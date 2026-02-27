import os
import base64
import time
import re
from datetime import datetime, timezone
from pathlib import Path
import json
import msal
import requests
from dotenv import load_dotenv

load_dotenv()


# ── Config ────────────────────────────────────────────────────────────────────

class Config:
    CLIENT_ID           = os.getenv("CLIENT_ID")
    USER_EMAIL          = os.getenv("USER_EMAIL")
    POLL_INTERVAL       = int(os.getenv("POLL_INTERVAL", 15))
    SAVE_ATTACHMENTS_TO = os.getenv("SAVE_ATTACHMENTS_TO", "./attachments")
    TOKEN_CACHE_FILE    = os.getenv("TOKEN_CACHE_FILE", ".token_cache.json")

    #they are what makes auth work.
    SCOPES    = ["Mail.Read"]
    AUTHORITY = "https://login.microsoftonline.com/common"

    GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


# ── Persistent token cache ────────────────────────────────────────────────────

class PersistentTokenCache(msal.SerializableTokenCache):
    def __init__(self, cache_file: str):
        super().__init__()
        self._cache_file = Path(cache_file)
        if self._cache_file.exists():
            self.deserialize(self._cache_file.read_text())

    def _save(self) -> None:
        if self.has_state_changed:
            self._cache_file.write_text(self.serialize())

    def add(self, event, **kwargs):
        super().add(event, **kwargs)
        self._save()

    def update_rt(self, rt_item, new_rt):
        super().update_rt(rt_item, new_rt)
        self._save()

    def remove_rt(self, rt_item):
        super().remove_rt(rt_item)
        self._save()

    def remove_at(self, at_item):
        super().remove_at(at_item)
        self._save()


# ── Auth ──────────────────────────────────────────────────────────────────────

class AuthManager:
    def __init__(self, config: Config):
        self._scopes     = config.SCOPES
        self._cache_file = config.TOKEN_CACHE_FILE
        self._cache      = PersistentTokenCache(config.TOKEN_CACHE_FILE)
        print("got the cache token")
        self._app        = msal.PublicClientApplication(
            config.CLIENT_ID,
            authority=config.AUTHORITY,
            token_cache=self._cache,
        )

    def get_token(self) -> str:
        accounts = self._app.get_accounts()
        if accounts:
            result = self._app.acquire_token_silent(self._scopes, account=accounts[0])
            if result and "access_token" in result:
                return result["access_token"]
            print("[AUTH] Cached token expired — re-authenticating.")
            self._clear_cache()

        print("[AUTH] No valid token found — one-time device code login required.")
        return self._device_flow_login()

    def _clear_cache(self) -> None:
        cache_path = Path(self._cache_file)
        if cache_path.exists():
            cache_path.unlink()
            print("[AUTH] Stale token cache cleared.")
        self._cache = PersistentTokenCache(self._cache_file)
        self._app._token_cache = self._cache

    def _device_flow_login(self) -> str:
        flow = self._app.initiate_device_flow(scopes=self._scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"Device flow failed: {flow}")

        print(flow["message"])
        result = self._app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            raise RuntimeError(f"Auth failed: {result.get('error_description')}")

        print("[AUTH] Login successful — token cached for future runs.\n")
        return result["access_token"]


# ── Graph API client ──────────────────────────────────────────────────────────

class GraphClient:
    def __init__(self, base_url: str):
        self._base_url = base_url
        self._token    = ""

    def set_token(self, token: str) -> None:
        self._token = token

    def _headers(self) -> dict:
        return {"Authorization": f"Bearer {self._token}"}

    def _get(self, endpoint: str, params: dict = None) -> dict:
        url      = f"{self._base_url}{endpoint}"
        response = requests.get(url, headers=self._headers(), params=params)
        response.raise_for_status()
        return response.json()

    def fetch_inbox_messages(self, since: datetime, top: int = 20) -> list[dict]:
        since_str = since.strftime("%Y-%m-%dT%H:%M:%SZ")
        params = {
            "$filter": f"receivedDateTime ge {since_str}",
            "$top":    top,
            "$select": "id,subject,from,receivedDateTime,body,hasAttachments,conversationId",
        }
        messages = self._get("/me/mailFolders/inbox/messages", params=params).get("value", [])
        return messages

    def fetch_attachments(self, email_id: str) -> list[dict]:
        return self._get(f"/me/messages/{email_id}/attachments").get("value", [])

    def fetch_latest_messages(self, top: int = 5) -> list[dict]:
        params = {"$top": top, "$select": "subject,receivedDateTime,from"}
        return self._get("/me/mailFolders/inbox/messages", params=params).get("value", [])

    def fetch_conversation_thread(self, conversation_id: str) -> list[dict]:
        params = {
            "$filter": f"conversationId eq '{conversation_id}'",
            "$select": "id,subject,from,receivedDateTime,body,hasAttachments",
        }
        messages = self._get("/me/mailFolders/inbox/messages", params=params).get("value", [])
        return sorted(messages, key=lambda m: m["receivedDateTime"])


# ── Attachments ───────────────────────────────────────────────────────────────

class AttachmentManager:
    def __init__(self, save_dir: str):
        self._save_dir = save_dir
        os.makedirs(self._save_dir, exist_ok=True)

    def save(self, attachment: dict) -> str | None:
        filename = attachment.get("name", "unnamed_file")
        content  = attachment.get("contentBytes", "")
        if not content:
            return None
        filepath = os.path.join(self._save_dir, filename)
        with open(filepath, "wb") as f:
            f.write(base64.b64decode(content))
        return filepath


# ── Printer ───────────────────────────────────────────────────────────────────
#use this class if you need to print the email in a readable format and not a json body.
# class EmailPrinter:
#     def print_email(self, email: dict, attachments: list[dict], saved_paths: list) -> None:
#         print("\n" + "━" * 60)
#         print("NEW EMAIL")
#         print("━" * 60)
#         sender = email["from"]["emailAddress"]
#         print(f"  From    : {sender['name']} <{sender['address']}>")
#         print(f"  Subject : {email['subject']}")
#         print(f"  Time    : {email['receivedDateTime']}")
#         print(f"  Preview : {email['bodyPreview'][:300]}")
#         if attachments:
#             print("\nAttachments:")
#             for att, path in zip(attachments, saved_paths):
#                 size_kb = round(att.get("size", 0) / 1024, 1)
#                 label   = f"saved to {path}" if path else "no content to save"
#                 print(f"    {att['name']} ({size_kb} KB) → {label}")
#         else:
#             print("\n  No attachments.")
#         print("━" * 60)



#use this class if you need to print the email in a json body.
class EmailPrinter:
    def _strip_html(self, html: str) -> str:
        return re.sub(r'<[^>]+>', '', html).strip()

    def print_email(self, email: dict, attachments: list[dict], saved_paths: list, thread: list = []) -> None:
        body = email.get("body", {})
        if body.get("contentType") == "html":
            email["body"]["content"] = self._strip_html(body["content"])

        for msg in thread:
            msg_body = msg.get("body", {})
            if msg_body.get("contentType") == "html":
                msg["body"]["content"] = self._strip_html(msg_body["content"])

        output = {
            "email": email,
            "thread": thread,
            "attachments": [
                {
                    "file_name": att.get("name"),
                    "saved_path": path,
                    "size_kb": round(att.get("size", 0) / 1024, 1)
                }
                for att, path in zip(attachments, saved_paths)
            ]
        }

        print("\n" + "━" * 60)
        print("New Email(JSON)")
        print("━" * 60)
        print(json.dumps(output, indent=4, ensure_ascii=False))
        print("━" * 60)

# ── Monitor ───────────────────────────────────────────────────────────────────

class EmailMonitor:
    def __init__(self, config: Config):
        self._config      = config
        self._auth        = AuthManager(config)
        self._graph       = GraphClient(config.GRAPH_BASE_URL)
        self._attachments = AttachmentManager(config.SAVE_ATTACHMENTS_TO)
        self._printer     = EmailPrinter()
        self._seen_ids: set[str] = set()

    def _refresh_token(self) -> None:
        token = self._auth.get_token()
        self._graph.set_token(token)

    def _process_email(self, email: dict) -> None:
        attachments, saved_paths = [], []
        if email.get("hasAttachments"):
            attachments = self._graph.fetch_attachments(email["id"])
            saved_paths = [self._attachments.save(a) for a in attachments]

        conversation_id = email.get("conversationId")
        thread = []
        if conversation_id:
            thread = self._graph.fetch_conversation_thread(conversation_id)

        self._printer.print_email(email, attachments, saved_paths, thread)

    def _poll(self, last_checked: datetime) -> datetime:
        self._refresh_token()
        emails = self._graph.fetch_inbox_messages(since=last_checked)
        now    = datetime.now(timezone.utc)
        for email in emails:
            if email["id"] not in self._seen_ids:
                self._seen_ids.add(email["id"])
                self._process_email(email)
        return now

    def _verify_mail_access(self) -> bool:
        try:
            messages = self._graph.fetch_latest_messages()
            print(f"Mail access confirmed — {len(messages)} message(s) visible in inbox.\n")
            print("pls send a new email to this mailbox to trigger the workflow")
            return True
        except requests.exceptions.HTTPError as e:
            print(f"Mail access check failed: {e}")
            return False

    def start(self) -> None:
        print(f"Watching inbox for: {self._config.USER_EMAIL}")
        print(f"Poll interval: {self._config.POLL_INTERVAL}s  (Ctrl+C to stop)\n")

        self._refresh_token()

        if not self._verify_mail_access():
            print("Cannot access mailbox. Exiting.")
            return

        last_checked = datetime.now(timezone.utc)

        while True:
            try:
                last_checked = self._poll(last_checked)
            except requests.exceptions.HTTPError as e:
                if e.response is not None and e.response.status_code == 401:
                    print("401 Unauthorized re-authenticating...pls wait")
                    self._auth._clear_cache()
                else:
                    print(f"HTTP error: {e}")
            except requests.exceptions.RequestException as e:
                print(f"Network: {e}")
            except Exception as e:
                print(f"error: {e}")
            time.sleep(self._config.POLL_INTERVAL)


# ── Entry point ───────────────────────────────────────────────────────────────

def main() -> None:
    config  = Config()
    monitor = EmailMonitor(config)
    monitor.start()


if __name__ == "__main__":
    main()
