Enter your credentials in the env and run the code.
It will cache the auth token once you login to your account and will create a file "token_cache.json".
This way everytime the code/service restarts you dont need to login again untill the access token expires.
The code can be run as a service in the background it acts like n8n's outlook trigger.
