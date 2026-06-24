#!/usr/bin/env python3
"""Re-authorize Google OAuth and save new token.json"""

import json
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
creds = flow.run_local_server(port=0)

token_data = {
    "token": creds.token,
    "refresh_token": creds.refresh_token,
    "token_uri": creds.token_uri,
    "client_id": creds.client_id,
    "client_secret": creds.client_secret,
    "scopes": list(creds.scopes),
}

with open("token.json", "w") as f:
    json.dump(token_data, f, indent=2)

print("token.json saved!")
print()
print("Copy this JSON to Streamlit Secrets as GOOGLE_TOKEN:")
print(json.dumps(token_data))
