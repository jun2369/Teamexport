import os

# Fill these in after creating the App Registration in Microsoft Entra admin center:
#   Entra admin center -> App registrations -> New registration
#   Name: Teams Message Exporter
#   Supported account types: Accounts in this organizational directory only
#   Authentication -> Add a platform -> Web -> redirect URI: <REDIRECT_URI below>
#   Certificates & secrets -> New client secret -> set MS_CLIENT_SECRET env var to its value
#   API permissions -> Microsoft Graph -> Delegated -> add: User.Read, Chat.Read, Files.Read
CLIENT_ID = "58511677-b8c8-4d6e-bf67-fa86799d13d0"
TENANT_ID = "a7f846ec-8219-438c-8896-9e45cdbbe994"

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Chat.Read", "Files.Read"]

# Client secret for the app registration and the key Flask uses to sign
# session cookies. Read from the environment rather than hardcoded here,
# since this project folder syncs to OneDrive.
CLIENT_SECRET = os.environ["MS_CLIENT_SECRET"]
FLASK_SECRET_KEY = os.environ["FLASK_SECRET_KEY"]

# Must exactly match a redirect URI registered on the app's "Web" platform
# in Entra. Override via REDIRECT_URI once this is hosted somewhere other
# than localhost.
REDIRECT_PATH = "/auth/callback"
REDIRECT_URI = os.environ.get("REDIRECT_URI", f"http://localhost:5000{REDIRECT_PATH}")

# The Teams group chat topic to look for by default. You can also pick
# a different chat from the dropdown in the web UI.
DEFAULT_CHAT_TOPIC = "AGS Nimbus Support"

# Per-user MSAL token caches are stored as one file per browser session
# under this directory (see auth.py).
SESSIONS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sessions")
