import dropbox
from dropbox import DropboxOAuth2FlowNoRedirect

APP_KEY = input("Enter your Dropbox APP_KEY: ")
APP_SECRET = input("Enter your Dropbox APP_SECRET: ")

auth_flow = DropboxOAuth2FlowNoRedirect(
    APP_KEY,
    APP_SECRET,
    token_access_type='offline'
)

authorize_url = auth_flow.start()
print(f"\n1. Go to: {authorize_url}")
print("2. Click 'Allow'")
print("3. Copy the authorization code\n")

auth_code = input("Enter the authorization code: ").strip()
oauth_result = auth_flow.finish(auth_code)

print(f"\n✅ DROPBOX_REFRESH_TOKEN = {oauth_result.refresh_token}")
print(f"✅ DROPBOX_ACCESS_TOKEN  = {oauth_result.access_token}")
