import os, json, sys
import msal, requests
from dotenv import load_dotenv
load_dotenv()

TENANT_ID = os.getenv('TENANT_ID')
CLIENT_ID = os.getenv('CLIENT_ID')
RAW_SCOPES = (os.getenv("SCOPES") or "").split()
RESERVED = {"openid", "profile", "offline_access"}
SCOPES = [s for s in RAW_SCOPES if s not in RESERVED]

print("Scopes going to MSAL:", SCOPES)  # should NOT show reserved ones



#Loading Client and Tenant ID from .env file, then checking for errors
if not TENANT_ID or not CLIENT_ID or not SCOPES:
    raise RuntimeError('Please make sure TENANT_ID, CLIENT_ID, and SCOPES environment variables are set')

#Constants used by graph API
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH = "https://graph.microsoft.com/v1.0"

class MicrosoftGraphClient:

    def __init__(self):
        #Create a client
        self.app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
        self.token_result = None #holds auth token from microsoft / set to None as placeholder.

    def get_token(self):

        if self.token_result:
            return self.token_result

        #first attempt to get an access token from cache, if not have user sign in.
        accounts = self.app.get_accounts()
        if accounts:
            result = self.app.acquire_token_silent(SCOPES, account)
            if result and "access_token" in result:
                return result

        #Go back to interactive Device Code authentication (this prints a code)
        #the flow variable sends a request to microsoft login server, which returns
        flow = self.app.initiate_device_flow(scopes=SCOPES)
        if not flow:
            print("Failed to start device flow (no response).")
            sys.exit(1)

        # See everything Microsoft returned (user_code, verification_uri, etc.)
        #print("\nFlow payload:")
        #(json.dumps(flow, indent=2))

        #Promt user for authentication
        print(f"\nGo to {flow['verification_uri']} and enter code: {flow['user_code']}")
        res = self.app.acquire_token_by_device_flow(flow) #blocks until finish sign in
        if "access_token" in res:
            self.token_result = res["access_token"]
        else:
            print("Failed to get access token.")
            sys.exit(1)
        print("\nAuth success. Token acquired.")
        return self.token_result

    def headers(self):
        if not self.token_result:
            raise RuntimeError("No token yet — call get_token() first")
        return {
            "Authorization": f"Bearer {self.token_result}",
            "Accept": "application/json"
        }

if __name__ == "__main__":
    Client = MicrosoftGraphClient()
    token = Client.get_token()

    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(f"{GRAPH}/me?$select=displayName,mail,userPrincipalName", headers=headers, timeout=30)
    print("\n/me response status:", r.status_code)
    print("Body:", r.text)

    if r.ok:
        me = r.json()
        print(f"\n✅ Signed in as: {me.get('displayName')} <{me.get('mail') or me.get('userPrincipalName')}>")
    else:
        print("\ncall failed. Check API permissions and consent.")
