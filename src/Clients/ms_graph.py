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
            result = self.app.acquire_token_silent(SCOPES, account[0])
            if result and "access_token" in result:
                return self.token_result

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

    def get_email_headers(self, message_id: str):
        headers = self.headers()  #  get auth headers for request
        url = f"{GRAPH}/me/messages/{message_id}?$select=internetMessageHeaders"

        r = requests.get(url, headers=headers, timeout=30)
        if not r.ok:
            print(f"❌ Error fetching headers: {r.status_code} {r.text}")
            return []

        data = r.json()
        headers_list = data.get("internetMessageHeaders", [])  #  lowercase key
        print(f"Retrieved {len(headers_list)} headers.")

        important = ['subject', 'x-sender-ip', 'authentication-results',
                     'received-spf', 'dkim-signature']
        for h in headers_list:
            name = h.get('name', '')
            val = h.get('value', '')
            if name.lower() in important:
                print(f"{name}: {val}")
        return headers_list

    @staticmethod
    def parse_spf_dkim(headers: str):
        spf_pattern = r"spf\s*=\s*(pass|fail|softfail|neutral|none)"
        dkim_pattern = r"dkim\s*=\s*(pass|fail|softfail|neutral|none)"
        spf_result, dkim_result = "unknown", "unknown"

        for h in headers:
            print(f"{h.get('name', '')}: {h.get('value', '')}")
            if h['name'].lower() == 'authentication-results':
                auth_value = h['value']
                spf_match = re.search(spf_pattern, auth_value, re.IGNORECASE)
                dkim_match = re.search(dkim_pattern, auth_value, re.IGNORECASE)
                if spf_match:
                    spf_result = spf_match.group(1).strip()
                if dkim_match:
                    dkim_result = dkim_match.group(1).strip()
        return {"spf": spf_result, "dkim": dkim_result}

    def get_latest_message(self):
        headers = self.headers()  # uses the bearer token you already stored
        url = f"{GRAPH}/me/messages"
        params = {
            "$top": 1,
            "$orderby": "receivedDateTime DESC",
            "$select": "id,subject,from,receivedDateTime,bodyPreview"
        }
        r = requests.get(url, headers=headers, params=params, timeout=30)
        if not r.ok:
            # show why it failed
            print("GET /me/messages failed:", r.status_code, repr(r.text))
            return None
        data = r.json()
        return (data.get("value") or [None])[0]