import os, json, sys
import msal, requests
from msal import SerializableTokenCache
from bs4 import BeautifulSoup
from dotenv import load_dotenv
import html
import re
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
        # path to persistent cache file
        self.cache_file = "msal_cache.json"
        self.cache = SerializableTokenCache()

        # load previous tokens if file exists
        if os.path.exists(self.cache_file):
            with open(self.cache_file, "r") as f:
                self.cache.deserialize(f.read())

        # create the app and attach cache
        self.app = msal.PublicClientApplication(
            CLIENT_ID,
            authority=AUTHORITY,
            token_cache=self.cache
        )

        self.token_result = None

    def get_token(self):

        accounts = self.app.get_accounts()
        if accounts:
            print("Existiing account found")
            result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self.token_result = result["access_token"]
                print(" Token loaded from cache.")
                return self.token_result
        print("🧾 No cached token found, starting device code flow...")
        flow = self.app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError("Failed to start device flow.")
        print(f"\nGo to {flow['verification_uri']} and enter code: {flow['user_code']}")
        result = self.app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            raise RuntimeError("Failed to acquire token from Microsoft Graph.")
        self.token_result = result["access_token"]
        print("Token acquired interactively.")

        # Save refreshed tokens to disk
        if self.cache.has_state_changed:
            with open(self.cache_file, "w") as f:
                f.write(self.cache.serialize())

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
            #if name.lower() in important:
                #print(f"{name}: {val}")
        return headers_list
    @staticmethod
    def printHeaderList(headers: []):
        if not headers:
            print("⚠️ No headers found")
            return
        for h in headers:
            name = h.get("name", "")
            value = h.get("value", "")
            print(f"{name}: {value}")
    @staticmethod
    def parse_spf_dkim(headers: str):
        spf_pattern = r"spf\s*=\s*(pass|fail|softfail|neutral|none)"
        dkim_pattern = r"dkim\s*=\s*(pass|fail|softfail|neutral|none)"
        spf_result, dkim_result = "unknown", "unknown"

        for h in headers:
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
        url = f"{GRAPH}/me/mailFolders/inbox/messages"
        params = {
            "$top": "1",
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,from,receivedDateTime,bodyPreview, replyTo, sender, hasAttachments, body"
        }
        r = requests.get(url, headers=headers, params=params, timeout=30)
        if not r.ok:
            # show why it failed
            print("GET /me/messages failed:", r.status_code, repr(r.text))
            return None
        data = r.json()
        return (data.get("value") or [None])[0]

    def debug_list_folders(self):
        url = f"{GRAPH}/me/mailFolders"
        params = {"$select": "displayName,id,totalItemCount,unreadItemCount"}
        r = requests.get(url, headers=self.headers(), params=params, timeout=30)
        r.raise_for_status()
        for f in r.json().get("value", []):
            print(
                f"{f['displayName']:<25}  items={f['totalItemCount']:<5}  unread={f['unreadItemCount']:<5}  id={f['id']}")

    def debug_who_am_i(self):
        r = requests.get(
            "https://graph.microsoft.com/v1.0/me?$select=id,displayName,mail,userPrincipalName",
            headers=self.headers(), timeout=30)
        r.raise_for_status()
        me = r.json()
        print("SIGNED IN AS:")
        print(" displayName       :", me.get("displayName"))
        print(" mail              :", me.get("mail"))
        print(" userPrincipalName :", me.get("userPrincipalName"))

        # Also dump tenant + preferred username from the token to catch MSA vs AAD
        import jwt  # pip install pyjwt if needed, but we can also parse manually
        token = self.headers().get("Authorization","").split()[-1]
        # decode header only (no verify) to get tid/upn; fallback if lib not installed
        try:
            claims = jwt.get_unverified_claims(token)
            print(" token.tid (tenant):", claims.get("tid"))
            print(" token.preferred_username:", claims.get("preferred_username"))
            print(" token.iss           :", claims.get("iss"))
        except Exception:
            pass
    @staticmethod
    def htmlToPlainText(body: str):
        if not isinstance(body, str):
            raise TypeError(f"body must be a string, {type(body)}")
        soup = BeautifulSoup(body, "html.parser")
        return soup.get_text(separator=' ', strip=True)