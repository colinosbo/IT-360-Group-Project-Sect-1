import requests
from src.Clients.ms_graph import MicrosoftGraphClient
import json
from bs4 import BeautifulSoup

#it360project
#it360password
# connect + authentication
client = MicrosoftGraphClient()
client.get_token()

# currently grabs the last email
email = client.get_latest_message()

id = email.get("id") #Retreives email ID
headers = client.get_email_headers(id) #Uses ID to grab email headers
#print(headers)
#print(id)

MicrosoftGraphClient.printHeaderList(headers) #printing all headers for testing

spf_dkim_result = client.parse_spf_dkim(headers) #Setting SPF and Dkim variables
spf = spf_dkim_result["spf"]
dkim = spf_dkim_result["dkim"]
print(f"SPF: {spf}") #Printing SPF and DKIM values
print(f"DKIM: {dkim}")

if not email:
    print("No messages found. Try g.debug_list_folders() or g.get_latest_received_anywhere().")
else:
    plain_body = client.htmlToPlainText(email["body"]["content"])

    print("\nNewest email")
    print("Subject :", email.get("subject"))
    print("From    :", email.get("from", {}).get("emailAddress", {}).get("address"))
    print("Received:", email.get("receivedDateTime"))
    print("Preview :", email.get("bodyPreview"))
    print(f"body:", plain_body)
    print(f"Message ID:", id)
    print("hasAttachments:", email.get("hasAttachments"))

