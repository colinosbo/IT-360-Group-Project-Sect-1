from src.Clients.ms_graph import MicrosoftGraphClient

g = MicrosoftGraphClient()
g.get_token()  # interactive once, then cache

msg = g.get_latest_message()
if not msg:
    print("No messages found.")
else:
    print("\n📩 Newest email")
    print("Subject :", msg.get("subject"))
    print("From    :", msg.get("from", {}).get("emailAddress", {}).get("address"))
    print("Received:", msg.get("receivedDateTime"))
    print("Preview :", msg.get("bodyPreview"))
