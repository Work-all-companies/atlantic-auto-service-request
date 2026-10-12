"""One-time: get a refresh token for AGC1130@outlook.com via device code.
Usage: MS_CLIENT_ID=... python device_login.py   -> prints code/URL, then polls and prints the refresh token."""
import os, json, time, urllib.request, urllib.parse
CID = os.environ["MS_CLIENT_ID"]; B = "https://login.microsoftonline.com/consumers/oauth2/v2.0/"
post = lambda u, d: json.load(urllib.request.urlopen(B + u, urllib.parse.urlencode(d).encode()))
r = post("devicecode", {"client_id": CID, "scope": "Calendars.ReadWrite offline_access"})
print(r["message"], flush=True)
while True:
    time.sleep(r.get("interval", 5))
    try: t = post("token", {"client_id": CID, "grant_type": "urn:ietf:params:oauth:grant-type:device_code", "device_code": r["device_code"]})
    except urllib.error.HTTPError as e:
        err = json.load(e).get("error")
        if err == "authorization_pending": continue
        raise SystemExit(err)
    open(os.environ.get("OUT", "refresh_token.txt"), "w").write(t["refresh_token"]); print("saved refresh token"); break
