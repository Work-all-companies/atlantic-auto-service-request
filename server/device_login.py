"""One-time: get a refresh token for AGC1130@outlook.com via device code.
Usage: MS_CLIENT_ID=... python device_login.py
Prints the code/URL, waits for sign-in, then prints the value to paste into the Render
secret MS_REFRESH_TOKEN (service aas-booking -> Environment). Nothing is written to disk."""
import os, json, time, urllib.request, urllib.parse, urllib.error
CID = os.environ["MS_CLIENT_ID"]; B = "https://login.microsoftonline.com/consumers/oauth2/v2.0/"
post = lambda u, d: json.load(urllib.request.urlopen(B + u, urllib.parse.urlencode(d).encode()))
r = post("devicecode", {"client_id": CID, "scope": "Calendars.ReadWrite offline_access"})
print(r["message"], flush=True)
while True:
    time.sleep(r.get("interval", 5))
    try: t = post("token", {"client_id": CID, "grant_type": "urn:ietf:params:oauth:grant-type:device_code", "device_code": r["device_code"]})
    except urllib.error.HTTPError as e:
        err = json.load(e).get("error")
        if err in ("authorization_pending", "slow_down"): continue
        raise SystemExit(err)
    print("\nPaste this into Render secret MS_REFRESH_TOKEN:\n" + t["refresh_token"], flush=True); break
