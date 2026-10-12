"""Atlantic Service & Repair booking server.
Outlook AAS SHOP is the record of bookings: availability is read live from Microsoft Graph on
every request. No local database or files (Render free plan wipes disk on restart/sleep).
Each slot holds 2 jobs (Brandon, Kolby), noon holds 1. A booking is check + create-Outlook-event
under one process lock, so two customers can never take the same last spot. Holds made in this
process are also kept in memory briefly in case Graph lags. Run with ONE worker.
Auth: MS_REFRESH_TOKEN env var (Render secret); rotated tokens are kept in memory only."""
import os, json, threading, datetime as dt, re, time, urllib.request, urllib.parse
from zoneinfo import ZoneInfo
from flask import Flask, request, jsonify

TZ = ZoneInfo("America/Chicago")
TECHS = ["Brandon", "Kolby"]                      # preference order
SLOT_HOURS = [8, 9, 10, 11, 12, 13, 14, 15, 16]
CAPACITY = {12: 1}                                # noon slot: one tech only; others hold len(TECHS)
cap = lambda h: CAPACITY.get(h, len(TECHS))
DAYS_AHEAD = 30
ALLOWED_ORIGIN = os.environ.get("ALLOWED_ORIGIN", "https://work-all-companies.github.io")
CLIENT_ID = os.environ.get("MS_CLIENT_ID", "")
CAL_ID = os.environ.get("AAS_CALENDAR_ID", "")
TOKEN_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
SCOPES = "Calendars.ReadWrite Mail.Send offline_access"
GRAPH = "https://graph.microsoft.com/v1.0"
NOTIFY_EMAILS = [e.strip() for e in os.environ.get("NOTIFY_EMAILS",
    "bkeegan@atlanticglassco.com,bschmidtke@atlanticautosalesco.com").split(",") if e.strip()]
FAKE_GRAPH = os.environ.get("FAKE_GRAPH")       # path to JSON list of events, for local tests

app = Flask(__name__)
LOCK = threading.Lock()

HOLDS = {}            # slot iso -> {tech: expiry_ts}; in-memory guard against Graph read lag
HOLD_SECS = 600
def holds_for(iso):
    now = time.time(); h = HOLDS.get(iso, {})
    for t in [t for t, x in h.items() if x < now]: del h[t]
    return set(h)

# ---------- Microsoft Graph ----------
_tok = {"access": None, "exp": 0}
def access_token():
    if _tok["access"] and time.time() < _tok["exp"] - 60: return _tok["access"]
    rt = _tok.get("refresh") or os.environ.get("MS_REFRESH_TOKEN", "")
    if not rt: raise RuntimeError("no refresh token")
    body = urllib.parse.urlencode({"client_id": CLIENT_ID, "grant_type": "refresh_token",
                                   "refresh_token": rt, "scope": SCOPES}).encode()
    r = json.load(urllib.request.urlopen(TOKEN_URL, body, timeout=20))
    _tok.update(access=r["access_token"], exp=time.time() + int(r["expires_in"]))
    if r.get("refresh_token"): _tok["refresh"] = r["refresh_token"]   # memory only; env token stays valid
    return _tok["access"]

def graph(method, path, body=None):
    req = urllib.request.Request(GRAPH + path, method=method, data=json.dumps(body).encode() if body else None,
        headers={"Authorization": "Bearer " + access_token(), "Content-Type": "application/json",
                 "Prefer": 'outlook.timezone="UTC"'})
    return json.load(urllib.request.urlopen(req, timeout=20))

def outlook_events(start, end):
    """Return [(start_utc, end_utc, subject)] for busy AAS SHOP events."""
    if FAKE_GRAPH:
        raw = json.load(open(FAKE_GRAPH))
    else:
        q = urllib.parse.urlencode({"startDateTime": start.astimezone(dt.timezone.utc).isoformat(),
            "endDateTime": end.astimezone(dt.timezone.utc).isoformat(), "$top": "500",
            "$select": "subject,start,end,showAs,isAllDay,isCancelled"})
        path = f"/me/calendars/{CAL_ID}/calendarView?{q}"; raw = []
        while path:
            r = graph("GET", path); raw += r["value"]
            path = r.get("@odata.nextLink", "").replace(GRAPH, "") or None
    out = []
    for e in raw:
        if e.get("isCancelled") or e.get("showAs") in ("free", "workingElsewhere"): continue
        if e.get("isAllDay"):
            s = dt.datetime.fromisoformat(e["start"]["dateTime"][:10]).replace(tzinfo=TZ)
            en = dt.datetime.fromisoformat(e["end"]["dateTime"][:10]).replace(tzinfo=TZ)
        else:
            s = dt.datetime.fromisoformat(e["start"]["dateTime"][:19]).replace(tzinfo=dt.timezone.utc)
            en = dt.datetime.fromisoformat(e["end"]["dateTime"][:19]).replace(tzinfo=dt.timezone.utc)
        out.append((s, en, e.get("subject") or ""))
    return out

def techs_in(subject):
    m = re.match(r"\s*(brandon|kolby)\b", subject, re.I)
    return {m.group(1).capitalize()} if m else set()   # no tech name at start -> ignored (slot stays open)

# ---------- availability ----------
def slot_list():
    today = dt.datetime.now(TZ).date(); out = []
    for i in range(1, DAYS_AHEAD + 1):
        d = today + dt.timedelta(days=i)
        if d.weekday() < 5:
            out += [dt.datetime.combine(d, dt.time(h), TZ) for h in SLOT_HOURS]
    return out

def busy_map(slots):
    """slot_start iso -> set of busy techs (live Outlook + in-memory holds)."""
    events = outlook_events(slots[0], slots[-1] + dt.timedelta(hours=1)) if slots else []
    res = {}
    for s in slots:
        e = s + dt.timedelta(hours=1); busy = holds_for(s.isoformat())
        for es, ee, subj in events:
            if es < e and ee > s: busy |= techs_in(subj)
        res[s.isoformat()] = busy
    return res

SENT_MAIL = []        # FAKE_GRAPH mode: messages recorded here instead of sent
def send_mail(subject, text):
    msg = {"message": {"subject": subject, "body": {"contentType": "Text", "content": text},
           "toRecipients": [{"emailAddress": {"address": a}} for a in NOTIFY_EMAILS]}, "saveToSentItems": True}
    if FAKE_GRAPH: SENT_MAIL.append(msg); return
    req = urllib.request.Request(GRAPH + "/me/sendMail", method="POST", data=json.dumps(msg).encode(),
        headers={"Authorization": "Bearer " + access_token(), "Content-Type": "application/json"})
    urllib.request.urlopen(req, timeout=20).read()   # 202, empty body

def notify(d, start, tech):
    slot = start.strftime("%a %b %d, %Y %I:%M %p").replace(" 0", " ") + " CT"
    lines = [f"Name: {d.get('name','')}", f"Phone: {d.get('phone','')}", f"Email: {d.get('email','') or '-'}",
             f"Vehicle: {d.get('vehicle','')}", f"VIN or plate: {d.get('vin_or_plate','')}",
             f"Issue: {d.get('issue','')}", f"Booked: {slot}", f"Tech: {tech}"]
    try: send_mail(f"Service request: {d['name']} - {d['vehicle']} - {slot}", "\n".join(lines))
    except Exception as e: app.logger.error("notify email failed: %s", e)

VIN_RE = re.compile(r"^[A-HJ-NPR-Z0-9]{17}$")
PLATE_RE = re.compile(r"^[A-Z0-9]{2,8}$")
def max_year(): return dt.datetime.now(TZ).year + 1

def validate(d):
    """Return (clean_dict, None) or (None, error message). Mirrors the checks in index.html."""
    g = lambda k: str(d.get(k) or "").strip()
    c = {k: g(k) for k in ("name", "email", "make", "model", "issue")}
    for k, lbl in (("name", "Name"), ("make", "Make"), ("model", "Model"), ("issue", "Description of the issue")):
        if not c[k]: return None, f"{lbl} is required."
    digits = re.sub(r"\D", "", g("phone"))
    if len(digits) == 11 and digits[0] == "1": digits = digits[1:]
    if len(digits) != 10: return None, "Phone number must be exactly 10 digits."
    c["phone"] = f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"
    y = g("year")
    if not re.fullmatch(r"\d{4}", y) or not 1950 <= int(y) <= max_year():
        return None, f"Vehicle year must be between 1950 and {max_year()}."
    c["year"] = y
    vin = re.sub(r"[\s-]", "", g("vin")).upper(); plate = re.sub(r"[\s-]", "", g("plate")).upper()
    if vin:
        if len(vin) != 17: return None, "VIN must be exactly 17 characters."
        if not VIN_RE.match(vin): return None, "VIN can only contain letters and numbers (no I, O or Q)."
        plate = plate if PLATE_RE.match(plate) else ""
    elif plate:
        if not PLATE_RE.match(plate): return None, "License plate must be 2-8 letters or numbers."
    else: return None, "Please enter a VIN or a license plate number."
    c["vin"], c["plate"] = vin, plate
    c["vehicle"] = f"{y} {c['make']} {c['model']}"
    c["vin_or_plate"] = (f"VIN {vin}" if vin else "") + (" / " if vin and plate else "") + (f"Plate {plate}" if plate else "")
    if c["email"] and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", c["email"]): return None, "Email address doesn't look right."
    if not g("slot_start"): return None, "Please choose a time."
    c["slot_start"] = g("slot_start")
    return c, None

@app.after_request
def cors(r):
    r.headers["Access-Control-Allow-Origin"] = ALLOWED_ORIGIN
    r.headers["Access-Control-Allow-Headers"] = "Content-Type"
    r.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    r.headers["Cache-Control"] = "no-store"
    return r

@app.get("/health")
def health(): return {"ok": True}

@app.get("/availability")
def availability():
    slots = slot_list(); bm = busy_map(slots)
    # customers only see open / taken
    return jsonify({"slots": [{"start": k, "taken": len(v) >= cap(dt.datetime.fromisoformat(k).hour)} for k, v in bm.items()]})

@app.route("/book", methods=["POST", "OPTIONS"])
def book():
    if request.method == "OPTIONS": return ("", 204)
    d = request.get_json(force=True, silent=True) or {}
    if d.get("_gotcha"): return {"ok": True}
    d, err = validate(d)
    if err: return {"ok": False, "error": err}, 400
    try: start = dt.datetime.fromisoformat(d["slot_start"]).astimezone(TZ)
    except Exception: return {"ok": False, "error": "Please choose a valid time."}, 400
    if start not in slot_list(): return {"ok": False, "error": "That time is no longer offered. Please choose another."}, 400
    with LOCK:
        busy = busy_map([start])[start.isoformat()]
        free = [t for t in TECHS if t not in busy]
        if not free or len(busy) >= cap(start.hour):
            return {"ok": False, "error": "taken"}, 409
        tech = free[0]
        if not FAKE_GRAPH:   # create the Outlook event while still holding the lock
            body = "\n".join(f"{k}: {d.get(k,'')}" for k in ("name","phone","email","vehicle","vin_or_plate","issue"))
            try:
                graph("POST", f"/me/calendars/{CAL_ID}/events", {
                    "subject": f"{tech} - Customer: {d['name']} - {d['vehicle']}",
                    "body": {"contentType": "text", "content": body + "\n\nBooked online. Call customer to confirm."},
                    "start": {"dateTime": start.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "Central Standard Time"},
                    "end": {"dateTime": (start + dt.timedelta(hours=1)).strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "Central Standard Time"},
                    "showAs": "busy"})
            except Exception as e:
                app.logger.error("outlook event failed: %s", e)
                return {"ok": False, "error": "calendar unavailable, please call 712-254-6914"}, 503
        HOLDS.setdefault(start.isoformat(), {})[tech] = time.time() + HOLD_SECS
    notify(d, start, tech)   # outside the lock; failure never fails the booking
    return {"ok": True, "slot_start": start.isoformat()}

if __name__ == "__main__":
    app.run(port=int(os.environ.get("PORT", 8000)), threaded=True)
