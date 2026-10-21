"""Atlantic Service & Repair booking server.
Outlook AAS SHOP is the record of bookings: availability is read live from Microsoft Graph on
every request. No local database or files (Render free plan wipes disk on restart/sleep).
Each slot holds 2 jobs (Brandon, Kolby), noon holds 1. A booking is check + create-Outlook-event
under one process lock, so two customers can never take the same last spot. Holds made in this
process are also kept in memory briefly in case Graph lags. Run with ONE worker.
Auth: MS_REFRESH_TOKEN env var (Render secret); rotated tokens are kept in memory only."""
import os, json, threading, hashlib, hmac, atexit, datetime as dt, re, time, urllib.request, urllib.parse
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
            "$select": "subject,start,end,showAs,isAllDay,isCancelled,originalStartTimeZone"})
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
            fixed = FIXED_CENTRAL.get(e.get("originalStartTimeZone") or "")
            if fixed is not None:
                # Typed in Outlook set to a no-DST "Central" zone (e.g. Central America, UTC-6):
                # the shop means the wall-clock time they typed, in Chicago time.
                s = (s + fixed).replace(tzinfo=None).replace(tzinfo=TZ)
                en = (en + fixed).replace(tzinfo=None).replace(tzinfo=TZ)
        out.append((s, en, e.get("subject") or ""))
    return out

FIXED_CENTRAL = {"Central America Standard Time": dt.timedelta(hours=-6),
                 "Canada Central Standard Time": dt.timedelta(hours=-6)}

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
    """Return (clean_dict, None) or (None, error message). Never raises."""
    try: return _validate(d if isinstance(d, dict) else {})
    except Exception as e:
        app.logger.error("validate failed: %s", e)
        return None, "Please check the form and try again."

def _legacy(d):
    """Old cached form: single 'vehicle' text + 'vin_or_plate'. Map onto new fields."""
    d = dict(d)
    if not any(str(d.get(k) or "").strip() for k in ("year", "make", "model")) and str(d.get("vehicle") or "").strip():
        v = str(d["vehicle"]).strip()
        m = re.match(r"^\s*'?(\d{2,4})\b\s*(.*)$", v)
        if m:
            y = m.group(1); y = ("20" + y if int(y) <= max_year() % 100 else "19" + y) if len(y) == 2 else y
            rest = m.group(2).split(None, 1)
            d["year"] = y; d["make"] = rest[0] if rest else ""; d["model"] = rest[1] if len(rest) > 1 else (rest[0] if rest else "")
        else:
            d["year"] = ""; d["make"] = v; d["model"] = v
    if not str(d.get("vin") or "").strip() and not str(d.get("plate") or "").strip() and str(d.get("vin_or_plate") or "").strip():
        x = re.sub(r"^(vin|plate)\s*", "", str(d["vin_or_plate"]).strip(), flags=re.I)
        x2 = re.sub(r"[\s-]", "", x)
        if len(x2) == 17: d["vin"] = x2
        else: d["plate"] = x2
    return d

def _validate(d):
    d = _legacy(d)
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


# ---------- private visit stats ----------
# Counts live in a separate Outlook calendar "AAS Stats" (not AAS SHOP) as a JSON extended property on
# one placeholder event dated 2000-01-01, so they survive Render restarts with no disk/GitHub access.
# Pings are batched in memory and flushed at most every STATS_FLUSH_SECS (and at shutdown).
STATS_KEY_SHA256 = os.environ.get("STATS_KEY_SHA256", "dc9f64169a06b5727748886111f58b0af51be3944c9216e7cdb9480d2409aa68")
STATS_CAL_NAME = "AAS Stats"; STATS_SUBJECT = "AAS-STATS-DATA (do not delete)"
STATS_PROP = "String {6f1c3a52-8f0e-4b8e-9f43-2a7d5c1e9b10} Name aasStats"
STATS_FLUSH_SECS = int(os.environ.get("STATS_FLUSH_SECS", "60"))
STATS_FAKE = os.environ.get("STATS_FAKE")           # path to local JSON file, for tests
SLOCK = threading.Lock(); PENDING = {}               # day -> {field: n}
_sref = {}
COUNT_FIELDS = ("qr", "link", "unique_qr", "unique_link", "new_visitors", "bookings")

def stat_bump(field, n=1):
    if field not in COUNT_FIELDS: return
    day = dt.datetime.now(TZ).date().isoformat()
    with SLOCK:
        d = PENDING.setdefault(day, {}); d[field] = d.get(field, 0) + n

def _stats_event():
    if "ev" in _sref: return _sref["cal"], _sref["ev"]
    q = urllib.parse.quote(f"name eq '{STATS_CAL_NAME}'")
    cals = graph("GET", f"/me/calendars?$filter={q}")["value"]
    cal = cals[0]["id"] if cals else graph("POST", "/me/calendars", {"name": STATS_CAL_NAME})["id"]
    pid = urllib.parse.quote(f"id eq '{STATS_PROP}'")
    sq = urllib.parse.quote(f"subject eq '{STATS_SUBJECT}'")
    evs = graph("GET", f"/me/calendars/{cal}/events?$filter={sq}&$select=id&$expand=singleValueExtendedProperties($filter={pid})")["value"]
    if evs: ev = evs[0]["id"]
    else:
        ev = graph("POST", f"/me/calendars/{cal}/events", {"subject": STATS_SUBJECT, "showAs": "free",
            "isReminderOn": False, "body": {"contentType": "text", "content": "Form visit counts. Do not delete."},
            "start": {"dateTime": "2000-01-01T00:00:00", "timeZone": "UTC"}, "end": {"dateTime": "2000-01-01T00:30:00", "timeZone": "UTC"},
            "singleValueExtendedProperties": [{"id": STATS_PROP, "value": "{}"}]})["id"]
    _sref.update(cal=cal, ev=ev); return cal, ev

def stats_load():
    if STATS_FAKE:
        try: return json.load(open(STATS_FAKE))
        except Exception: return {}
    cal, ev = _stats_event()
    pid = urllib.parse.quote(f"id eq '{STATS_PROP}'")
    r = graph("GET", f"/me/calendars/{cal}/events/{ev}?$select=id&$expand=singleValueExtendedProperties($filter={pid})")
    p = r.get("singleValueExtendedProperties") or []
    return json.loads(p[0]["value"]) if p else {}

def stats_save(data):
    if STATS_FAKE: json.dump(data, open(STATS_FAKE, "w")); return
    cal, ev = _stats_event()
    graph("PATCH", f"/me/calendars/{cal}/events/{ev}", {"singleValueExtendedProperties": [{"id": STATS_PROP, "value": json.dumps(data, separators=(",", ":"))}]})

def _merge(base, add):
    out = {k: dict(v) for k, v in base.items()}
    for day, f in add.items():
        d = out.setdefault(day, {})
        for k, n in f.items(): d[k] = d.get(k, 0) + n
    return out

FLUSH_LOCK = threading.Lock()
def stats_flush():
    with FLUSH_LOCK:
        with SLOCK:
            if not PENDING: return True
            batch = {k: dict(v) for k, v in PENDING.items()}; PENDING.clear()
        try: stats_save(_merge(stats_load(), batch)); return True
        except Exception as e:
            app.logger.error("stats flush failed: %s", e)
            with SLOCK:
                for day, f in batch.items():
                    d = PENDING.setdefault(day, {})
                    for k, n in f.items(): d[k] = d.get(k, 0) + n
            return False

def _flusher():
    while True:
        time.sleep(STATS_FLUSH_SECS); stats_flush()
if not os.environ.get("STATS_NO_THREAD"):
    threading.Thread(target=_flusher, daemon=True).start()
atexit.register(stats_flush)

def stats_summary(data):
    days = sorted(data); tot = {k: 0 for k in COUNT_FIELDS}
    daily = []
    for d in days:
        row = {k: data[d].get(k, 0) for k in COUNT_FIELDS}
        row["visits"] = row["qr"] + row["link"]; row["unique"] = row["unique_qr"] + row["unique_link"]
        for k in COUNT_FIELDS: tot[k] += row[k]
        daily.append(dict(date=d, **row))
    tot["visits"] = tot["qr"] + tot["link"]; tot["unique"] = tot["unique_qr"] + tot["unique_link"]
    today = dt.datetime.now(TZ).date(); wk = (today - dt.timedelta(days=6)).isoformat()
    last7 = {k: sum(r[k] for r in daily if r["date"] >= wk) for k in list(COUNT_FIELDS) + ["visits", "unique"]}
    return {"timezone": "America/Chicago", "totals": tot, "last_7_days": last7, "daily": daily,
            "note": "unique = first visit that day per browser; new_visitors = first visit ever per browser"}

@app.after_request
def cors(r):
    r.headers["Access-Control-Allow-Origin"] = ALLOWED_ORIGIN
    r.headers["Access-Control-Allow-Headers"] = "Content-Type"
    r.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    r.headers["Cache-Control"] = "no-store"
    return r

@app.errorhandler(Exception)
def on_error(e):
    from werkzeug.exceptions import HTTPException
    if isinstance(e, HTTPException):
        r = jsonify({"ok": False, "error": e.description or e.name}); r.status_code = e.code or 500
    else:
        app.logger.exception("unhandled error")
        r = jsonify({"ok": False, "error": "Something went wrong on our end. Please call the shop at 712-254-6914."}); r.status_code = 500
    return cors(r)

@app.get("/health")
def health(): return {"ok": True}


@app.route("/visit", methods=["POST", "GET", "OPTIONS"])
def visit():
    if request.method == "OPTIONS": return ("", 204)
    d = request.get_json(force=True, silent=True) or {}
    if not isinstance(d, dict): d = {}
    src = "qr" if str(d.get("src") or request.args.get("src") or "").lower() == "qr" else "link"
    stat_bump(src)
    if d.get("first_today"): stat_bump("unique_" + src)
    if d.get("new"): stat_bump("new_visitors")
    return ("", 204)

@app.get("/stats")
def stats():
    k = request.args.get("key") or request.headers.get("X-Stats-Key") or ""
    if not hmac.compare_digest(hashlib.sha256(k.encode()).hexdigest(), STATS_KEY_SHA256):
        return {"ok": False, "error": "not found"}, 404
    if request.args.get("flush"): stats_flush()
    try: data = stats_load()
    except Exception as e:
        app.logger.error("stats load failed: %s", e); data = {}
    with SLOCK: data = _merge(data, PENDING)
    return jsonify(stats_summary(data))

@app.get("/availability")
def availability():
    slots = slot_list(); bm = busy_map(slots)
    # customers only see open / taken
    return jsonify({"slots": [{"start": k, "taken": len(v) >= cap(dt.datetime.fromisoformat(k).hour)} for k, v in bm.items()]})

@app.route("/book", methods=["POST", "OPTIONS"])
def book():
    if request.method == "OPTIONS": return ("", 204)
    d = request.get_json(force=True, silent=True)
    if not isinstance(d, dict): d = {}
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
    stat_bump("bookings")
    notify(d, start, tech)   # outside the lock; failure never fails the booking
    return {"ok": True, "slot_start": start.isoformat()}

if __name__ == "__main__":
    app.run(port=int(os.environ.get("PORT", 8000)), threaded=True)
