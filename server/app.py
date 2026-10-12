"""Atlantic Service & Repair booking server.
Availability = Outlook AAS SHOP events (live via Microsoft Graph) + local bookings table.
Each slot holds 2 jobs (Brandon, Kolby). Bookings are check-and-insert under a process lock
+ SQLite BEGIN IMMEDIATE transaction, so two customers can never take the same last spot.
Run with ONE worker (gunicorn -w 1 --threads 8)."""
import os, json, sqlite3, threading, datetime as dt, re, time, urllib.request, urllib.parse
from zoneinfo import ZoneInfo
from flask import Flask, request, jsonify

TZ = ZoneInfo("America/Chicago")
TECHS = ["Brandon", "Kolby"]                      # preference order
SLOT_HOURS = [8, 9, 10, 11, 13, 14, 15, 16]
DAYS_AHEAD = 30
DB = os.environ.get("DB_PATH", "bookings.db")
ALLOWED_ORIGIN = os.environ.get("ALLOWED_ORIGIN", "https://work-all-companies.github.io")
CLIENT_ID = os.environ.get("MS_CLIENT_ID", "")
CAL_ID = os.environ.get("AAS_CALENDAR_ID", "")
TOKEN_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
SCOPES = "Calendars.ReadWrite offline_access"
GRAPH = "https://graph.microsoft.com/v1.0"
FAKE_GRAPH = os.environ.get("FAKE_GRAPH")       # path to JSON list of events, for local tests

app = Flask(__name__)
LOCK = threading.Lock()

def db():
    c = sqlite3.connect(DB, timeout=30, isolation_level=None)
    c.execute("""CREATE TABLE IF NOT EXISTS bookings(id INTEGER PRIMARY KEY, slot_start TEXT, tech TEXT,
        name TEXT, phone TEXT, email TEXT, vehicle TEXT, vin TEXT, issue TEXT, created TEXT, event_id TEXT,
        UNIQUE(slot_start, tech))""")
    c.execute("CREATE TABLE IF NOT EXISTS kv(k TEXT PRIMARY KEY, v TEXT)")
    return c

# ---------- Microsoft Graph ----------
_tok = {"access": None, "exp": 0}
def access_token():
    if _tok["access"] and time.time() < _tok["exp"] - 60: return _tok["access"]
    c = db(); row = c.execute("SELECT v FROM kv WHERE k='refresh_token'").fetchone()
    rt = row[0] if row else os.environ.get("MS_REFRESH_TOKEN", "")
    if not rt: raise RuntimeError("no refresh token")
    body = urllib.parse.urlencode({"client_id": CLIENT_ID, "grant_type": "refresh_token",
                                   "refresh_token": rt, "scope": SCOPES}).encode()
    r = json.load(urllib.request.urlopen(TOKEN_URL, body, timeout=20))
    _tok.update(access=r["access_token"], exp=time.time() + int(r["expires_in"]))
    if r.get("refresh_token"):   # rotated token persisted
        c.execute("INSERT OR REPLACE INTO kv VALUES('refresh_token',?)", (r["refresh_token"],))
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
    return {m.group(1).capitalize()} if m else set(TECHS)   # no tech name -> blocks both

# ---------- availability ----------
def slot_list():
    today = dt.datetime.now(TZ).date(); out = []
    for i in range(1, DAYS_AHEAD + 1):
        d = today + dt.timedelta(days=i)
        if d.weekday() < 5:
            out += [dt.datetime.combine(d, dt.time(h), TZ) for h in SLOT_HOURS]
    return out

def busy_map(slots, conn):
    """slot_start iso -> set of busy techs."""
    events = outlook_events(slots[0], slots[-1] + dt.timedelta(hours=1)) if slots else []
    bk = {}
    for s, t, eid in conn.execute("SELECT slot_start, tech, event_id FROM bookings"):
        bk.setdefault(s, set()).add(t)
    res = {}
    for s in slots:
        e = s + dt.timedelta(hours=1); busy = set(bk.get(s.isoformat(), ()))
        for es, ee, subj in events:
            if es < e and ee > s: busy |= techs_in(subj)
        res[s.isoformat()] = busy
    return res

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
    slots = slot_list(); bm = busy_map(slots, db())
    # customers only see open / taken
    return jsonify({"slots": [{"start": k, "taken": len(v) >= len(TECHS)} for k, v in bm.items()]})

@app.route("/book", methods=["POST", "OPTIONS"])
def book():
    if request.method == "OPTIONS": return ("", 204)
    d = request.get_json(force=True, silent=True) or {}
    if d.get("_gotcha"): return {"ok": True}
    for k in ("name", "phone", "vehicle", "vin_or_plate", "issue", "slot_start"):
        if not str(d.get(k, "")).strip(): return {"ok": False, "error": "missing " + k}, 400
    try: start = dt.datetime.fromisoformat(d["slot_start"]).astimezone(TZ)
    except Exception: return {"ok": False, "error": "bad slot"}, 400
    if start not in slot_list(): return {"ok": False, "error": "slot not offered"}, 400
    with LOCK:
        c = db(); c.execute("BEGIN IMMEDIATE")
        try:
            busy = busy_map([start], c)[start.isoformat()]
            free = [t for t in TECHS if t not in busy]
            if not free:
                c.execute("ROLLBACK"); return {"ok": False, "error": "taken"}, 409
            tech = free[0]
            c.execute("INSERT INTO bookings(slot_start,tech,name,phone,email,vehicle,vin,issue,created) VALUES(?,?,?,?,?,?,?,?,?)",
                (start.isoformat(), tech, d["name"], d["phone"], d.get("email", ""), d["vehicle"],
                 d["vin_or_plate"], d["issue"], dt.datetime.now(TZ).isoformat()))
            bid = c.execute("SELECT last_insert_rowid()").fetchone()[0]
            c.execute("COMMIT")
        except Exception:
            c.execute("ROLLBACK"); raise
    # Outlook event (outside the lock; the DB row already holds the spot)
    try:
        if not FAKE_GRAPH:
            body = "\n".join(f"{k}: {d.get(k,'')}" for k in ("name","phone","email","vehicle","vin_or_plate","issue"))
            ev = graph("POST", f"/me/calendars/{CAL_ID}/events", {
                "subject": f"{tech} - Customer: {d['name']} - {d['vehicle']}",
                "body": {"contentType": "text", "content": body + "\n\nBooked online. Call customer to confirm."},
                "start": {"dateTime": start.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "Central Standard Time"},
                "end": {"dateTime": (start + dt.timedelta(hours=1)).strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "Central Standard Time"},
                "showAs": "busy"})
            c = db(); c.execute("UPDATE bookings SET event_id=? WHERE id=?", (ev["id"], bid))
    except Exception as e:
        app.logger.error("outlook event failed for booking %s: %s", bid, e)
    return {"ok": True, "slot_start": start.isoformat()}

if __name__ == "__main__":
    app.run(port=int(os.environ.get("PORT", 8000)), threaded=True)
