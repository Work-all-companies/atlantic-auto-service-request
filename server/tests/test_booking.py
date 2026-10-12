"""Run: cd server && python -m pytest tests -q   (uses FAKE_GRAPH, no network)"""
import os, json, tempfile, datetime as dt, threading
f = tempfile.NamedTemporaryFile("w", suffix=".json", delete=False); json.dump([], f); f.close()
os.environ["FAKE_GRAPH"] = f.name
import sys; sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import app as A

def hammer(hour):
    A.HOLDS.clear()
    s = [x for x in A.slot_list() if x.hour == hour][0].isoformat()
    c = A.app.test_client(); res = []
    def go(i):
        r = c.post("/book", json={"name": f"c{i}", "phone": "1", "vehicle": "v", "vin_or_plate": "x", "issue": "i", "slot_start": s})
        res.append(r.status_code)
    ts = [threading.Thread(target=go, args=(i,)) for i in range(30)]
    [t.start() for t in ts]; [t.join() for t in ts]
    return res, s

def test_two_per_hour():
    res, s = hammer(9); assert res.count(200) == 2 and res.count(409) == 28
    assert A.holds_for(s) == {"Brandon", "Kolby"}

def test_noon_one():
    res, _ = hammer(12); assert res.count(200) == 1 and res.count(409) == 29

def test_outlook_event_blocks():
    A.HOLDS.clear(); s = [x for x in A.slot_list() if x.hour == 10][0]
    u = s.astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")
    e = (s + dt.timedelta(hours=1)).astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")
    json.dump([{"subject": "oil change", "start": {"dateTime": u}, "end": {"dateTime": e}}], open(f.name, "w"))
    try:
        r = A.app.test_client().get("/availability").get_json()
        assert [x for x in r["slots"] if x["start"] == s.isoformat()][0]["taken"]
    finally: json.dump([], open(f.name, "w"))
