"""Run: cd server && python -m pytest tests -q   (uses FAKE_GRAPH, no network)"""
import os, json, tempfile, datetime as dt, threading
f = tempfile.NamedTemporaryFile("w", suffix=".json", delete=False); json.dump([], f); f.close()
os.environ["FAKE_GRAPH"] = f.name
import sys; sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import app as A
GOOD = {"name": "Jo", "phone": "(712) 555-1234", "email": "jo@x.com", "year": "2015", "make": "Ford",
        "model": "F-150", "vin": "", "plate": "abc123", "issue": "brakes"}

def hammer(hour):
    A.HOLDS.clear()
    s = [x for x in A.slot_list() if x.hour == hour][0].isoformat()
    c = A.app.test_client(); res = []
    def go(i):
        r = c.post("/book", json=dict(GOOD, name=f"c{i}", slot_start=s))
        res.append(r.status_code)
    ts = [threading.Thread(target=go, args=(i,)) for i in range(30)]
    [t.start() for t in ts]; [t.join() for t in ts]
    return res, s

def test_two_per_hour():
    res, s = hammer(9); assert res.count(200) == 2 and res.count(409) == 28
    assert A.holds_for(s) == {"Brandon", "Kolby"}

def test_noon_one():
    res, _ = hammer(12); assert res.count(200) == 1 and res.count(409) == 29

def _avail_with(subject, hour):
    A.HOLDS.clear(); s = [x for x in A.slot_list() if x.hour == hour][0]
    u = s.astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")
    e = (s + dt.timedelta(hours=1)).astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")
    json.dump([{"subject": subject, "start": {"dateTime": u}, "end": {"dateTime": e}}], open(f.name, "w"))
    try:
        r = A.app.test_client().get("/availability").get_json()
        return [x for x in r["slots"] if x["start"] == s.isoformat()][0], A.busy_map([s])[s.isoformat()]
    finally: json.dump([], open(f.name, "w"))

def test_untagged_event_ignored():
    slot, busy = _avail_with("2014 CHRYSLER TOWN AND COUNTRY", 10)
    assert not slot["taken"] and busy == set()
    slot, busy = _avail_with("oil change", 12)
    assert not slot["taken"] and busy == set()

def test_tagged_event_blocks_only_that_tech():
    slot, busy = _avail_with("kolby - 2014 Chrysler", 10)
    assert busy == {"Kolby"} and not slot["taken"]
    slot, busy = _avail_with("Brandon - Customer: X - Y", 12)
    assert busy == {"Brandon"} and slot["taken"]

def _book(hour):
    A.HOLDS.clear(); s = [x for x in A.slot_list() if x.hour == hour][0].isoformat()
    return A.app.test_client().post("/book", json=dict(GOOD, slot_start=s))

def test_mail_both_and_failure_safe():
    A.SENT_MAIL.clear()
    r = _book(14); assert r.status_code == 200
    m = A.SENT_MAIL[-1]["message"]
    assert [x["emailAddress"]["address"] for x in m["toRecipients"]] == [
        "bkeegan@atlanticglassco.com", "bschmidtke@atlanticautosalesco.com"]
    assert m["subject"].startswith("Service request: Jo - 2015 Ford F-150 - ")
    for k in ("Phone: (712) 555-1234", "jo@x.com", "ABC123", "brakes", "Tech: Brandon", "2:00 PM"): assert k in m["body"]["content"]
    orig = A.send_mail
    def boom(*a): raise RuntimeError("graph down")
    A.send_mail = boom
    try: assert _book(15).status_code == 200
    finally: A.send_mail = orig

def _slot(h=10): return [x for x in A.slot_list() if x.hour == h][0].isoformat()

def _bad(**kw):
    A.HOLDS.clear(); A.SENT_MAIL.clear()
    r = A.app.test_client().post("/book", json={**GOOD, "slot_start": _slot(), **kw})
    assert r.status_code == 400, r.get_json(); assert not A.SENT_MAIL and not A.HOLDS
    return r.get_json()["error"]

def test_required_fields():
    for k in ("name", "make", "model", "issue", "phone", "year", "slot_start"):
        assert _bad(**{k: ""})
    assert "VIN or a license plate" in _bad(vin="", plate="")
    assert A.validate(dict(GOOD, email="", slot_start="x"))[1] is None   # email optional

def test_phone():
    for p in ("555-1234", "712555123", "71255512345", "abc"): assert "10 digits" in _bad(phone=p)
    for p in ("7125551234", "712.555.1234", "+1 712 555 1234"):
        assert A.validate(dict(GOOD, phone=p, slot_start="x"))[0]["phone"] == "(712) 555-1234"

def test_vin():
    assert "17" in _bad(vin="1FTEW1EP5JFA1234")
    for v in ("1FTEW1EP5JFA1234I", "1FTEW1EP5JFA1234O", "1FTEW1EP5JFA1234Q", "1FTEW1EP5JFA1234!"):
        assert "no I, O or Q" in _bad(vin=v)
    c, e = A.validate(dict(GOOD, vin="1ftew1ep5jfa12345", plate="", slot_start="x"))
    assert e is None and c["vin"] == "1FTEW1EP5JFA12345" and c["vin_or_plate"] == "VIN 1FTEW1EP5JFA12345"

def test_plate():
    for p in ("A", "ABCDEFGHI", "AB#1"): assert "2-8" in _bad(vin="", plate=p)
    assert A.validate(dict(GOOD, plate="ab 12", slot_start="x"))[0]["plate"] == "AB12"

def test_year_dynamic():
    mx = A.dt.datetime.now(A.TZ).year + 1
    for y in ("1949", str(mx + 1), "15", "abcd", "20155"): assert str(mx) in _bad(year=y)
    for y in ("1950", str(mx)): assert A.validate(dict(GOOD, year=y, slot_start="x"))[1] is None

def test_event_title_and_mail_use_new_fields():
    A.SENT_MAIL.clear(); A.HOLDS.clear()
    r = A.app.test_client().post("/book", json=dict(GOOD, vin="1FTEW1EP5JFA12345", slot_start=_slot(11)))
    assert r.status_code == 200
    m = A.SENT_MAIL[-1]["message"]
    assert "Vehicle: 2015 Ford F-150" in m["body"]["content"] and "VIN 1FTEW1EP5JFA12345" in m["body"]["content"]

def _legacy_payload(**kw):
    p = {"name": "Old Form", "phone": "7125551234", "email": "", "vehicle": "2029 Ford F150",
         "vin_or_plate": "XYZ9876", "issue": "noise", "slot_start": _slot(9)}
    p.update(kw); return p

def test_legacy_year_too_high_400():
    r = A.app.test_client().post("/book", json=_legacy_payload())
    assert r.status_code == 400 and "year" in r.get_json()["error"].lower()
    assert r.headers.get("Access-Control-Allow-Origin")

def test_legacy_good_accepted():
    r = A.app.test_client().post("/book", json=_legacy_payload(vehicle="2018 Ford F-150", slot_start=_slot(14)))
    assert r.status_code == 200, r.get_json()

def test_legacy_no_year_400():
    r = A.app.test_client().post("/book", json=_legacy_payload(vehicle="Ford truck"))
    assert r.status_code == 400

def test_garbage_never_raises():
    for body in ([1, 2], "x", {"year": {"a": 1}, "make": ["x"]}, {"vehicle": 5}, None):
        r = A.app.test_client().post("/book", json=body)
        assert r.status_code == 400, (body, r.status_code)

def test_500_is_json_with_cors(monkeypatch):
    def boom(*a, **k): raise RuntimeError("x")
    monkeypatch.setattr(A, "slot_list", boom)
    A.app.config["PROPAGATE_EXCEPTIONS"] = False
    r = A.app.test_client().get("/availability")
    assert r.status_code == 500 and r.get_json()["ok"] is False
    assert r.headers.get("Access-Control-Allow-Origin")
