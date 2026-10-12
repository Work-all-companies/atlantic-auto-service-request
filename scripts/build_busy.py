import os, json, datetime as dt, urllib.request
from zoneinfo import ZoneInfo
import icalendar, recurring_ical_events
TZ = ZoneInfo("America/Chicago")
raw = urllib.request.urlopen(os.environ["AAS_ICS_URL"], timeout=60).read()
cal = icalendar.Calendar.from_ical(raw)
now = dt.datetime.now(TZ)
start = dt.datetime.combine(now.date(), dt.time(), TZ)
end = start + dt.timedelta(days=43)
def to_local(v, is_end=False):
    if isinstance(v, dt.datetime):
        return (v if v.tzinfo else v.replace(tzinfo=TZ)).astimezone(TZ)
    return dt.datetime.combine(v, dt.time(), TZ)  # all-day: DTEND is exclusive midnight
blocks = set()
for ev in recurring_ical_events.of(cal).between(start, end):
    if str(ev.get("TRANSP", "OPAQUE")).upper() == "TRANSPARENT": continue
    if str(ev.get("X-MICROSOFT-CDO-BUSYSTATUS", "BUSY")).upper() == "FREE": continue
    if str(ev.get("STATUS", "")).upper() == "CANCELLED": continue
    s = to_local(ev["DTSTART"].dt); e = to_local(ev["DTEND"].dt) if ev.get("DTEND") else s + dt.timedelta(hours=1)
    if e <= start or s >= end: continue
    blocks.add((s.isoformat(timespec="minutes"), e.isoformat(timespec="minutes")))
out = {"timezone": "America/Chicago", "busy": [{"start": s, "end": e} for s, e in sorted(blocks)]}
new = json.dumps(out, indent=1) + "\n"
if not os.path.exists("busy.json") or open("busy.json").read() != new:
    open("busy.json", "w").write(new); print("updated", len(blocks))
else: print("unchanged")
