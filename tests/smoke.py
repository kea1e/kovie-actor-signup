"""Smoke test for the Kovie actor signup (live system).

Only Reese's own UARK address is ever booked, so no email reaches anyone else.
Every booking made here is cancelled at the end. Results go to smoke_results.json.
"""
import json, re, subprocess, sys, time, urllib.request, urllib.parse, concurrent.futures as cf

API = "https://script.google.com/macros/s/AKfycbxe89dKjVF6gacjOxsNZMcQFYFHRqqUDTnfRrGX5yC_NJYQZOAcY8l8MILyFUDTYrxM/exec"
SITE = "https://kea1e.github.io/kovie-actor-signup/"
ME = sys.argv[1]            # owner test address, passed in so it is not in this file
results = []

def rec(name, ok, detail=""):
    results.append({"check": name, "pass": bool(ok), "detail": str(detail)[:300]})
    print(("PASS " if ok else "FAIL ") + name + ("  | " + str(detail)[:200] if detail else ""))

def get(params=None, url=API):
    q = ("?" + urllib.parse.urlencode(params)) if params else ""
    with urllib.request.urlopen(url + q, timeout=60) as r:
        return r.status, r.read().decode()

def post(body):
    req = urllib.request.Request(API, data=json.dumps(body).encode(), headers={"Content-Type": "text/plain"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read().decode())

def state():
    return json.loads(get({"action": "state"})[1])

def seat(st, slot, n):
    return [s for s in st["slots"] if s["id"] == slot][0]["seats"][n - 1]

# ---------------------------------------------------------------- public surface
st_code, html = get(url=SITE)
rec("page loads", st_code == 200)
rec("noindex meta present", 'name="robots" content="noindex' in html)
_, cfg = get(url=SITE + "config.js")
rec("config.js points at the live API", API in cfg)
for path in ["apps-script/Code.gs", "apps-script/Secret.gs", ".relay_secret", ".clasp.json", "power-automate/clientdata.secret.json"]:
    try:
        c, _ = get(url=SITE + path)
    except urllib.error.HTTPError as e:
        c = e.code
    expect = 404 if ("Secret" in path or "secret" in path or "clasp" in path) else 200
    rec(f"{path} served? -> {c}", c == expect, "secret files must 404; source .gs is public by design" if expect == 200 else "")

st = state()
raw = json.dumps(st)
rec("state API ok", st.get("ok") is True)
rec("state API exposes no email addresses", "@" not in raw, f"{raw.count('@')} @ signs")
rec("state API exposes no tokens/ids", not re.search(r"[a-f0-9]{32}|K\d{10}", raw))
open_before = sum(x["status"] == "open" for s in st["slots"] for x in s["seats"])
rec("all 12 seats open before test", open_before == 12, open_before)
rec("unknown GET action falls back to state only", json.loads(get({"action": "dump"})[1]).get("slots") is not None)

# ---------------------------------------------------------------- validation (no email is sent on any of these)
r = post({"action": "hold", "slotId": "nope", "seat": 1}); rec("bad slot rejected", not r["ok"], r.get("error"))
r = post({"action": "hold", "slotId": "s1100", "seat": 3}); rec("seat out of range rejected", not r["ok"], r.get("error"))
r = post({"action": "hold", "slotId": "s1100", "seat": "1; drop"}); rec("non-numeric seat rejected", not r["ok"], r.get("error"))
r = post({"action": "book", "slotId": "s1100", "seat": 1, "name": "Test Person", "email": "someone@gmail.com"}); rec("non-UARK email rejected", not r["ok"], r.get("error"))
r = post({"action": "book", "slotId": "s1100", "seat": 1, "name": "Prince", "email": ME}); rec("single-word name rejected", not r["ok"], r.get("error"))
r = post({"action": "book", "slotId": "s1100", "seat": 1, "name": "A B", "email": "x@uark.edu.evil.com"}); rec("look-alike domain rejected", not r["ok"], r.get("error"))
r = post({"action": "nuke"}); rec("unknown POST action rejected", not r["ok"], r.get("error"))
r = post({"action": "cancel", "token": "0" * 32}); rec("bogus cancel token rejected", not r["ok"], r.get("error"))
r = post({"action": "cancel", "token": "../../etc"}); rec("malformed cancel token rejected", not r["ok"], r.get("error"))
r = json.loads(get({"action": "lookup", "token": "f" * 32})[1]); rec("bogus lookup rejected", not r["ok"], r.get("error"))
try:
    req = urllib.request.Request(API, data=b"not json", headers={"Content-Type": "text/plain"})
    r = json.loads(urllib.request.urlopen(req, timeout=60).read().decode())
    rec("garbage POST body handled", r.get("ok") is False, r.get("error"))
except Exception as e:
    rec("garbage POST body handled", False, e)

# ---------------------------------------------------------------- hold race: two people tap the same seat at once
with cf.ThreadPoolExecutor(2) as ex:
    holds = list(ex.map(lambda _: post({"action": "hold", "slotId": "s1300", "seat": 1}), range(2)))
winners = [h for h in holds if h.get("ok")]
rec("simultaneous holds: exactly one wins", len(winners) == 1, [h.get("ok") for h in holds])
rec("held seat shows as 'held' publicly", seat(state(), "s1300", 1)["status"] == "held")
r = post({"action": "book", "slotId": "s1300", "seat": 1, "name": "Sneaky Person", "email": ME, "holdId": "not-mine"})
rec("cannot book a seat someone else holds", not r["ok"], r.get("error"))
r = post({"action": "release", "slotId": "s1300", "seat": 1, "holdId": "not-mine"})
rec("cannot release someone else's hold", seat(state(), "s1300", 1)["status"] == "held")

# ---------------------------------------------------------------- real booking (sends Reese one confirmation + invite)
tricky = "Smoke \"Test\" O'Crawford & <b>Bold</b>"
r = post({"action": "book", "slotId": "s1300", "seat": 1, "name": tricky, "email": ME, "holdId": winners[0]["holdId"]})
rec("booking with my hold succeeds", r.get("ok"), r)
stored = r.get("name", "")
rec("angle brackets stripped from name", "<" not in stored and ">" not in stored, stored)
s = seat(state(), "s1300", 1)
rec("booked seat shows name, no email", s["status"] == "taken" and "@" not in json.dumps(s), s)

# double booking
r = post({"action": "hold", "slotId": "s1300", "seat": 1}); rec("taken seat cannot be held", not r["ok"], r.get("error"))
h = post({"action": "hold", "slotId": "s1300", "seat": 2})
r = post({"action": "book", "slotId": "s1300", "seat": 2, "name": "Smoke Test", "email": ME, "holdId": h.get("holdId")})
rec("same person, same time, other seat: refused", not r["ok"], r.get("error"))
post({"action": "release", "slotId": "s1300", "seat": 2, "holdId": h.get("holdId")})
rec("released hold frees the seat", seat(state(), "s1300", 2)["status"] == "open")
r = post({"action": "book", "slotId": "s1300", "seat": 2, "name": "Smoke Test", "email": ME.upper()})
rec("same-time check is case-insensitive", not r["ok"], r.get("error"))

json.dump({"results": results, "booked": "s1300/1"}, open("/tmp/kovie_smoke_results.json", "w"), indent=1)
print("\nphase 1 done:", sum(x["pass"] for x in results), "/", len(results))
