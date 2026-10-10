"""
PVH intel export -- notes, stops and queries from the shared Supabase project,
joined to the PVH records and turned into links between drivers and vehicles.

Usage:
    python pvh_intel.py                 sign in, pull, build everything
    python pvh_intel.py --offline       rebuild from the last pull (intel/snapshot.json)
    python pvh_intel.py --snapshot F --out DIR   (testing: any snapshot, any folder)

Signs in as you (the email is remembered in pvh_local_config.json as
"intel_email"; the password is asked every run and never stored). Reads only
what the app itself can read, with the same public key the app uses.

Writes to intel/ (gitignored -- it is personal information):
    PVH_Intel.xlsx      Profiles, Driver-Vehicle, Links, Notes, Activity,
                        Monthly, Officers, Hour x Weekday
    PVH_Intel_Map.html  clickable relationship map (opens in a browser)
    gephi_nodes.csv / gephi_edges.csv   Gephi "Import spreadsheet" format
    PVH_Intel.graphml   Gephi, Cytoscape, yEd, Neo4j
    snapshot.json       the raw pull, for --offline reruns

How links are found:
    On file        owner -> vehicle, from the PVH records
    Note mention   a note on one record names another: plate, vehicle or
                   operator licence number, "deck 211", or a full name
    Same encounter one officer queried / stopped / noted a vehicle and a
                   person within 15 minutes of each other
A driver-vehicle link that is not backed by the records is flagged NOT ON FILE.
"""

import os
import re
import sys
import json
import getpass
import argparse
import datetime
import urllib.request
import urllib.error
from collections import Counter, defaultdict
from zoneinfo import ZoneInfo

HERE = os.path.dirname(os.path.abspath(__file__))
CFG_PATH = os.path.join(HERE, "pvh_local_config.json")
DATA_PATH = os.path.join(HERE, "PVH_data.json")
TZ = ZoneInfo("America/Halifax")
ENCOUNTER_SECS = 15 * 60
TABLES = ("notes", "checks", "queries")


# ------------------------------------------------------------------ helpers
def fmt_when(dt):
    """10/09/26 @ 1535Hrs -- same as the app."""
    return dt.strftime("%m/%d/%y @ %H%MHrs") if dt else ""


def parse_ts(s):
    if not s:
        return None
    s = re.sub(r"(\.\d{6})\d+", r"\1", str(s)).replace("Z", "+00:00")
    try:
        return datetime.datetime.fromisoformat(s).astimezone(TZ)
    except ValueError:
        return None


def s_(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v)


def alnum(s):
    return re.sub(r"[^A-Z0-9]", "", s_(s).upper())


# ------------------------------------------------------------------ record keys
# These mirror clKey / clKeyV / nkW in the app exactly; a mismatch would
# orphan every note on that record.
def key_op(o):
    s = s_(o.get("Master Number")).strip().upper()
    s = re.sub(r"^4A\b", "", s)
    s = re.sub(r"^0*4\s+", "", s)
    s = re.sub(r"\s+", "", s)
    return "o:" + (s or "#" + (s_(o.get("ID")) or "?"))


def key_veh(v):
    s = re.sub(r"\s+", "", s_(v.get("Licence No")).strip().upper())
    return "v:" + (s or "#" + (s_(v.get("Vehicle ID")) or "?"))


def key_owner(w):
    nm = alnum(s_(w.get("Owner Last Name")) + s_(w.get("Owner First Name")) + s_(w.get("Business Name")))
    return "w:" + s_(w.get("Owner ID")).strip() + "|" + nm


# ------------------------------------------------------------------ Supabase
def http(method, url, key, token=None, body=None):
    h = {"apikey": key, "Content-Type": "application/json"}
    if token:
        h["Authorization"] = "Bearer " + token
    req = urllib.request.Request(url, method=method, headers=h,
                                 data=json.dumps(body).encode() if body is not None else None)
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.status, json.loads(r.read() or b"null")
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.loads(e.read() or b"null")
        except ValueError:
            return e.code, None


def pull(cfg):
    url = s_(cfg.get("supabase_url")).strip().rstrip("/")
    key = s_(cfg.get("supabase_anon_key")).strip()
    if not url or not key:
        sys.exit("pvh_local_config.json has no supabase_url / supabase_anon_key.")
    email = s_(cfg.get("intel_email")).strip()
    typed = input("Supabase email" + (" [" + email + "]" if email else "") + ": ").strip()
    if typed:
        email = typed
    pw = getpass.getpass("Password (not shown, not saved): ")
    st, j = http("POST", url + "/auth/v1/token?grant_type=password", key, body={"email": email, "password": pw})
    pw = None
    if st != 200 or not j or not j.get("access_token"):
        sys.exit("Sign-in refused: " + s_((j or {}).get("error_description") or (j or {}).get("msg") or st))
    token = j["access_token"]
    if typed and typed != cfg.get("intel_email"):
        try:
            raw = json.load(open(CFG_PATH, encoding="utf-8"))
            raw["intel_email"] = typed
            json.dump(raw, open(CFG_PATH, "w", encoding="utf-8"), indent=2)
        except (OSError, ValueError):
            pass
    snap = {"pulled_at": datetime.datetime.now(TZ).isoformat(), "by": email}
    for t in TABLES:
        rows, off = [], 0
        while True:
            st, page = http("GET", url + "/rest/v1/" + t + "?select=*&order=created_at.asc&limit=1000&offset=" + str(off), key, token)
            if st == 404 and t == "queries":
                print("  queries table not found -- run supabase/003_queries.sql. Continuing without it.")
                break
            if st != 200 or not isinstance(page, list):
                sys.exit("Could not read " + t + ": HTTP " + str(st) + " " + s_(page))
            rows += page
            if len(page) < 1000:
                break
            off += 1000
        snap[t] = rows
        print("  " + t + ": " + str(len(rows)))
    return snap


# ------------------------------------------------------------------ model
class Model:
    def __init__(self, db):
        self.V, self.O, self.W = db["vehicles"], db["operators"], db.get("owners", [])
        self.built = db.get("built", "")
        self.nodes = {}            # id -> dict
        self.alias = {}            # rec_key -> node id (owner-operators fold into the operator)
        self.veh_owner = {}        # vehicle node -> owner node (on file)
        for i, o in enumerate(self.O):
            k = key_op(o)
            self.add(k, "Operator", (" ".join(x for x in (s_(o.get("First Name")), s_(o.get("Middle")), s_(o.get("Last Name"))) if x) or "?"),
                     licence=s_(o.get("Licence Number")), company=s_(o.get("Business Name")),
                     detail=s_(o.get("Operator Type")) + " operator" + (" (inactive/cancelled)" if o.get("In Active") is True or o.get("Cancelled") is True else ""))
        for w in self.W:
            nk = key_owner(w)
            name = " ".join(x for x in (s_(w.get("Owner First Name")), s_(w.get("Owner Last Name"))) if x) or s_(w.get("Business Name")) or "?"
            if w.get("_op") is not None and w["_op"] < len(self.O):
                ok = key_op(self.O[w["_op"]])
                self.alias[nk] = ok
                self.nodes[ok]["type"] = "Owner-operator"
            else:
                self.add(nk, "Owner", name, company=s_(w.get("Business Name")), detail="Owner")
        for v in self.V:
            k = key_veh(v)
            deck = s_(v.get("Deck No"))
            lab = " · ".join(x for x in (("Deck " + deck) if deck else "", s_(v.get("Plate No")),
                                         (s_(v.get("Make Model")) + " " + s_(v.get("Vehicle Color"))).strip()) if x)
            self.add(k, "Vehicle", lab or k, licence=s_(v.get("Licence No")), plate=s_(v.get("Plate No")), deck=deck,
                     company=s_(v.get("Business Name")), detail=s_(v.get("Vehicle Type")))
            ow = v.get("_owner")
            if ow is not None and ow < len(self.W):
                self.veh_owner[k] = self.resolve(key_owner(self.W[ow]))
        self.build_patterns()

    def add(self, nid, typ, label, **kw):
        if nid in self.nodes:
            return self.nodes[nid]
        n = {"id": nid, "type": typ, "label": label, "licence": "", "plate": "", "deck": "", "company": "",
             "detail": "", "in_data": True, "notes": 0, "stops": 0, "queries": 0, "events": [], "officers": Counter()}
        n.update(kw)
        self.nodes[nid] = n
        return n

    def resolve(self, rec_key):
        return self.alias.get(rec_key, rec_key)

    def node_for(self, rec_key, rec_label="", rec_type=""):
        nid = self.resolve(rec_key)
        if nid not in self.nodes:
            typ = {"v": "Vehicle", "o": "Operator", "w": "Owner"}.get(rec_key[:1], (rec_type or "Record").title())
            self.add(nid, typ, rec_label or rec_key, in_data=False, detail="Not in the current PVH data")
        return self.nodes[nid]

    # --- mention patterns, built once from the records
    def build_patterns(self):
        pats = []   # (compiled regex, [node ids], method, confidence)

        def sep_rx(code):
            return r"[\s-]?".join(re.escape(c) for c in code)

        plates, lic, decks, names = defaultdict(list), defaultdict(list), defaultdict(list), defaultdict(list)
        for v in self.V:
            k = key_veh(v)
            if len(alnum(v.get("Plate No"))) >= 4:
                plates[alnum(v.get("Plate No"))].append(k)
            if len(alnum(v.get("Licence No"))) >= 3:
                lic[alnum(v.get("Licence No"))].append(k)
            if s_(v.get("Deck No")).strip():
                decks[s_(v.get("Deck No")).strip()].append(k)
        for o in self.O:
            k = key_op(o)
            if len(alnum(o.get("Licence Number"))) >= 3:
                lic[alnum(o.get("Licence Number"))].append(k)
            f, m, l = (s_(o.get(x)).strip() for x in ("First Name", "Middle", "Last Name"))
            if len(f) >= 2 and len(l) >= 2:
                names[(f + " " + l).upper()].append(k)
                if m:
                    names[(f + " " + m + " " + l).upper()].append(k)
                names[(l + ", " + f).upper()].append(k)
        for w in self.W:
            if w.get("_op") is not None:
                continue
            f, l = s_(w.get("Owner First Name")).strip(), s_(w.get("Owner Last Name")).strip()
            if len(f) >= 2 and len(l) >= 2:
                names[(f + " " + l).upper()].append(key_owner(w))
        for code, ks in plates.items():
            pats.append((re.compile(r"(?<![A-Z0-9])" + sep_rx(code) + r"(?![A-Z0-9])", re.I), ks, "plate"))
        for code, ks in lic.items():
            m = re.match(r"([A-Z]+)(\d+)$", code)
            rx = (re.escape(m.group(1)) + r"\s?-?" + m.group(2)) if m else re.escape(code)
            pats.append((re.compile(r"(?<![A-Z0-9])" + rx + r"(?![A-Z0-9])", re.I), ks, "licence no."))
        for nm, ks in names.items():
            rx = r"\s+".join(re.escape(p) for p in nm.split())
            pats.append((re.compile(r"(?<![A-Z])" + rx + r"(?![A-Z])", re.I), sorted(set(ks)), "name"))
        self.pats = pats
        self.decks = decks
        self.deck_rx = re.compile(r"\b(?:deck|dk|car|unit|light)\s*(?:no\.?|number|#)?\s*#?\s*(\d{1,4})\b|#\s?(\d{1,4})\b", re.I)

    def mentions(self, text, self_id):
        out, seen = [], set()
        for rx, ks, method in self.pats:
            for m in rx.finditer(text):
                ids = sorted({self.resolve(k) for k in ks} - {self_id})
                for nid in ids:
                    if (nid, method) in seen:
                        continue
                    seen.add((nid, method))
                    out.append({"id": nid, "method": method, "text": m.group(0),
                                "confidence": "exact" if len(ids) == 1 else "possible (" + str(len(ids)) + " share it)"})
        for m in self.deck_rx.finditer(text):
            d = m.group(1) or m.group(2)
            ids = sorted(set(self.decks.get(d, [])) - {self_id})
            for nid in ids:
                if (nid, "deck") in seen or any(x["id"] == nid for x in out):
                    continue
                seen.add((nid, "deck"))
                out.append({"id": nid, "method": "deck", "text": m.group(0),
                            "confidence": "likely" if len(ids) == 1 else "possible (" + str(len(ids)) + " share it)"})
        # One mention per record per note: "<full name>, TD0000" is one
        # sighting, found two ways. Keep the strongest confidence.
        rank = lambda c: 0 if c == "exact" else (1 if c == "likely" else 2)
        merged = {}
        for m in out:
            cur = merged.get(m["id"])
            if not cur:
                merged[m["id"]] = dict(m)
                continue
            cur["method"] += " + " + m["method"]
            cur["text"] += "”, “" + m["text"]
            if rank(m["confidence"]) < rank(cur["confidence"]):
                cur["confidence"] = m["confidence"]
        return list(merged.values())


# ------------------------------------------------------------------ analysis
def analyse(model, snap):
    events, notes, links = [], [], {}
    officers = {}

    def officer(row):
        oid = s_(row.get("author_id")) or ("name:" + s_(row.get("author_name")))
        officers.setdefault(oid, s_(row.get("author_name")) or "?")
        return oid

    def link(a, b, rel, when, ev):
        if a == b:
            return
        a, b = sorted((a, b))
        L = links.setdefault((a, b, rel), {"source": a, "target": b, "relation": rel, "count": 0,
                                          "first": None, "last": None, "officers": Counter(), "evidence": []})
        L["count"] += 1
        if when:
            L["first"] = min(L["first"] or when, when)
            L["last"] = max(L["last"] or when, when)
        if ev.get("officer"):
            L["officers"][ev["officer"]] += 1
        if len(L["evidence"]) < 20:
            L["evidence"].append(ev)

    for kind, table in (("Stop", "checks"), ("Query", "queries")):
        for r in snap.get(table, []):
            n = model.node_for(s_(r.get("rec_key")), s_(r.get("rec_label")))
            when = parse_ts(r.get("written_at"))
            oid = officer(r)
            n["stops" if kind == "Stop" else "queries"] += 1
            n["officers"][officers[oid]] += 1
            events.append({"kind": kind, "node": n["id"], "when": when, "oid": oid, "officer": officers[oid], "text": ""})

    for r in snap.get("notes", []):
        n = model.node_for(s_(r.get("rec_key")), s_(r.get("rec_label")), s_(r.get("rec_type")))
        when = parse_ts(r.get("written_at"))
        oid = officer(r)
        removed = bool(r.get("removed_at"))
        body = s_(r.get("body"))
        ms = [] if removed else model.mentions(body, n["id"])
        status = ("Removed by " + s_(r.get("removed_by"))) if removed else \
                 ("Done by " + s_(r.get("done_by"))) if r.get("done_at") else ("Open" if r.get("action_needed") else "")
        notes.append({"node": n["id"], "when": when, "officer": officers[oid], "body": body,
                      "action": bool(r.get("action_needed")), "status": status, "removed": removed,
                      "edited": (fmt_when(parse_ts(r.get("edited_at"))) + " by " + s_(r.get("edited_by"))) if r.get("edited_at") else "",
                      "mentions": ms})
        if not removed:
            n["notes"] += 1
            n["officers"][officers[oid]] += 1
            events.append({"kind": "Note", "node": n["id"], "when": when, "oid": oid, "officer": officers[oid], "text": body})
        for m in ms:
            model.node_for(m["id"])
            link(n["id"], m["id"], "Note mention", when,
                 {"officer": officers[oid], "when": when, "how": m["method"] + " \u201c" + m["text"] + "\u201d, " + m["confidence"],
                  "text": body, "confidence": m["confidence"]})

    # same encounter: one officer, a vehicle and a person within 15 minutes
    by_off = defaultdict(list)
    for e in events:
        if e["when"]:
            by_off[e["oid"]].append(e)
    pair_last = {}
    for oid, evs in by_off.items():
        evs.sort(key=lambda e: e["when"])
        for i, a in enumerate(evs):
            for b in evs[i + 1:]:
                if (b["when"] - a["when"]).total_seconds() > ENCOUNTER_SECS:
                    break
                ta, tb = model.nodes[a["node"]]["type"], model.nodes[b["node"]]["type"]
                if (ta == "Vehicle") == (tb == "Vehicle"):
                    continue
                pk = (oid,) + tuple(sorted((a["node"], b["node"])))
                lw = pair_last.get(pk)
                if lw and (a["when"] - lw).total_seconds() <= ENCOUNTER_SECS:
                    continue
                pair_last[pk] = b["when"]
                link(a["node"], b["node"], "Same encounter", a["when"],
                     {"officer": a["officer"], "when": a["when"],
                      "how": a["kind"] + " + " + b["kind"] + " within " + str(int((b["when"] - a["when"]).total_seconds() // 60)) + " min",
                      "text": a["text"] or b["text"], "confidence": "likely"})

    for e in events:
        model.nodes[e["node"]]["events"].append(e)

    # driver <-> vehicle view, with the on-file check
    dv = []
    for L in links.values():
        a, b = model.nodes[L["source"]], model.nodes[L["target"]]
        if (a["type"] == "Vehicle") == (b["type"] == "Vehicle"):
            continue
        veh, per = (a, b) if a["type"] == "Vehicle" else (b, a)
        L["on_file"] = model.veh_owner.get(veh["id"]) == per["id"]
        dv.append((veh, per, L))
    return events, notes, links, dv, officers


def active_set(model, links):
    """Records with any activity, everything they link to, and their on-file
    owner / vehicles -- the part of the registry worth drawing."""
    act = {nid for nid, n in model.nodes.items() if n["notes"] or n["stops"] or n["queries"]}
    for L in links.values():
        act |= {L["source"], L["target"]}
    ring = set(act)
    for v, o in model.veh_owner.items():
        if v in act or o in act:
            ring |= {v, o}
    return act, ring


# ------------------------------------------------------------------ outputs
def write_xlsx(path, model, events, notes, links, dv, officers, act, meta):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    head = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="1F4E79")
    red = PatternFill("solid", fgColor="F8D7DA")
    wrap = Alignment(wrap_text=True, vertical="top")

    def sheet(title, cols, rows, widths=None, wrap_cols=()):
        ws = wb.create_sheet(title)
        ws.append(cols)
        for c in ws[1]:
            c.font, c.fill = head, fill
        for r in rows:
            ws.append(r)
        ws.freeze_panes = "A2"
        if rows:
            ws.auto_filter.ref = ws.dimensions
        for i, c in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = (widths or {}).get(c, max(10, min(40, len(c) + 4)))
            if c in wrap_cols:
                for cell in ws[get_column_letter(i)][1:]:
                    cell.alignment = wrap
        return ws

    N = model.nodes
    lab = lambda nid: N[nid]["label"]

    ws = wb.active
    ws.title = "Read me"
    for line in [
        ["PVH intel export"], [],
        ["Pulled", meta["pulled"]], ["Pulled by", meta["by"]], ["PVH data built", model.built],
        ["Notes / stops / queries", "%d / %d / %d" % (meta["n_notes"], meta["n_stops"], meta["n_queries"])], [],
        ["Sheets"],
        ["Profiles", "Every record with activity: counts, last activity, who, and what it links to."],
        ["Driver-Vehicle", "Drivers linked to vehicles by notes or by the same encounter. NOT ON FILE = not the registered owner."],
        ["Links", "Every link with its evidence."],
        ["Notes", "All notes, with the records each one mentions."],
        ["Activity", "One row per query / stop / note -- use for pivots."],
        ["Monthly", "Activity per record per month."],
        ["Officers", "Activity per officer."],
        ["Hour x Weekday", "When activity happens."], [],
        ["Link types"],
        ["On file", "Owner -> vehicle in the PVH records."],
        ["Note mention", "A note on one record names another (plate, licence no., deck, full name)."],
        ["Same encounter", "One officer touched a vehicle and a person within 15 minutes."],
        ["Confidence", "exact = one match; likely = deck number or encounter; possible = several records share that name or number -- verify."], [],
        ["This file holds personal information. Keep it on this PC; do not email or share it."],
    ]:
        ws.append(line)
    ws["A1"].font = Font(bold=True, size=14)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 100

    now = datetime.datetime.now(TZ)
    adj = defaultdict(set)
    for L in links.values():
        adj[L["source"]].add(L["target"])
        adj[L["target"]].add(L["source"])
    nof = Counter()
    for veh, per, L in dv:
        if not L["on_file"]:
            nof[veh["id"]] += 1
            nof[per["id"]] += 1
    prof = []
    for nid in sorted(act, key=lambda k: -(N[k]["notes"] + N[k]["stops"] + N[k]["queries"])):
        n = N[nid]
        evs = sorted((e for e in n["events"] if e["when"]), key=lambda e: e["when"])
        last = evs[-1] if evs else None
        hrs = Counter(e["when"].hour for e in evs)
        prof.append([n["type"], n["label"], n["licence"] or n["plate"], n["company"], n["notes"], n["stops"], n["queries"],
                     sum(1 for e in evs if (now - e["when"]).days < 30),
                     fmt_when(last["when"]) if last else "", (last["kind"] + " \u00B7 " + last["officer"]) if last else "",
                     len(n["officers"]), ", ".join(o for o, _ in n["officers"].most_common(3)),
                     ("%02d00-%02d59" % (hrs.most_common(1)[0][0], hrs.most_common(1)[0][0])) if hrs else "",
                     nof[nid] or "", "; ".join(sorted(lab(x) for x in adj[nid]))[:500],
                     "" if n["in_data"] else "Not in current data", nid])
    sheet("Profiles", ["Type", "Record", "Licence / plate", "Company", "Notes", "Stops", "Queries", "Last 30 days",
                       "Last activity", "Last by", "Officers", "Top officers", "Busiest hour", "NOT ON FILE links",
                       "Linked to", "Flag", "Key"], prof,
          {"Record": 34, "Company": 24, "Linked to": 60, "Last activity": 18, "Last by": 22, "Top officers": 30, "Key": 22})

    rows = []
    for veh, per, L in sorted(dv, key=lambda x: (x[2]["on_file"], -x[2]["count"])):
        ev = L["evidence"][0] if L["evidence"] else {}
        rows.append([per["label"], per["licence"], veh["label"], veh["plate"], veh["deck"],
                     N[model.veh_owner[veh["id"]]]["label"] if veh["id"] in model.veh_owner else "",
                     "Yes" if L["on_file"] else "NOT ON FILE", L["relation"], L["count"],
                     fmt_when(L["first"]), fmt_when(L["last"]), ", ".join(L["officers"]),
                     ev.get("confidence", ""), ev.get("how", ""), ev.get("text", "")[:500]])
    ws = sheet("Driver-Vehicle", ["Driver / person", "Licence", "Vehicle", "Plate", "Deck", "Registered owner", "On file",
                                  "Evidence", "Times", "First", "Last", "Officers", "Confidence", "How found", "Note"],
               rows, {"Driver / person": 26, "Vehicle": 34, "Registered owner": 24, "First": 18, "Last": 18,
                      "How found": 36, "Note": 60}, wrap_cols=("Note",))
    for row in ws.iter_rows(min_row=2):
        if row[6].value == "NOT ON FILE":
            for c in row:
                c.fill = red

    rows = []
    for L in sorted(links.values(), key=lambda L: -L["count"]):
        for ev in L["evidence"]:
            rows.append([lab(L["source"]), N[L["source"]]["type"], lab(L["target"]), N[L["target"]]["type"], L["relation"],
                         fmt_when(ev.get("when")), ev.get("officer", ""), ev.get("confidence", ""), ev.get("how", ""),
                         ev.get("text", "")[:500]])
    sheet("Links", ["Record A", "Type A", "Record B", "Type B", "Link", "When", "Officer", "Confidence", "How found", "Note"],
          rows, {"Record A": 30, "Record B": 30, "When": 18, "How found": 36, "Note": 60}, wrap_cols=("Note",))

    rows = []
    for nt in sorted(notes, key=lambda x: x["when"] or now, reverse=True):
        n = N[nt["node"]]
        rows.append([fmt_when(nt["when"]), nt["officer"], n["type"], n["label"], nt["body"], "Yes" if nt["action"] else "",
                     nt["status"], nt["edited"], "; ".join(lab(m["id"]) + " (" + m["method"] + ")" for m in nt["mentions"])])
    sheet("Notes", ["Written", "Officer", "Record type", "Record", "Note", "Action needed", "Status", "Edited", "Mentions"],
          rows, {"Written": 18, "Record": 32, "Note": 70, "Status": 22, "Edited": 26, "Mentions": 40}, wrap_cols=("Note", "Mentions"))

    rows = []
    for e in sorted(events, key=lambda e: e["when"] or now, reverse=True):
        n, w = N[e["node"]], e["when"]
        rows.append([w.replace(tzinfo=None) if w else None, fmt_when(w), w.strftime("%A") if w else "", w.hour if w else None,
                     w.strftime("%Y-%m") if w else "", e["kind"], e["officer"], n["type"], n["label"], e["node"]])
    ws = sheet("Activity", ["Date/time", "Written", "Weekday", "Hour", "Month", "Kind", "Officer", "Record type", "Record", "Key"],
               rows, {"Date/time": 18, "Written": 18, "Record": 34, "Key": 22})
    for c in ws["A"][1:]:
        c.number_format = "yyyy-mm-dd hh:mm"

    months = sorted({e["when"].strftime("%Y-%m") for e in events if e["when"]})
    per = defaultdict(Counter)
    for e in events:
        if e["when"]:
            per[e["node"]][e["when"].strftime("%Y-%m")] += 1
    rows = [[N[k]["type"], N[k]["label"]] + [per[k][m] or None for m in months] + [sum(per[k].values())]
            for k in sorted(per, key=lambda k: -sum(per[k].values()))]
    ws = sheet("Monthly", ["Type", "Record"] + months + ["Total"], rows, {"Record": 34})
    if rows and months:
        rng = "C2:" + get_column_letter(2 + len(months)) + str(len(rows) + 1)
        ws.conditional_formatting.add(rng, ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="5B9BD5"))

    offs = defaultdict(lambda: {"Query": 0, "Stop": 0, "Note": 0, "recs": set(), "hrs": Counter(), "last": None})
    for e in events:
        o = offs[e["officer"]]
        o[e["kind"]] += 1
        o["recs"].add(e["node"])
        if e["when"]:
            o["hrs"][e["when"].hour] += 1
            o["last"] = max(o["last"] or e["when"], e["when"])
    rows = [[k, o["Query"], o["Stop"], o["Note"], len(o["recs"]),
             ("%02d00" % o["hrs"].most_common(1)[0][0]) if o["hrs"] else "", fmt_when(o["last"])]
            for k, o in sorted(offs.items(), key=lambda kv: -(kv[1]["Query"] + kv[1]["Stop"] + kv[1]["Note"]))]
    sheet("Officers", ["Officer", "Queries", "Stops", "Notes", "Records touched", "Busiest hour", "Last activity"], rows,
          {"Officer": 22, "Last activity": 18})

    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    grid = Counter((e["when"].weekday(), e["when"].hour) for e in events if e["when"])
    rows = [["%02d00" % h] + [grid[(d, h)] or None for d in range(7)] for h in range(24)]
    ws = sheet("Hour x Weekday", ["Hour"] + days, rows)
    ws.conditional_formatting.add("B2:H25", ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="F4B183"))
    wb.save(path)


def graph_rows(model, links, ring, officers_too=True):
    N = model.nodes
    nodes = [N[k] for k in sorted(ring)]
    edges = []
    for v, o in model.veh_owner.items():
        if v in ring and o in ring:
            edges.append({"source": o, "target": v, "relation": "On file", "count": 1, "first": None, "last": None,
                          "on_file": True, "officers": Counter(), "evidence": []})
    edges += list(links.values())
    off_nodes, off_edges = {}, []
    if officers_too:
        for k in ring:
            for kind_counts in [Counter((e["officer"], e["kind"]) for e in N[k]["events"])]:
                for (off, kind), c in kind_counts.items():
                    oid = "officer:" + off
                    off_nodes[oid] = {"id": oid, "type": "Officer", "label": off, "licence": "", "plate": "", "deck": "",
                                      "company": "", "detail": "Officer", "in_data": True, "notes": 0, "stops": 0,
                                      "queries": 0, "events": [], "officers": Counter()}
                    off_edges.append({"source": oid, "target": k, "relation": {"Query": "Queried", "Stop": "Stopped", "Note": "Noted"}[kind],
                                      "count": c, "first": None, "last": None, "officers": Counter(), "evidence": []})
    return nodes + list(off_nodes.values()), edges + off_edges


def write_gephi(dir_, model, links, ring):
    import csv
    nodes, edges = graph_rows(model, links, ring)
    with open(os.path.join(dir_, "gephi_nodes.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Id", "Label", "Category", "Licence", "Plate", "Deck", "Company", "Notes", "Stops", "Queries", "InCurrentData"])
        for n in nodes:
            w.writerow([n["id"], n["label"], n["type"], n["licence"], n["plate"], n["deck"], n["company"],
                        n["notes"], n["stops"], n["queries"], n["in_data"]])
    with open(os.path.join(dir_, "gephi_edges.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        # Gephi reads "Type" as Directed/Undirected, so the link kind goes in Relation.
        w.writerow(["Source", "Target", "Type", "Weight", "Relation", "OnFile", "First", "Last"])
        for e in edges:
            w.writerow([e["source"], e["target"], "Undirected", e["count"], e["relation"],
                        "" if "on_file" not in e else ("yes" if e["on_file"] else "NOT ON FILE"),
                        e["first"].isoformat() if e["first"] else "", e["last"].isoformat() if e["last"] else ""])
    from xml.sax.saxutils import escape as x
    na = [("label", "string"), ("category", "string"), ("licence", "string"), ("plate", "string"), ("company", "string"),
          ("notes", "int"), ("stops", "int"), ("queries", "int")]
    ea = [("relation", "string"), ("weight", "double"), ("onfile", "string"), ("first", "string"), ("last", "string")]
    out = ['<?xml version="1.0" encoding="UTF-8"?>',
           '<graphml xmlns="http://graphml.graphdrawing.org/xmlns">']
    out += ['<key id="n_%s" for="node" attr.name="%s" attr.type="%s"/>' % (k, k, t) for k, t in na]
    out += ['<key id="e_%s" for="edge" attr.name="%s" attr.type="%s"/>' % (k, k, t) for k, t in ea]
    out.append('<graph id="PVH" edgedefault="undirected">')
    for n in nodes:
        vals = {"label": n["label"], "category": n["type"], "licence": n["licence"], "plate": n["plate"], "company": n["company"],
                "notes": n["notes"], "stops": n["stops"], "queries": n["queries"]}
        out.append('<node id="%s">' % x(n["id"], {'"': "&quot;"}) + "".join('<data key="n_%s">%s</data>' % (k, x(s_(vals[k]))) for k, _ in na) + "</node>")
    for i, e in enumerate(edges):
        vals = {"relation": e["relation"], "weight": e["count"],
                "onfile": "" if "on_file" not in e else ("yes" if e["on_file"] else "NOT ON FILE"),
                "first": e["first"].isoformat() if e["first"] else "", "last": e["last"].isoformat() if e["last"] else ""}
        out.append('<edge id="e%d" source="%s" target="%s">' % (i, x(e["source"], {'"': "&quot;"}), x(e["target"], {'"': "&quot;"})) +
                   "".join('<data key="e_%s">%s</data>' % (k, x(s_(vals[k]))) for k, _ in ea) + "</edge>")
    out += ["</graph>", "</graphml>"]
    open(os.path.join(dir_, "PVH_Intel.graphml"), "w", encoding="utf-8").write("\n".join(out))


def write_map(path, model, links, notes, ring, meta):
    nodes, edges = graph_rows(model, links, ring)
    N = model.nodes
    by_node = defaultdict(list)
    for nt in notes:
        if not nt["removed"]:
            by_node[nt["node"]].append({"w": fmt_when(nt["when"]), "o": nt["officer"], "b": nt["body"]})
    data = {"nodes": [], "edges": [], "meta": meta}
    for n in nodes:
        evs = sorted((e for e in n["events"] if e["when"]), key=lambda e: e["when"], reverse=True)
        data["nodes"].append({"id": n["id"], "label": n["label"], "type": n["type"], "lic": n["licence"] or n["plate"],
                              "co": n["company"], "n": n["notes"], "s": n["stops"], "q": n["queries"], "in": n["in_data"],
                              "ev": [[e["kind"], fmt_when(e["when"]), e["officer"]] for e in evs[:40]],
                              "notes": by_node.get(n["id"], [])[:40]})
    for e in edges:
        data["edges"].append({"from": e["source"], "to": e["target"], "rel": e["relation"], "c": e["count"],
                              "nof": ("on_file" in e and not e["on_file"]),
                              "ev": [[fmt_when(v.get("when")), v.get("officer", ""), v.get("how", ""), v.get("text", "")[:300]] for v in e["evidence"][:10]]})
    html = MAP_HTML.replace("__DATA__", json.dumps(data).replace("</", "<\\/"))
    open(path, "w", encoding="utf-8").write(html)


MAP_HTML = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>PVH Intel Map</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/vis-network/9.1.9/standalone/umd/vis-network.min.js"></script>
<style>
:root{--bg:#f4f6f9;--panel:#fff;--line:#d8dee6;--text:#1b2430;--dim:#5b6675;--accent:#0a62c6;--bad:#c42b2b}
@media (prefers-color-scheme:dark){:root{--bg:#10151c;--panel:#18202a;--line:#2b3542;--text:#e6ebf1;--dim:#93a0b0;--accent:#5faeff;--bad:#ff6363}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font:14px/1.4 -apple-system,"Segoe UI",Roboto,sans-serif;height:100vh;display:flex;flex-direction:column}
header{display:flex;gap:10px;align-items:center;flex-wrap:wrap;padding:8px 12px;border-bottom:1px solid var(--line);background:var(--panel)}
header b{font-size:15px;letter-spacing:.04em}header small{color:var(--dim)}
input[type=search]{flex:0 1 240px;padding:6px 10px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--text)}
label{display:inline-flex;gap:4px;align-items:center;color:var(--dim);cursor:pointer}
#wrap{flex:1;display:flex;min-height:0}#net{flex:1;min-width:0}
#side{width:380px;max-width:45vw;overflow:auto;border-left:1px solid var(--line);background:var(--panel);padding:12px}
#side h2{font-size:16px;margin:0 0 4px}.k{color:var(--dim);font-size:12px}
.sec{margin-top:12px;font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--dim)}
.it{border-top:1px solid var(--line);padding:6px 0}.it .k{display:block}.nof{color:var(--bad);font-weight:700}
.lg{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:3px;vertical-align:middle}
a{color:var(--accent);cursor:pointer}
</style></head><body>
<header><b>PVH INTEL MAP</b><small id="meta"></small>
<input type="search" id="q" placeholder="Find name, plate, licence…">
<span id="types"></span>
<label><input type="checkbox" id="onlynof"> Only NOT ON FILE</label>
</header>
<div id="wrap"><div id="net"></div><div id="side"><div class="k">Click a record or a link for details.<br><br>
Lines: <span style="color:#9aa5b1">grey</span> on file · <span style="color:#d9a21b">amber</span> note mention ·
<span style="color:#3d8bfd">blue</span> same encounter · <span style="color:#c42b2b">red dashed</span> NOT ON FILE · dotted = officer activity</div></div></div>
<script>
const D=__DATA__;
const COL={"Vehicle":"#f0b429","Operator":"#3d8bfd","Owner-operator":"#7b61ff","Owner":"#8a94a3","Officer":"#2fb67c","Record":"#bbb"};
const SHAPE={"Vehicle":"box","Officer":"diamond"};
const ECOL={"On file":"#9aa5b1","Note mention":"#d9a21b","Same encounter":"#3d8bfd"};
const esc=s=>String(s==null?"":s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
document.getElementById("meta").textContent="pulled "+D.meta.pulled+" · "+D.meta.n_notes+" notes · "+D.meta.n_stops+" stops · "+D.meta.n_queries+" queries";
const byId={};D.nodes.forEach(n=>byId[n.id]=n);
const show={};Object.keys(COL).forEach(t=>show[t]=t!=="Officer");
document.getElementById("types").innerHTML=Object.keys(COL).filter(t=>D.nodes.some(n=>n.type===t)).map(t=>
  '<label><input type="checkbox" data-t="'+t+'" '+(show[t]?"checked":"")+'><span class="lg" style="background:'+COL[t]+'"></span>'+t+'</label>').join(" ");
const act=n=>n.n+n.s+n.q;
const nodes=new vis.DataSet(D.nodes.map(n=>({id:n.id,label:n.label.length>34?n.label.slice(0,32)+"…":n.label,
  shape:SHAPE[n.type]||"dot",size:10+Math.min(30,act(n)*2),color:{background:COL[n.type]||"#bbb",border:n.in?"#0000":"#c42b2b"},
  borderWidth:n.in?0:2,font:{color:getComputedStyle(document.body).color,size:12},title:n.type+": "+n.label})));
const edges=new vis.DataSet(D.edges.map((e,i)=>({id:i,from:e.from,to:e.to,width:Math.min(8,1+e.c),
  color:{color:e.nof?"#c42b2b":(ECOL[e.rel]||"#2fb67c"),opacity:.85},dashes:e.nof?[8,5]:(/Queried|Stopped|Noted/.test(e.rel)?[2,4]:false),
  title:e.rel+(e.nof?" — NOT ON FILE":"")+(e.c>1?" ×"+e.c:"")})));
const view=new vis.DataView(nodes,{filter:n=>{const d=byId[n.id];if(!show[d.type])return false;
  if(document.getElementById("onlynof").checked){return D.edges.some(e=>e.nof&&(e.from===n.id||e.to===n.id));}return true;}});
const net=new vis.Network(document.getElementById("net"),{nodes:view,edges:edges},
  {physics:{solver:"forceAtlas2Based",stabilization:{iterations:300}},interaction:{hover:true,multiselect:false}});
document.getElementById("types").addEventListener("change",e=>{show[e.target.dataset.t]=e.target.checked;view.refresh();});
document.getElementById("onlynof").addEventListener("change",()=>view.refresh());
document.getElementById("q").addEventListener("keydown",e=>{if(e.key!=="Enter")return;const s=e.target.value.trim().toLowerCase();if(!s)return;
  const hit=D.nodes.find(n=>show[n.type]&&(n.label.toLowerCase().includes(s)||String(n.lic).toLowerCase().includes(s)));
  if(hit){net.selectNodes([hit.id]);net.focus(hit.id,{scale:1.4,animation:true});side(hit.id);}});
function side(id){const n=byId[id];if(!n)return;
  const ls=D.edges.map((e,i)=>[e,i]).filter(([e])=>e.from===id||e.to===id);
  let h='<h2>'+esc(n.label)+'</h2><div class="k">'+esc(n.type)+(n.lic?' · '+esc(n.lic):'')+(n.co?' · '+esc(n.co):'')+(n.in?'':' · <span class="nof">not in current data</span>')+'</div>'+
    '<div class="k">'+n.n+' notes · '+n.s+' stops · '+n.q+' queries</div>';
  h+='<div class="sec">Links ('+ls.length+')</div>'+ls.map(([e,i])=>{const o=byId[e.from===id?e.to:e.from];
    return '<div class="it"><a onclick="pick(\''+esc(o.id).replace(/'/g,"\\'")+'\')">'+esc(o.label)+'</a> <span class="k">'+esc(e.rel)+(e.c>1?' ×'+e.c:'')+'</span>'+(e.nof?' <span class="nof">NOT ON FILE</span>':'')+
      e.ev.map(v=>'<span class="k">'+esc(v[0])+' · '+esc(v[1])+' · '+esc(v[2])+(v[3]?' — '+esc(v[3]):'')+'</span>').join("")+'</div>';}).join("");
  if(n.notes.length)h+='<div class="sec">Notes</div>'+n.notes.map(x=>'<div class="it"><span class="k">'+esc(x.w)+' · '+esc(x.o)+'</span>'+esc(x.b)+'</div>').join("");
  if(n.ev.length)h+='<div class="sec">Activity</div>'+n.ev.map(x=>'<div class="it"><b>'+esc(x[0])+'</b> '+esc(x[1])+' · '+esc(x[2])+'</div>').join("");
  document.getElementById("side").innerHTML=h;}
function pick(id){const d=byId[id];if(d&&!show[d.type]){show[d.type]=true;document.querySelector('[data-t="'+d.type+'"]').checked=true;view.refresh();}
  net.selectNodes([id]);net.focus(id,{scale:1.3,animation:true});side(id);}
net.on("click",p=>{if(p.nodes.length)side(p.nodes[0]);else if(p.edges.length){const e=D.edges[p.edges[0]];side(e.from);}});
</script></body></html>
"""


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser(description="PVH intel export")
    ap.add_argument("--offline", action="store_true", help="rebuild from the last pull")
    ap.add_argument("--snapshot", help="snapshot file to read instead of pulling")
    ap.add_argument("--out", default=os.path.join(HERE, "intel"), help="output folder")
    ap.add_argument("--data", default=DATA_PATH, help="PVH_data.json")
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    snap_path = a.snapshot or os.path.join(a.out, "snapshot.json")
    if a.offline or a.snapshot:
        if not os.path.exists(snap_path):
            sys.exit("No snapshot at " + snap_path + " -- run once without --offline.")
        snap = json.load(open(snap_path, encoding="utf-8"))
        print("Using " + snap_path + " (pulled " + s_(snap.get("pulled_at")) + ")")
    else:
        cfg = json.load(open(CFG_PATH, encoding="utf-8")) if os.path.exists(CFG_PATH) else {}
        print("Pulling from Supabase…")
        snap = pull(cfg)
        json.dump(snap, open(snap_path, "w", encoding="utf-8"))
    if not os.path.exists(a.data):
        sys.exit("No " + a.data + " -- run the app build first.")
    model = Model(json.load(open(a.data, encoding="utf-8")))
    events, notes, links, dv, officers = analyse(model, snap)
    act, ring = active_set(model, links)
    meta = {"pulled": fmt_when(parse_ts(snap.get("pulled_at"))), "by": s_(snap.get("by")),
            "n_notes": len(snap.get("notes", [])), "n_stops": len(snap.get("checks", [])), "n_queries": len(snap.get("queries", []))}
    xl = os.path.join(a.out, "PVH_Intel.xlsx")
    try:
        write_xlsx(xl, model, events, notes, links, dv, officers, act, meta)
    except PermissionError:
        sys.exit("PVH_Intel.xlsx is open in Excel -- close it and run again.")
    write_gephi(a.out, model, links, ring)
    write_map(os.path.join(a.out, "PVH_Intel_Map.html"), model, links, notes, ring, meta)
    nof = sum(1 for _, _, L in dv if not L["on_file"])
    print("Records with activity: %d | links: %d | driver-vehicle: %d (%d NOT ON FILE)" % (len(act), len(links), len(dv), nof))
    print("Written to " + a.out)


if __name__ == "__main__":
    main()
