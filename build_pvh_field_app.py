#!/usr/bin/env python3
"""
PVH Field App builder.

Reads the two PVH exports and produces one self-contained offline HTML app:
searchable by deck number, plate number and names, with per-record detail
screens and operator-to-vehicle linkage by owner name.

Usage:
    python build_pvh_field_app.py [vehicles.xlsx] [operators.xlsx] [addresses.xlsx] [output.html]

The optional addresses file is the "All_Vehicle_LastInspection" export; when
supplied, owner mailing addresses are merged into vehicle records by VIN
(fallback: deck + plate).

Defaults:
    vehicles : All_Active_Vehicles.xlsx
    operators: PVH_ALL_Operators.xlsx
    output   : PVH_Field_App.html

No payment, refund, SAP or criminal-check fields are embedded.
"""
import re
import sys
import json
import hashlib
import datetime
from collections import Counter
import pandas as pd

VEH_FILE = sys.argv[1] if len(sys.argv) > 1 else "All_Active_Vehicles.xlsx"
OP_FILE = sys.argv[2] if len(sys.argv) > 2 else "PVH_ALL_Operators.xlsx"
ADDR_FILE = None
OUT_FILE = "PVH_Field_App.html"
_rest = sys.argv[3:]
for a in _rest:
    if a.lower().endswith(".html"):
        OUT_FILE = a
    else:
        ADDR_FILE = a

# Fields embedded in the app. Everything else in the exports is dropped.
VEH_FIELDS = [
    "Vehicle Type", "Business Name", "Owner Last Name", "Owner First Name",
    "Owner ID", "VIN", "Make Model", "Vehicle Color", "v Year", "Deck No",
    "Plate No", "NSVehicle Permit Expiry", "District", "First MVIDate",
    "Insurance Expiry", "Notes", "Vehicle ID", "Inspection Date",
    "Licence No", "Expiry Date",
]
OP_FIELDS = [
    "Operator Type", "ID", "Business Name", "Last Name", "First Name",
    "Middle", "Address1", "Address2", "City", "Province", "Postal Code",
    "Phone", "Cell Phone", "Master Number", "Licence Number", "v Year",
    "NSDLExpired Date", "Approval Date", "Renewal Date", "District",
    "In Active", "Cancelled", "Notes",
]


def find_header_row(path, marker):
    """Locate the real header row in a report-style export."""
    probe = pd.read_excel(path, header=None, nrows=12)
    for i in range(len(probe)):
        if str(probe.iloc[i, 0]).strip() == marker:
            return i
    hint = ""
    if marker == "Vehicle Type":
        probe2 = pd.read_excel(path, header=None, nrows=8)
        flat = " ".join(str(x) for x in probe2.values.flatten())
        if "Owner Type" in flat:
            hint = ("\nThis looks like the owner-keyed 'Active Owners & Active Vehicles' report, "
                    "which lacks insurance, MVI and licence expiry fields.\n"
                    "Export the 'All Active Vehicles' report instead.")
    raise SystemExit(f"Could not find header row (looked for '{marker}' in column A) in {path}{hint}")


def clean_value(v):
    if pd.isna(v):
        return None
    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)):
        return v.strftime("%m/%d/%Y")
    if isinstance(v, bool):
        return v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str):
        return v.strip()
    return v


# Which source column feeds which on-screen status row. The two permit
# columns are easy to transpose, so state it once here:
#
#   "Expiry Date"              -> PVH License  (the CBRM taxi licence)
#   "NSVehicle Permit Expiry"  -> NS Permit    (the provincial vehicle permit)
#
# The column names mean exactly what they say. An earlier version of this
# build had the two swapped, which showed vehicles as PVH-expired in the
# field when only their NS permit had lapsed. Verified 2026-08-24 against
# deck 87 (licence TO4351): Expiry Date 01/31/2027, NSVehicle Permit
# Expiry 07/31/2026 -- its PVH licence is current. Do not transpose these.
VEH_DATES = ["NSVehicle Permit Expiry", "First MVIDate", "Insurance Expiry",
             "Inspection Date", "Expiry Date"]
OP_DATES = ["NSDLExpired Date", "Approval Date", "Renewal Date"]

DATE_WARNINGS = []
DEDUP_LOG = []
OWNER_NAME_DRIFT = []

def load(path, marker, fields, date_cols):
    hdr = find_header_row(path, marker)
    df = pd.read_excel(path, header=hdr)
    missing = [f for f in fields if f not in df.columns]
    if missing:
        raise SystemExit(f"{path}: expected columns missing: {missing}")
    df = df[fields]
    df = df.dropna(how="all")
    DATE_FORMATS = ["%m/%d/%Y", "%m/%d/%y", "%Y/%m/%d", "%Y-%m-%d", "%m%d%Y", "%d/%m/%Y"]

    def parse_any(x):
        if pd.isna(x):
            return pd.NaT
        if isinstance(x, (pd.Timestamp, datetime.datetime, datetime.date)):
            return pd.Timestamp(x)
        s = str(x).strip()
        for fmt in DATE_FORMATS:
            try:
                return pd.Timestamp(datetime.datetime.strptime(s, fmt))
            except ValueError:
                continue
        return pd.NaT

    for c in date_cols:
        raw = df[c].copy()
        df[c] = raw.map(parse_any)
        lost = raw.notna() & df[c].isna()
        for idx in df.index[lost]:
            DATE_WARNINGS.append(f"{path.split('/')[-1]} \u00B7 {c}: unparseable value '{raw[idx]}' (row {idx+hdr+2}) \u2014 treated as none on file")
    records = []
    for _, row in df.iterrows():
        records.append({k: clean_value(v) for k, v in row.items()})
    return records


def norm_part(s):
    n = re.sub(r"[^A-Z]", "", (s or "").upper())
    if n.startswith("MC") and not n.startswith("MAC"):
        n = "MAC" + n[2:]
    return n


def norm_name(last, first):
    l, f = norm_part(last), norm_part(first)
    return (l, f) if (l or f) else None


def norm_master(m):
    """Normalize a Master Number for identity matching.

    Source files carry inconsistent prefixes and spacing on the same operator:
    '4A  MACGI...', '4A MACGI...', '04 KUKKA...', '4 SINGH...'. Strip a leading
    4A / 04 / 4 prefix, collapse all whitespace, uppercase. Returns None for
    blank so blank masters are never collapsed together.
    """
    if m is None:
        return None
    s = str(m).strip().upper()
    if not s:
        return None
    s = re.sub(r"^4A\b", "", s)     # strip leading 4A
    s = re.sub(r"^0*4\s+", "", s)   # strip leading "4"/"04" prefix token (only
                                     # when followed by whitespace, so a plain
                                     # numeric master number like "4444444" is
                                     # left untouched instead of losing a digit)
    s = re.sub(r"\s+", "", s)       # collapse all internal whitespace
    return s or None


def _op_renewal_key(rec):
    """Sort key for choosing the surviving row in an operator dup cluster:
    latest Renewal Date wins; tiebreak on highest TD licence number."""
    r = _pdate(rec.get("Renewal Date"))
    rd = r or datetime.date.min
    m = re.search(r"(\d+)", str(rec.get("Licence Number") or ""))
    td = int(m.group(1)) if m else -1
    return (rd, td)


def dedupe_operators(operators):
    """Collapse operator rows sharing a normalized Master Number to one row.
    Newest Renewal Date wins (tiebreak: highest TD). Dropped rows are logged.
    Rows with a blank Master Number are never merged and always kept."""
    groups = {}
    passthrough = []
    for op in operators:
        key = norm_master(op.get("Master Number"))
        if key is None:
            passthrough.append(op)
        else:
            groups.setdefault(key, []).append(op)
    kept = []
    for key, rows in groups.items():
        if len(rows) == 1:
            kept.append(rows[0])
            continue
        rows_sorted = sorted(rows, key=_op_renewal_key, reverse=True)
        winner = rows_sorted[0]
        kept.append(winner)
        for loser in rows_sorted[1:]:
            DEDUP_LOG.append(
                "operator dup [master %s]: kept %s (ID %s, renewed %s) \u2014 dropped %s (ID %s, renewed %s)"
                % (key, winner.get("Licence Number"), winner.get("ID"), winner.get("Renewal Date"),
                   loser.get("Licence Number"), loser.get("ID"), loser.get("Renewal Date")))
    result = kept + passthrough
    return result


def build_owners(vehicles, op_by_name):
    """Group vehicles into first-class Owner records.

    Grouped primarily by Owner ID -- a fully-populated identity key on
    every vehicle row, more reliable than the name-string matching used
    elsewhere in this file as a display-only fallback. BUT Owner ID is not
    always a clean 1:1 key: the source data has at least one confirmed
    case of the same Owner ID reused across two unrelated people (Owner ID
    1: "Colemon, Thomas", one Limo vehicle, vs "MacGillivary, Michael",
    five Tour vehicles). A same-ID group is therefore split further by
    normalized owner name so unrelated people never get merged onto one
    Owner card; a split logs a warning, since a reused Owner ID is a real
    source-data problem worth flagging, not cosmetic spelling drift.

    Owner-to-Operator linkage reuses the SAME exact-match name index the
    rest of the app already treats as reliable (op_by_name) -- it
    deliberately does NOT use the fuzzy tier, so an "Owner-Operator" label
    is never a guess. Mutates vehicles in place, setting v["_owner"].
    """
    id_groups = {}
    for i, v in enumerate(vehicles):
        oid = v.get("Owner ID")
        if oid in (None, ""):
            continue
        id_groups.setdefault(oid, []).append(i)

    def most_common(rows, field):
        vals = [r.get(field) for r in rows if r.get(field) not in (None, "")]
        return Counter(vals).most_common(1)[0][0] if vals else None

    owners = []
    for oid, idxs in id_groups.items():
        by_identity = {}
        for i in idxs:
            v = vehicles[i]
            key = norm_name(v.get("Owner Last Name"), v.get("Owner First Name")) or ("", str(i))
            by_identity.setdefault(key, []).append(i)
        if len(by_identity) > 1:
            names = "; ".join(
                f"{vehicles[sub[0]].get('Owner Last Name')}, {vehicles[sub[0]].get('Owner First Name')}"
                for sub in by_identity.values()
            )
            OWNER_NAME_DRIFT.append(
                f"Owner ID {oid} reused across different identities (kept separate): {names}"
            )
        for key, sub_idxs in by_identity.items():
            rows = [vehicles[i] for i in sub_idxs]
            last = most_common(rows, "Owner Last Name")
            first = most_common(rows, "Owner First Name")
            nkey = norm_name(last, first)
            owners.append({
                "Owner ID": oid,
                "Owner Last Name": last,
                "Owner First Name": first,
                "Business Name": most_common(rows, "Business Name"),
                "Owner Address": most_common(rows, "Owner Address"),
                "_veh": sub_idxs,
                "_op": op_by_name.get(nkey) if nkey else None,
            })
    owners.sort(key=lambda o: (o["Owner Last Name"] or "", o["Owner First Name"] or ""))
    for i, o in enumerate(owners):
        o["_i"] = i
        for ix in o["_veh"]:
            vehicles[ix]["_owner"] = i
    return owners


def lev1(a, b):
    """True if edit distance between a and b is <= 1."""
    if a == b:
        return True
    la, lb = len(a), len(b)
    if abs(la - lb) > 1:
        return False
    i = j = edits = 0
    while i < la and j < lb:
        if a[i] == b[j]:
            i += 1; j += 1; continue
        edits += 1
        if edits > 1:
            return False
        if la == lb:
            i += 1; j += 1
        elif la > lb:
            i += 1
        else:
            j += 1
    return edits + (la - i) + (lb - j) <= 1


def main():
    vehicles = load(VEH_FILE, "Vehicle Type", VEH_FIELDS, VEH_DATES)
    operators = load(OP_FILE, "Operator Type", OP_FIELDS, OP_DATES)
    # Collapse duplicate operator rows (same person across licence renewals /
    # stale LD records) BEFORE building name-link index arrays, so positional
    # indices stay valid. Vehicles are intentionally NOT deduped: Vehicle ID is
    # reused across different physical vehicles in the source, so collapsing on
    # it would silently drop a real vehicle. Collisions are reported instead.
    operators = dedupe_operators(operators)
    merge_addresses(vehicles)

    # Link operators to vehicles by normalized owner name; fuzzy tier for near-misses.
    veh_by_name = {}
    for i, v in enumerate(vehicles):
        key = norm_name(v["Owner Last Name"], v["Owner First Name"])
        if key:
            veh_by_name.setdefault(key, []).append(i)
    owner_keys = list(veh_by_name.keys())
    op_by_name = {}
    for j, op in enumerate(operators):
        k = norm_name(op["Last Name"], op["First Name"])
        if k and k not in op_by_name:
            op_by_name[k] = j
    for v in vehicles:
        k = norm_name(v["Owner Last Name"], v["Owner First Name"])
        v["_op"] = op_by_name.get(k) if k else None
    for op in operators:
        key = norm_name(op["Last Name"], op["First Name"])
        op["_veh"] = veh_by_name.get(key, []) if key else []
        fz = []
        if key and len(key[0]) >= 3:
            for ok in owner_keys:
                if ok == key or len(ok[0]) < 3:
                    continue
                if lev1(key[0], ok[0]) and lev1(key[1], ok[1]):
                    fz.extend(veh_by_name[ok])
        op["_vfz"] = [i for i in fz if i not in op["_veh"]][:6]

    owners = build_owners(vehicles, op_by_name)

    quality_report(vehicles, operators, owners)
    diff_report(OUT_FILE, vehicles, operators)

    data = {
        "built": datetime.datetime.now().strftime("%b %d, %Y %H:%M"),
        "sources": {
            "vehicles": VEH_FILE.split("/")[-1],
            "operators": OP_FILE.split("/")[-1],
            "owners": "derived from Owner ID in " + VEH_FILE.split("/")[-1],
        },
        "counts": {"vehicles": len(vehicles), "operators": len(operators), "owners": len(owners)},
        "vehicles": vehicles,
        "operators": operators,
        "owners": owners,
    }
    payload_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    # Escape "</" so the JSON can never terminate the script tag.
    payload = payload_json.replace("</", "<\\/")

    html = (TEMPLATE
            .replace("__HEAD_EXTRA__", "")
            .replace("__DATA_SCRIPT__", '<script id="data" type="application/json">' + payload + "</script>")
            .replace("__BOOT__", SINGLE_BOOT)
            .replace("__APPVER__", shell_version()[0]))
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)
    kb = len(html.encode("utf-8")) // 1024
    addr_src = ADDR_FILE.split("/")[-1] if ADDR_FILE else "none (owner addresses skipped)"
    print(f"Built {OUT_FILE} ({kb} KB) | vehicles: {len(vehicles)} | operators: {len(operators)} | owners: {len(owners)}")
    print(f"Sources: {VEH_FILE.split('/')[-1]} + {OP_FILE.split('/')[-1]} + {addr_src}")
    write_shell(payload_json)


def _pdate(s):
    try:
        return datetime.datetime.strptime(s, "%m/%d/%Y").date() if s else None
    except (ValueError, TypeError):
        return None


def _expired(s):
    d = _pdate(s)
    return d is not None and d < datetime.date.today()


def _mvi_due(s):
    d = _pdate(s)
    if not d:
        return None
    try:
        return d.replace(year=d.year + 1)
    except ValueError:
        return d.replace(year=d.year + 1, day=28)


def merge_addresses(vehicles):
    for v in vehicles:
        v["Owner Address"] = None
    if not ADDR_FILE:
        return
    hdr = find_header_row(ADDR_FILE, "Vehicle Type")
    df = pd.read_excel(ADDR_FILE, header=hdr)
    need = ["VIN", "Deck No", "Plate No", "Owner Address1", "Owner Address2",
            "City", "Province", "Postal Code"]
    missing = [c for c in need if c not in df.columns]
    if missing:
        print(f"Address file {ADDR_FILE}: columns missing {missing}; addresses skipped")
        return
    by_vin, by_dp = {}, {}
    for _, r in df.iterrows():
        parts = [r["Owner Address1"], r["Owner Address2"],
                 ", ".join(str(x) for x in [r["City"], r["Province"], r["Postal Code"]] if pd.notna(x))]
        addr = " \u00B7 ".join(str(p).strip() for p in parts if pd.notna(p) and str(p).strip()) or None
        if addr is None:
            continue
        vin = str(r["VIN"]).strip().upper() if pd.notna(r["VIN"]) else None
        if vin and vin not in by_vin:
            by_vin[vin] = addr
        dp = (str(r["Deck No"]).strip(), str(r["Plate No"]).strip())
        if dp not in by_dp:
            by_dp[dp] = addr
    matched = 0
    for v in vehicles:
        vin = str(v["VIN"]).strip().upper() if v["VIN"] else None
        addr = by_vin.get(vin) if vin else None
        if addr is None:
            addr = by_dp.get((str(v["Deck No"]), str(v["Plate No"])))
        if addr:
            v["Owner Address"] = addr
            matched += 1
    print(f"Owner addresses merged: {matched}/{len(vehicles)} vehicles"
          + ("" if matched == len(vehicles) else " (rest have no address on file)"))


def quality_report(vehicles, operators, owners=None):
    warn = list(DATE_WARNINGS) + list(OWNER_NAME_DRIFT)
    decks, plates = {}, {}
    for v in vehicles:
        d, p = v["Deck No"], v["Plate No"]
        if d not in (None, ""):
            decks.setdefault(str(d), []).append(p or "?")
        if p:
            plates[p] = plates.get(p, 0) + 1
    dup_d = {k: pl for k, pl in decks.items() if len(pl) > 1}
    if dup_d:
        warn.append("Duplicate deck numbers: " + ", ".join(
            f"{k} ({'/'.join(pl)})" for k, pl in sorted(dup_d.items())))
    dup_p = [p for p, c in plates.items() if c > 1]
    if dup_p:
        warn.append("Duplicate plates: " + ", ".join(sorted(dup_p)))
    # Vehicle ID collisions: same ID on rows with different VINs = source-system
    # ID reuse across different physical vehicles. Both rows are KEPT (not deduped);
    # this only surfaces the collision so it can be reconciled at source.
    vid = {}
    for v in vehicles:
        k = v.get("Vehicle ID")
        if k not in (None, ""):
            vid.setdefault(k, []).append(v)
    for k, rows in vid.items():
        vins = {str(r.get("VIN") or "").upper() for r in rows}
        if len(rows) > 1 and len(vins) > 1:
            desc = " vs ".join(
                f"{r.get('Plate No') or '?'}/{(r.get('Owner Last Name') or r.get('Business Name') or '?')}"
                for r in rows)
            warn.append(f"Vehicle ID {k} reused across different vehicles (both kept): {desc}")
    for field, label in [("Plate No", "plate"), ("VIN", "VIN"),
                         ("First MVIDate", "MVI date"),
                         ("Insurance Expiry", "insurance expiry"),
                         ("Expiry Date", "PVH licence expiry"),
                         ("NSVehicle Permit Expiry", "NS permit expiry")]:
        m = [v for v in vehicles if v[field] in (None, "")]
        if m:
            ids = ", ".join(str(x["Deck No"] or x["Plate No"] or x["VIN"] or "?") for x in m[:10])
            warn.append(f"Vehicles missing {label}: {len(m)}" + (f" (decks/plates: {ids})" if len(m) <= 10 else ""))
    active = [o for o in operators if o.get("In Active") is not True and o.get("Cancelled") is not True]
    for field, label in [("Renewal Date", "renewal date"), ("NSDLExpired Date", "NS DL expiry")]:
        m = [o for o in active if o[field] in (None, "")]
        if m:
            warn.append(f"Active operators missing {label}: {len(m)}")
    if warn:
        print("\n=== DATA QUALITY WARNINGS ===")
        for w in warn:
            print(" ! " + w)
    else:
        print("Data quality: no issues found")
    if DEDUP_LOG:
        print(f"\n=== OPERATOR DEDUP ({len(DEDUP_LOG)} row(s) dropped) ===")
        for d in DEDUP_LOG:
            print(" - " + d)


def _veh_key(v):
    return str(v.get("Vehicle ID") or v.get("VIN") or f"{v.get('Deck No')}|{v.get('Plate No')}")


def _op_key(o):
    return f"{o.get('Operator Type')}|{o.get('ID')}|{norm_part(o.get('Last Name'))}"


def _veh_flags(v):
    f = set()
    due = _mvi_due(v["First MVIDate"])
    if due and due < datetime.date.today():
        f.add("MVI")
    for field, label in [("Expiry Date", "PVH License"),
                         ("Insurance Expiry", "Insurance"),
                         ("NSVehicle Permit Expiry", "NS Permit")]:
        if _expired(v[field]):
            f.add(label)
    return f


def diff_report(out_path, vehicles, operators):
    try:
        old_html = open(out_path, encoding="utf-8").read()
    except FileNotFoundError:
        print("No previous build found; skipping diff")
        return
    m = re.search(r'<script id="data" type="application/json">(.*?)</script>', old_html, re.S)
    if not m:
        return
    old = json.loads(m.group(1).replace("<\\/", "</"))
    ov = {_veh_key(v): v for v in old["vehicles"]}
    nv = {_veh_key(v): v for v in vehicles}
    oo = {_op_key(o): o for o in old["operators"]}
    no = {_op_key(o): o for o in operators}
    lines = []
    for k in nv:
        if k not in ov:
            v = nv[k]
            lines.append(f"+ vehicle: deck {v['Deck No']} plate {v['Plate No']} ({v['Owner Last Name']})")
    for k in ov:
        if k not in nv:
            v = ov[k]
            lines.append(f"- vehicle removed: deck {v['Deck No']} plate {v['Plate No']} ({v['Owner Last Name']})")
    for k in no:
        if k not in oo:
            o = no[k]
            lines.append(f"+ operator: {o['Last Name']}, {o['First Name']} ({o['Operator Type']})")
    for k in oo:
        if k not in no:
            o = oo[k]
            lines.append(f"- operator removed: {o['Last Name']}, {o['First Name']} ({o['Operator Type']})")
    for k in nv:
        if k in ov:
            newly = _veh_flags(nv[k]) - _veh_flags(ov[k])
            if newly:
                v = nv[k]
                lines.append(f"! newly expired ({', '.join(sorted(newly))}): deck {v['Deck No']} plate {v['Plate No']}")
    if lines:
        print(f"\n=== CHANGES SINCE LAST BUILD ({old.get('built')}) ===")
        for ln in lines:
            print(" " + ln)
    else:
        print(f"No changes since last build ({old.get('built')})")


SINGLE_BOOT = "boot(JSON.parse(document.getElementById(\"data\").textContent));"

SHELL_BOOT = r"""window.__SHELL__=true;
if("serviceWorker" in navigator){navigator.serviceWorker.register("sw.js").catch(function(){});}
(function(){
var DKEY="pvh_data", UKEY="pvh_data_url", CKEY="pvh_last_check", CHECK_MS=600000, JKEY="pvh_just_updated";
/* Flips true on the first tap or keystroke anywhere in the app — attached
   at the document level so it doesn't depend on what boot() wires up, and
   fires for a card tap, a settings link, typing in search, all of it.
   A cold-start auto-update is only safe to apply without asking while this
   is still false — see check("cold"). */
var interacted=false;
function markInteracted(){interacted=true;}
document.addEventListener("pointerdown",markInteracted,{once:true,passive:true});
document.addEventListener("keydown",markInteracted,{once:true});
function lsGet(k){try{return localStorage.getItem(k);}catch(e){return null;}}
function lsSet(k,v){try{localStorage.setItem(k,v);}catch(e){}}
function lsDel(k){try{localStorage.removeItem(k);}catch(e){}}
function eh(s){return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;")
  .replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
function mainEl(){return document.getElementById("main");}

/* A Dropbox share link opens a preview page, not the file. The download host
   serves the bytes and allows other sites to read them; the rlkey parameter in
   newer links is part of the credential and has to survive. Any other link is
   used exactly as pasted, so a direct URL from any host still works. */
function directUrl(u){
  u=(u||"").trim();
  if(!u)return "";
  if(/^https?:\/\/(www\.)?dropbox\.com\//i.test(u)){
    var x=u.replace(/^https?:\/\/(www\.)?dropbox\.com/i,"https://dl.dropboxusercontent.com");
    if(/[?&]dl=0(&|$)/i.test(x))x=x.replace(/([?&])dl=0(&|$)/i,"$1dl=1$2");
    else if(!/[?&]dl=1(&|$)/i.test(x))x+=(x.indexOf("?")>-1?"&":"?")+"dl=1";
    return x;
  }
  return u;
}
function fetchData(u){
  return fetch(directUrl(u),{cache:"no-store"}).catch(function(){
    /* fetch only rejects outright for network-level failures, and for a link
       that resolves this almost always means the host refused to let this app
       read it (CORS), which is what a host that blocks other sites does. */
    throw new Error("the host would not let this app read the file (CORS), or the link is unreachable");
  }).then(function(r){
    if(!r.ok)throw new Error("HTTP "+r.status+(r.status===404?" - nothing at that link":
      (r.status===401||r.status===403?" - the link needs a sign-in this app cannot do":"")));
    return r.text();
  }).then(function(t){
    var d;
    try{d=JSON.parse(t);}
    catch(e){throw new Error("that link gave back a web page, not PVH_data.json - it has to be the direct link to the file itself");}
    if(!d||!d.vehicles||!d.operators)throw new Error("that file is not a PVH data file");
    return {text:t,data:d};
  });
}
function toast(html,variant){
  var el=document.getElementById("toast");
  el.className=variant==="ok"?"ok":"";
  el.innerHTML=html;el.style.display="block";
}
var CLOSE=' · <span class="link" onclick="window.__closeToast()">Close</span>';
var PENDING=null;
function builtNow(){try{return (JSON.parse(lsGet(DKEY)||"{}")).built||null;}catch(e){return null;}}
function stampCheck(){lsSet(CKEY,new Date().toISOString());}
/* Shared apply path for every route that swaps in new data: sets the data,
   drops a one-shot marker with the build stamp so the NEXT load can show a
   confirmation, then reloads. A reload is required either way (search index
   and header listeners are rebuilt in boot(), which is not safe to re-run
   into a live DOM) so this never tries to hot-swap in place. */
function applyDataAndReload(res){
  lsSet(DKEY,res.text);lsSet(JKEY,res.data.built||"");stampCheck();location.reload();
}
function check(mode){
  var manual=mode==="manual";
  var u=lsGet(UKEY);
  if(!u){
    if(manual)toast('No data source saved yet. <span class="link" onclick="window.__datasrc()">Set one up</span>'+CLOSE);
    return;
  }
  if(manual)toast("Checking the data source…");
  fetchData(u).then(function(res){
    stampCheck();
    var cur=builtNow();
    if(res.data.built&&cur&&res.data.built===cur){
      if(manual)toast("Already up to date (data built "+eh(cur)+")."+CLOSE);
      return;
    }
    /* Cold start, nothing tapped yet: apply straight away, no prompt — there
       is no in-progress screen to lose. If the officer has already started
       using the app by the time this resolves (dueForCheck ran, but the
       fetch was slow), fall through to the same ask-first banner as a
       mid-session recheck instead of reloading out from under them. */
    if(mode==="cold"&&!interacted){applyDataAndReload(res);return;}
    PENDING=res;
    toast("New data available (built "+eh(res.data.built||"?")+"). "+
      '<span class="link" onclick="window.__applyUpdate()">Update now</span> · '+
      '<span class="link" onclick="window.__closeToast()">Later</span>');
  }).catch(function(e){
    if(manual)toast("Could not read the data source: "+eh(e.message)+". "+
      '<span class="link" onclick="window.__datasrc()">Check the link</span>'+CLOSE);
  });
}
window.__closeToast=function(){var el=document.getElementById("toast");if(el)el.style.display="none";};
window.__checkNow=function(){check("manual");};
window.__applyUpdate=function(){
  if(!PENDING){window.__closeToast();return;}
  applyDataAndReload(PENDING);
};
window.__syncLine=function(){
  var u=lsGet(UKEY),c=lsGet(CKEY),s="";
  if(!u)return "No automatic data source set — data changes only when you load a file.<br>";
  if(c){try{s=" · last checked "+new Date(c).toLocaleString();}catch(e){}}
  return "Auto-update on"+s+"<br>";
};
function srcHost(u){
  var m=String(u).match(/^https?:\/\/([^\/?#]+)/i);
  return m?m[1].replace(/^www\./i,""):"saved link";
}
function srcCode(u){
  /* Short, non-secret fingerprint of the link. Lets two devices be compared
     ("does yours show A3F9 as well?") without either screen ever putting the
     URL itself in front of whoever is standing there. */
  var h=5381;
  for(var i=0;i<u.length;i++)h=((h<<5)+h+u.charCodeAt(i))>>>0;
  return ("000"+h.toString(36).toUpperCase()).slice(-4);
}
window.__datasrc=function(replace){
  var cur=lsGet(UKEY)||"";
  /* A saved link is never rendered back to the screen. This page was the one
     place the URL sat in the clear, so anyone holding the handset could read
     the whole dataset's share link off it. Replacing therefore means pasting a
     fresh link rather than editing the old one -- the link's source of truth is
     the build machine's config, not the phone. */
  var editing=!cur||replace===true;
  mainEl().innerHTML='<div class="seclabel">Automatic data source</div>'+
    (editing
      ? '<div class="notes">Put PVH_data.json in cloud storage, copy its share link and paste it below. '+
        'Every time the app opens it checks that link and offers the file whenever the build stamp changes. '+
        'The link and the data are kept on this device only.\n\n'+
        'In Dropbox use Share → Copy link and paste the whole thing exactly as copied — this screen turns it '+
        'into a direct file link for you. A direct link from any other host works too.\n\n'+
        'Keep overwriting that same file each build. Deleting it and uploading a new one gives it a new link, '+
        'which quietly stops every device that was set up with the old one.</div>'+
        '<input id="dsurl" class="dsinput" type="url" inputmode="url" autocomplete="off" spellcheck="false" '+
          'placeholder="https://… link to PVH_data.json" value="">'+
        '<button class="copybtn" id="dssave">Save and check now</button>'
      : '<div class="notes">This device already has a data source. The link is not shown here — '+
        'if it needs to change, paste a fresh one and it replaces the old.</div>'+
        '<div class="srccard">Source · '+eh(srcHost(cur))+'<br>Link code <span class="code">'+eh(srcCode(cur))+'</span></div>'+
        '<button class="copybtn" id="dsedit">Replace this link</button>')+
    (cur?'<button class="copybtn" id="dsclear" style="color:var(--bad)">Remove this link</button>':'')+
    '<div class="stamp" id="dsmsg"></div>'+
    '<div class="hint"><span class="link" onclick="window.__back()">Back</span> · '+
      '<span class="link" onclick="window.__reimport()">Load a file instead…</span></div>';
  var msg=document.getElementById("dsmsg");
  var sv=document.getElementById("dssave");
  if(sv)sv.addEventListener("click",function(){
    var u=document.getElementById("dsurl").value.trim();
    if(!u){msg.textContent="Paste a link first.";return;}
    msg.textContent="Checking…";
    fetchData(u).then(function(res){
      lsSet(UKEY,u);lsSet(DKEY,res.text);stampCheck();
      msg.textContent="Saved. Loading data built "+(res.data.built||"?")+"…";
      setTimeout(function(){location.reload();},600);
    }).catch(function(e){
      msg.textContent="Could not read that link: "+e.message+".";
    });
  });
  var eb=document.getElementById("dsedit");
  if(eb)eb.addEventListener("click",function(){window.__datasrc(true);});
  var cb=document.getElementById("dsclear");
  if(cb)cb.addEventListener("click",function(){lsDel(UKEY);lsDel(CKEY);window.__datasrc();});
};
window.__back=function(){if(window.__route)window.__route();else location.reload();};
function showImport(msg){
  mainEl().innerHTML='<div class="hint" style="padding-top:50px">'+(msg||"No data loaded on this device yet.")+'</div>'+
    '<label class="copybtn" style="text-align:center;display:block">Load PVH_data.json'+
    '<input type="file" accept=".json,application/json" style="display:none" id="datafile"></label>'+
    '<div class="stamp">Pick the PVH_data.json produced by the build script (Dropbox / Files).</div>'+
    '<div class="hint"><span class="link" onclick="window.__datasrc()">Or set up automatic updates from a Dropbox link…</span></div>';
  document.getElementById("datafile").addEventListener("change",function(ev){
    var f=ev.target.files[0]; if(!f)return;
    var r=new FileReader();
    r.onload=function(){
      try{
        var d=JSON.parse(r.result);
        if(!d.vehicles||!d.operators)throw 0;
        lsSet(DKEY,r.result);
        boot(d);
      }catch(e){showImport("That file is not a valid PVH data file. Try again.");}
    };
    r.readAsText(f);
  });
}
window.__reimport=function(){showImport("Load the new PVH_data.json.");};
function dueForCheck(){
  if(!lsGet(UKEY))return false;
  var c=lsGet(CKEY), t=c?Date.parse(c):NaN;
  return isNaN(t)||(Date.now()-t)>CHECK_MS;
}
(function start(){
  document.body.classList.add("light");
  var saved=lsGet(DKEY), loaded=false;
  if(saved){
    try{boot(JSON.parse(saved));loaded=true;}
    catch(e){lsDel(DKEY);}
  }
  var justUpdated=lsGet(JKEY);
  if(justUpdated){lsDel(JKEY);toast("Data updated — built "+eh(justUpdated)+"."+CLOSE,"ok");}
  if(loaded){
    if(dueForCheck())check("cold");
  }else if(lsGet(UKEY)){
    mainEl().innerHTML='<div class="hint" style="padding-top:50px">Loading data from the saved link…</div>';
    fetchData(lsGet(UKEY)).then(function(res){
      lsSet(DKEY,res.text);stampCheck();boot(res.data);
    }).catch(function(e){
      showImport("Could not load from the saved link ("+eh(e.message)+"). Load the file by hand, or "+
        '<span class="link" onclick="window.__datasrc()">check the link</span>:');
    });
  }else{
    showImport();
  }
  document.addEventListener("visibilitychange",function(){
    if(document.visibilityState==="visible"&&dueForCheck())check("mid");
  });
})();
})();"""

SHELL_HEAD = """<link rel="manifest" href="manifest.webmanifest">
<link rel="apple-touch-icon" href="icon-192.png">
<meta name="theme-color" content="#0A62C6">"""

MANIFEST = """{"name":"PVH Field Lookup","short_name":"PVH","start_url":"./","scope":"./","display":"standalone","background_color":"#EFF2F6","theme_color":"#0A62C6","icons":[{"src":"icon-192.png","sizes":"192x192","type":"image/png"},{"src":"icon-512.png","sizes":"512x512","type":"image/png"}]}"""

SW_JS = r"""const C="pvh-shell-__BUILD__";
self.addEventListener("install",e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(["./","index.html","manifest.webmanifest","icon-192.png","icon-512.png"])).then(()=>self.skipWaiting()))});
self.addEventListener("activate",e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==C).map(x=>caches.delete(x)))).then(()=>self.clients.claim()))});
self.addEventListener("fetch",e=>{
  if(e.request.method!=="GET")return;
  const u=new URL(e.request.url);
  /* Only the app shell is cached. Data fetches (Dropbox and friends) always go
     to the network, so a new build is actually seen and no record data is left
     behind in the cache. */
  if(u.origin!==self.location.origin)return;
  if(/\.json(\?|$)/i.test(u.pathname+u.search))return;
  /* The page itself is network-first: a redeployed app is picked up on the next
     launch that has a signal, instead of the cached copy being served forever.
     The cache is the fallback, so offline still works, and a 3s cap means a
     flaky connection in the field falls back rather than hanging. */
  const isPage=e.request.mode==="navigate"||/(^|\/)(index\.html)?$/.test(u.pathname);
  if(isPage){
    e.respondWith(Promise.race([
      fetch(e.request).then(res=>{
        const cl=res.clone();
        caches.open(C).then(c=>c.put("index.html",cl));
        return res;
      }),
      new Promise(r=>setTimeout(()=>r(null),3000))
    ]).then(res=>res||caches.match("index.html",{ignoreSearch:true}).then(c=>c||fetch(e.request)))
      .catch(()=>caches.match("index.html",{ignoreSearch:true})));
    return;
  }
  e.respondWith(caches.match(e.request,{ignoreSearch:true}).then(r=>r||fetch(e.request).then(res=>{const cl=res.clone();caches.open(C).then(c=>c.put(e.request,cl));return res;}).catch(()=>caches.match("index.html"))));
});"""


def make_icons(folder):
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        print("Pillow not installed; skipping icon generation (pip install pillow)")
        return False
    for px in (192, 512):
        img = Image.new("RGBA", (px, px), (18, 22, 27, 255))
        d = ImageDraw.Draw(img)
        m = px * 0.16
        d.ellipse([m, m, px - m, px - m], outline=(247, 190, 74, 255), width=max(6, px // 14))
        w, h = px * 0.44, px * 0.15
        d.rounded_rectangle([(px - w) / 2, (px - h) / 2, (px + w) / 2, (px + h) / 2],
                            radius=h * 0.25, fill=(239, 242, 246, 255))
        img.save(f"{folder}/icon-{px}.png")
    return True


VERSION_FILE = "app_version.json"
_SHELL_VER = None


def shell_version():
    """Give the app a version worth reading, that still moves only on a real change.

    Returns (label, digest): the label is what an officer sees and can say out
    loud ("v3, shipped the 23rd"); the digest keys the service-worker cache.

    The shell's own code is hashed, and app_version.json remembers which hash
    the current number was issued for. An unchanged rebuild reuses the recorded
    number and date verbatim, so the 07:30 run stays byte-identical and nobody
    is asked to re-download a shell they already have. A real edit rolls the
    number forward and stamps the day it shipped.

    Hashed before __APPVER__ is substituted, so the value cannot depend on
    itself. Records are excluded on purpose: this versions the app, not the data
    it loads -- those are separate lines on screen because they answer separate
    questions.

    The ledger is committed deliberately. Without it a fresh clone would restart
    at v1 and start reissuing numbers that already mean something else on the
    handsets.
    """
    global _SHELL_VER
    if _SHELL_VER is not None:
        return _SHELL_VER
    base = (TEMPLATE
            .replace("__HEAD_EXTRA__", SHELL_HEAD)
            .replace("__DATA_SCRIPT__", "")
            .replace("__BOOT__", SHELL_BOOT))
    digest = hashlib.sha256(base.encode("utf-8")).hexdigest()[:8]

    try:
        with open(VERSION_FILE, encoding="utf-8") as f:
            rec = json.load(f)
    except (OSError, ValueError):
        # Missing or unreadable ledger starts the count rather than failing the
        # build; a stale number is worse than an obviously fresh one.
        rec = {}

    if rec.get("hash") != digest:
        rec = {"version": int(rec.get("version", 0)) + 1,
               "date": datetime.date.today().isoformat(),
               "hash": digest}
        with open(VERSION_FILE, "w", encoding="utf-8") as f:
            json.dump(rec, f, indent=2)
            f.write("\n")

    _SHELL_VER = (f"v{rec['version']} · {rec['date']}", digest)
    return _SHELL_VER


def write_shell(payload_json):
    import os
    os.makedirs("docs", exist_ok=True)
    label, _ = shell_version()
    shell_html = (TEMPLATE
                  .replace("__HEAD_EXTRA__", SHELL_HEAD)
                  .replace("__DATA_SCRIPT__", "")
                  .replace("__BOOT__", SHELL_BOOT)
                  .replace("__APPVER__", label))
    open("docs/index.html", "w", encoding="utf-8").write(shell_html)
    # Key the cache on the bytes actually served, not on the pre-substitution
    # digest: the version label is baked into this file, so a build that changes
    # only the label still needs a new cache key. Deterministic either way --
    # the label is settled before this is computed.
    cache_key = hashlib.sha256(shell_html.encode("utf-8")).hexdigest()[:8]
    open("docs/sw.js", "w", encoding="utf-8").write(SW_JS.replace("__BUILD__", cache_key))
    open("docs/manifest.webmanifest", "w", encoding="utf-8").write(MANIFEST)
    make_icons("docs")
    open("PVH_data.json", "w", encoding="utf-8").write(payload_json)
    print(f"Shell written to docs/ | app version {label} | "
          f"data written to PVH_data.json (do NOT commit)")


TEMPLATE = r"""<!DOCTYPE html>
<html lang="en-CA">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<title>PVH Field Lookup</title>
__HEAD_EXTRA__
<style>
:root{
  --bg:#12161B; --panel:#1C232C; --panel2:#242E39; --line:#37434F;
  --text:#F2F6FA; --dim:#A9B7C6; --faint:#7A8896;
  --accent:#5FAEFF; --ok:#43D384; --warn:#F7BE4A; --bad:#FF6363;
  --taxi:#F7BE4A; --limo:#B99BFF; --tour:#5FAEFF; --shuttle:#43D384;
  --platebg:#0D1115; --plateline:#3A4653;
  --mono:"SF Mono",ui-monospace,Menlo,Consolas,monospace;
}
body.light{
  --bg:#EFF2F6; --panel:#FFFFFF; --panel2:#E7EBF0; --line:#CFD7DF;
  --text:#131C26; --dim:#465666; --faint:#71808F;
  --accent:#0A62C6; --ok:#0C8747; --warn:#9A6A08; --bad:#C42B2B;
  --taxi:#8F6404; --limo:#6236C9; --tour:#0A62C6; --shuttle:#0C8747;
  --platebg:#FFFFFF; --plateline:#8E9CAA;
}
body.light .t-Taxi{background:rgba(143,100,4,.12)}
body.light .t-Limo{background:rgba(98,54,201,.10)}
body.light .t-Tour{background:rgba(10,98,198,.10)}
body.light .t-Shuttle{background:rgba(12,135,71,.12)}
body.light .b-ok{background:rgba(12,135,71,.12)}
body.light .b-warn{background:rgba(154,106,8,.13)}
body.light .b-bad{background:rgba(196,43,43,.11)}
body.light .alertbar{background:rgba(196,43,43,.10);border-color:rgba(196,43,43,.4)}
*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}
html,body{height:100%}
body{background:var(--bg);color:var(--text);
  font:16px/1.45 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;
  overscroll-behavior:none}
#app{max-width:640px;margin:0 auto;min-height:100%;display:flex;flex-direction:column}

/* header */
header{position:sticky;top:0;z-index:10;background:var(--bg);
  border-bottom:1px solid var(--line);padding:10px 12px 8px}
.hrow{display:flex;align-items:center;gap:10px}
.navbtn{flex:0 0 44px;height:44px;border:1px solid var(--line);border-radius:10px;
  background:var(--panel);color:var(--text);font-size:20px;display:flex;
  align-items:center;justify-content:center;cursor:pointer}
.navbtn:active{background:var(--panel2)}
.navbtn[disabled]{opacity:.3;pointer-events:none}
h1{font-size:15px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;flex:1}
h1 small{display:block;font-size:11px;font-weight:400;color:var(--faint);
  letter-spacing:.02em;text-transform:none}
#searchwrap{margin-top:8px;position:relative}
#q{width:100%;height:48px;border:1px solid var(--line);border-radius:12px;
  background:var(--panel);color:var(--text);font-size:17px;padding:0 44px 0 14px;
  outline:none}
#q:focus{border-color:var(--accent)}
#clr{position:absolute;right:4px;top:4px;width:40px;height:40px;border:none;
  background:none;color:var(--faint);font-size:20px;cursor:pointer;display:none}

/* body */
main{flex:1;padding:10px 12px 40px}
.hint{color:var(--faint);font-size:13px;text-align:center;padding:26px 20px}
.counts{color:var(--dim);font-size:12px;text-align:center;padding-top:6px}
.stamp{color:var(--faint);font-size:11px;text-align:center;padding:14px 0 4px}
.seclabel{font-size:11px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;
  color:var(--faint);margin:16px 2px 6px}

/* result cards */
.card{background:var(--panel);border:1px solid var(--line);border-radius:12px;
  padding:13px 12px;margin-bottom:8px;cursor:pointer;display:flex;gap:11px;
  align-items:center;min-height:70px}
.card:active{background:var(--panel2)}
.deck{flex:0 0 52px;height:52px;border-radius:50%;border:2px solid var(--faint);
  color:var(--text);display:flex;flex-direction:column;align-items:center;
  justify-content:center;font-family:var(--mono);font-weight:700;font-size:17px;
  line-height:1}
.deck.taxi{border-color:var(--taxi);color:var(--taxi)}
.deck.limo{border-color:var(--limo);color:var(--limo)}
.deck.tour{border-color:var(--tour);color:var(--tour)}
.deck.shuttle{border-color:var(--shuttle);color:var(--shuttle)}
.deck span{font-size:8px;letter-spacing:.08em;font-weight:600;margin-bottom:2px}
.deck.none{border-style:dashed;color:var(--faint);border-color:var(--line);font-size:11px}
.opdot{flex:0 0 52px;height:52px;border-radius:50%;background:var(--panel2);
  border:1px solid var(--line);display:flex;align-items:center;justify-content:center;
  font-weight:700;font-size:18px;color:var(--dim)}
.cmain{flex:1;min-width:0}
.cname{font-weight:600;font-size:16px;white-space:nowrap;overflow:hidden;
  text-overflow:ellipsis}
.csub{color:var(--dim);font-size:13px;white-space:nowrap;overflow:hidden;
  text-overflow:ellipsis;margin-top:1px}
.crow2{display:flex;gap:6px;margin-top:5px;flex-wrap:wrap}
.plate{font-family:var(--mono);font-weight:700;font-size:13px;background:var(--platebg);
  border:1px solid var(--plateline);border-radius:5px;padding:2px 7px;letter-spacing:.06em}
.chip{font-size:11px;font-weight:700;letter-spacing:.05em;border-radius:5px;
  padding:2px 7px;text-transform:uppercase}
.t-Taxi{background:rgba(245,185,66,.14);color:var(--taxi)}
.t-Limo{background:rgba(176,140,255,.14);color:var(--limo)}
.t-Tour{background:rgba(77,163,255,.14);color:var(--tour)}
.clchk{background:rgba(95,174,255,.16);color:var(--accent)}
.clseen{background:var(--panel2);color:var(--faint)}
.t-Shuttle{background:rgba(63,203,126,.14);color:var(--shuttle)}
.chev{color:var(--faint);font-size:18px}

/* status badges */
.badge{display:inline-flex;align-items:center;gap:6px;border-radius:7px;
  padding:5px 9px;font-size:12px;font-weight:700}
.b-ok{background:rgba(63,203,126,.12);color:var(--ok)}
.b-warn{background:rgba(245,185,66,.14);color:var(--warn)}
.b-bad{background:rgba(255,93,93,.14);color:var(--bad)}
.b-na{background:var(--panel2);color:var(--faint)}
.badge b{font-weight:800}

/* detail */
.dhead{background:var(--panel);border:1px solid var(--line);border-radius:14px;
  padding:16px;margin-bottom:10px}
.dtop{display:flex;gap:14px;align-items:center}
.dhead .deck{flex:0 0 66px;height:66px;font-size:22px}
.dtitle{font-size:20px;font-weight:700;line-height:1.2}
.dsub{color:var(--dim);font-size:14px;margin-top:2px}
.bigplate{font-family:var(--mono);font-weight:800;font-size:22px;background:var(--platebg);
  border:1.5px solid var(--plateline);border-radius:8px;padding:4px 12px;
  letter-spacing:.1em;display:inline-block;margin-top:8px}
.badges{display:flex;flex-wrap:wrap;gap:6px;margin-top:12px}
.alertbar{border-radius:10px;padding:10px 12px;font-weight:700;font-size:14px;
  margin-bottom:10px;background:rgba(255,93,93,.14);color:var(--bad);
  border:1px solid rgba(255,93,93,.35)}
.grid{background:var(--panel);border:1px solid var(--line);border-radius:14px;
  overflow:hidden;margin-bottom:10px}
.grid .row{display:flex;padding:10px 14px;border-bottom:1px solid var(--line);
  gap:12px}
.grid .row:last-child{border-bottom:none}
.grid .k{flex:0 0 44%;color:var(--dim);font-size:13px;padding-top:1px}
.grid .v{flex:1;font-size:15px;font-weight:500;word-break:break-word}
.grid .v.mono{font-family:var(--mono);font-size:14px;letter-spacing:.03em}
.tap{cursor:pointer;border-bottom:1px dashed var(--faint)}
.tap:active{color:var(--accent)}
.copied{color:var(--ok)!important;border-bottom-color:var(--ok)!important}
.link{color:var(--accent);cursor:pointer;text-decoration:none}
.notes{background:var(--panel);border:1px solid var(--line);border-radius:14px;
  padding:12px 14px;margin-bottom:10px;font-size:14px;white-space:pre-wrap;
  color:var(--text);line-height:1.5}
.empty{color:var(--faint);font-size:13px;padding:8px 2px}
a.tel{color:var(--accent);text-decoration:none}
.fchips{display:flex;flex-wrap:wrap;gap:6px;padding:2px 0 10px}
.fchip{flex:0 1 auto;max-width:100%;border:1px solid var(--line);background:var(--panel);
  color:var(--dim);border-radius:20px;padding:7px 13px;font-size:13px;line-height:1.2;
  font-weight:600;cursor:pointer;white-space:normal;text-align:center;min-height:34px;
  display:inline-flex;align-items:center;justify-content:center}
.fchip.on{background:var(--accent);border-color:var(--accent);color:#fff}
.browsebar{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:10px}
.browsebar select{flex:1 1 160px;min-width:0;height:42px;border:1px solid var(--line);
  border-radius:10px;background:var(--panel);color:var(--text);font-size:14px;
  padding:0 32px 0 10px;-webkit-appearance:none;appearance:none;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 12 8'%3E%3Cpath d='M1 2l5 5 5-5' fill='none' stroke='%2371808F' stroke-width='1.6' stroke-linecap='round'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right 11px center;background-size:11px 7px}
.cardbadges{display:flex;flex-wrap:wrap;gap:5px;margin-top:6px}
.cardbadges .badge{padding:3px 7px;font-size:11px}
.copybtn{display:block;width:100%;padding:13px;border-radius:12px;
  border:1px solid var(--line);background:var(--panel);color:var(--accent);
  font-size:15px;font-weight:700;margin-bottom:10px;cursor:pointer}
.copybtn:active{background:var(--panel2)}

/* update banner + data-source form (PWA shell)
   In normal document flow right under the sticky header, not fixed/floating —
   guarantees it can never sit under the header or overlap the search box,
   and it's the first thing visible with no scroll, no position math. */
#toast{margin:10px 16px 0;padding:12px 14px;background:var(--panel);
  border:1px solid var(--accent);border-left:4px solid var(--accent);border-radius:10px;
  font-size:14px;line-height:1.45;color:var(--text);display:none;
  box-shadow:0 4px 14px rgba(0,0,0,.18)}
#toast.ok{border-color:var(--ok);border-left-color:var(--ok)}
#toast .link{font-weight:700;white-space:nowrap}
.dsinput{width:100%;height:46px;border:1px solid var(--line);border-radius:10px;
  background:var(--panel);color:var(--text);font-size:15px;padding:0 12px;
  margin-bottom:10px;outline:none}
.dsinput:focus{border-color:var(--accent)}
/* Stands in for the link input once a source is saved -- confirms a source is
   set and which host it points at, without showing the link itself. */
.srccard{border:1px solid var(--line);border-radius:10px;background:var(--panel);
  color:var(--dim);font-size:14px;line-height:1.5;padding:12px 14px;margin-bottom:10px}
.srccard .code{font-weight:700;color:var(--text);letter-spacing:.08em}

/* wide displays (tablet / desktop browser) */
@media (min-width:720px){
  #app{max-width:880px}
  header{padding:12px 18px 10px}
  main{padding:14px 18px 56px}
  h1{font-size:16px}
  .fchip{font-size:14px;padding:8px 15px}
  .browsebar select{height:44px;font-size:15px}
  .grid .k{flex:0 0 32%;max-width:250px;font-size:14px}
  .card{padding:14px}
}
@media (min-width:1100px){
  #app{max-width:1040px}
}
/* mouse / trackpad affordances */
@media (hover:hover) and (pointer:fine){
  .card:hover{background:var(--panel2)}
  .fchip:hover{border-color:var(--accent);color:var(--text)}
  .fchip.on:hover{color:#fff}
  .navbtn:hover,.copybtn:hover{background:var(--panel2)}
  .browsebar select:hover{border-color:var(--accent)}
}
.browsebar select:focus,.fchip:focus-visible,.copybtn:focus-visible,
.card:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
</style>
</head>
<body>
<div id="app">
  <header>
    <div class="hrow">
      <button class="navbtn" id="back" onclick="history.back()" style="display:none">&#8592;</button>
      <h1>PVH Field Lookup<small id="stamp"></small></h1>
      <button class="navbtn" id="theme" title="Toggle dark mode">&#9789;</button>
    </div>
    <div id="searchwrap">
      <input id="q" type="search" placeholder="Deck # / plate / name / licence&#8230;"
        autocomplete="off" autocorrect="off" autocapitalize="characters" spellcheck="false">
      <button id="clr">&#10005;</button>
    </div>
  </header>
  <div id="toast" style="display:none"></div>
  <main id="main"></main>
</div>
__DATA_SCRIPT__
<script>
"use strict";
/* Bumped whenever the app itself changes, so a phone can be checked against
   what was deployed. Shown on the home screen under the data build stamp. */
var APP_VERSION="__APPVER__";
function boot(DB){
const V = DB.vehicles, O = DB.operators, W = DB.owners||[];
document.getElementById("stamp").textContent =
  "Data: " + DB.built + " \u00B7 " + DB.counts.vehicles + " vehicles \u00B7 " + DB.counts.operators +
  " operators \u00B7 " + (DB.counts.owners||0) + " owners";

/* ---------- search index ---------- */
const alnum = s => (s||"").toString().toUpperCase().replace(/[^A-Z0-9]/g,"");
const lc = s => (s||"").toString().toLowerCase();
V.forEach((v,i)=>{v._i=i;
  v._deck=alnum(v["Deck No"]); v._plate=alnum(v["Plate No"]); v._vin=alnum(v["VIN"]);
  v._name=lc(v["Owner Last Name"])+" "+lc(v["Owner First Name"]);
  v._biz=lc(v["Business Name"]); v._lic=alnum(v["Licence No"]);});
O.forEach((o,i)=>{o._i=i;
  o._name=lc(o["Last Name"])+" "+lc(o["First Name"])+" "+lc(o["Middle"]);
  o._biz=lc(o["Business Name"]); o._lic=alnum(o["Licence Number"]);});
W.forEach((w,i)=>{w._i=i;
  w._name=lc(w["Owner Last Name"])+" "+lc(w["Owner First Name"]);
  w._biz=lc(w["Business Name"]);});

function search(qRaw){
  const q=qRaw.trim(); if(q.length<2 && !/^\d$/.test(q)) return null;
  const qa=alnum(q), ql=lc(q);
  const vres=[], ores=[], wres=[];
  for(const v of V){
    let score=-1;
    if(qa && v._deck && v._deck===qa) score=100;
    else if(qa && v._plate && (v._plate===qa?1:0)) score=95;
    else if(qa && v._plate && v._plate.startsWith(qa) && qa.length>=3) score=80;
    else if(ql.length>=2 && v._name.includes(ql)) score=60;
    else if(ql.length>=2 && v._biz.includes(ql)) score=50;
    else if(qa && v._lic && v._lic===qa) score=90;
    else if(qa.length>=5 && v._vin && v._vin.includes(qa)) score=70;
    else if(qa && v._deck && v._deck.startsWith(qa) && qa.length<v._deck.length) score=40;
    if(score>=0) vres.push([score,v]);
  }
  for(const o of O){
    let score=-1;
    if(ql.length>=2 && o._name.includes(ql)) score=60;
    else if(ql.length>=2 && o._biz.includes(ql)) score=50;
    else if(qa && o._lic && o._lic===qa) score=90;
    if(score>=0) ores.push([score,o]);
  }
  for(const w of W){
    let score=-1;
    if(ql.length>=2 && w._name.includes(ql)) score=60;
    else if(ql.length>=2 && w._biz.includes(ql)) score=50;
    if(score>=0) wres.push([score,w]);
  }
  vres.sort((a,b)=>b[0]-a[0]); ores.sort((a,b)=>b[0]-a[0]); wres.sort((a,b)=>b[0]-a[0]);
  return {v:vres.map(x=>x[1]).slice(0,60), o:ores.map(x=>x[1]).slice(0,60), w:wres.map(x=>x[1]).slice(0,60)};
}

/* ---------- date status ---------- */
function parseD(s){
  if(!s) return null;
  const m=/^(\d{2})\/(\d{2})\/(\d{4})$/.exec(s);
  return m? new Date(+m[3],+m[1]-1,+m[2]) : null;
}
function status(label,dateStr){
  const d=parseD(dateStr);
  if(!d) return '<span class="badge b-na">'+label+': \u2014</span>';
  const days=Math.floor((d-new Date().setHours(0,0,0,0))/864e5);
  if(days<0)  return '<span class="badge b-bad">'+label+' <b>EXPIRED</b> '+dateStr+'</span>';
  if(days<=30)return '<span class="badge b-warn">'+label+' expires '+dateStr+'</span>';
  return '<span class="badge b-ok">'+label+' valid to '+dateStr+'</span>';
}

/* ---------- rendering ---------- */
const main=document.getElementById("main");
const esc=s=>(s==null?"":String(s)).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
const dash=s=>(s==null||s==="")?"\u2014":esc(s);

function deckHTML(v,big){
  const d=v["Deck No"];
  const t=(v["Vehicle Type"]||"").toLowerCase();
  if(d==null||d==="") return '<div class="deck none">NO<br>DECK</div>';
  return '<div class="deck '+t+'"><span>DECK</span>'+esc(d)+'</div>';
}
function ownerName(v){
  const n=[v["Owner First Name"],v["Owner Last Name"]].filter(Boolean).join(" ");
  return n||v["Business Name"]||"\u2014";
}
function opName(o){
  return [o["First Name"],o["Middle"],o["Last Name"]].filter(Boolean).join(" ")||"\u2014";
}

function vehCard(v,extra){
  return '<div class="card" onclick="go(\'v/'+v._i+'\')">'+deckHTML(v)+
    '<div class="cmain"><div class="cname">'+esc(v["Make Model"]||"Vehicle")+
    (v["Vehicle Color"]?' \u00B7 '+esc(v["Vehicle Color"]):'')+'</div>'+
    '<div class="csub">'+esc(ownerName(v))+(v["Business Name"]?' \u00B7 '+esc(v["Business Name"]):'')+'</div>'+
    '<div class="crow2"><span class="chip t-'+esc(v["Vehicle Type"])+'">'+esc(v["Vehicle Type"])+'</span>'+
    (v["Plate No"]?'<span class="plate">'+esc(v["Plate No"])+'</span>':'')+
    '</div>'+(extra||'')+'</div><div class="chev">&#8250;</div></div>';
}
/* Check history -- a rolling, per-device record of when a driver was last
   looked at, and when one was last deliberately marked as checked. The point
   is to stop the same people being pulled over twice in a week by different
   patrols.

   Two things it deliberately does. It lives in its own localStorage key, so
   the daily refresh -- which replaces pvh_data wholesale -- leaves it alone.
   And it keys on the normalized Master Number rather than the array index the
   routes use, because those indices shift on every rebuild. The normalizer
   mirrors norm_master() in the builder so an entry survives the same prefix
   and spacing noise the operator dedup already absorbs.

   Per-device only. Nothing syncs; another officer's checks are not visible
   here, and clearing the browser's storage loses it. */
var CLOG_KEY="pvh_check_log", CLOG_DAYS=180, CLOG_MAX=900, CLOG=null;
function clKey(o){
  var s=String(o["Master Number"]==null?"":o["Master Number"]).trim().toUpperCase();
  s=s.replace(/^4A\b/,"").replace(/^0*4\s+/,"").replace(/\s+/g,"");
  return s||("#"+(o["ID"]==null?"?":o["ID"]));
}
function clLoad(){
  if(CLOG)return CLOG;
  try{
    var r=localStorage.getItem(CLOG_KEY), o=r?JSON.parse(r):{};
    CLOG=(o&&typeof o==="object"&&!Array.isArray(o))?o:{};
  }catch(e){CLOG={};}
  return CLOG;
}
/* Entries are [lastSeen, lastChecked] in epoch seconds, 0 for never. Pruned
   on every write so the log cannot grow without bound on a handset. */
function clSave(){
  var m=clLoad(), cut=Math.floor(Date.now()/1000)-CLOG_DAYS*86400;
  Object.keys(m).forEach(function(k){
    var e=m[k];
    if(!e||Math.max(e[0]||0,e[1]||0)<cut)delete m[k];
  });
  var ks=Object.keys(m);
  if(ks.length>CLOG_MAX){
    ks.sort(function(a,b){return Math.max(m[b][0]||0,m[b][1]||0)-Math.max(m[a][0]||0,m[a][1]||0);});
    ks.slice(CLOG_MAX).forEach(function(k){delete m[k];});
  }
  try{localStorage.setItem(CLOG_KEY,JSON.stringify(m));}catch(e){}
}
function clGet(o){
  var e=clLoad()[clKey(o)];
  return {seen:(e&&e[0])||0, checked:(e&&e[1])||0};
}
function clMark(o,deliberate){
  var m=clLoad(), k=clKey(o), e=m[k]||[0,0];
  e[deliberate?1:0]=Math.floor(Date.now()/1000);
  m[k]=e; clSave();
}
function clDays(ts){return Math.floor((Date.now()/1000-ts)/86400);}
function clAgo(ts){
  if(!ts)return "";
  var d=clDays(ts);
  if(d<=0)return "today";
  if(d===1)return "yesterday";
  if(d<7)return d+" days ago";
  if(d<14)return "last week";
  if(d<60)return Math.round(d/7)+" weeks ago";
  return Math.round(d/30)+" months ago";
}
function clShort(ts){
  if(!ts)return "";
  var d=clDays(ts);
  return d<=0?"today":(d<7?d+"d":(d<60?Math.round(d/7)+"w":Math.round(d/30)+"mo"));
}
/* A marked check outranks a passing look, so only the stronger one shows. */
function clChips(o){
  var s=clGet(o);
  if(s.checked)return '<span class="chip clchk">Checked '+clShort(s.checked)+'</span>';
  if(s.seen)return '<span class="chip clseen">seen '+clShort(s.seen)+'</span>';
  return "";
}
function clBtnLabel(s){
  return s.checked?("Checked "+clAgo(s.checked)+" \u00B7 mark again"):"Mark as checked";
}
function markChecked(i){
  var o=O[i]; if(!o)return;
  clMark(o,true);
  renderOperator(i);
}
function clearCheckLog(){
  if(!confirm("Clear this device's check history? It cannot be undone, and this is the only copy."))return;
  try{localStorage.removeItem(CLOG_KEY);}catch(e){}
  CLOG=null;
  renderHome(q.value);
}
function checkedCards(){
  var m=clLoad(), ks=Object.keys(m).filter(function(k){return m[k][1];});
  if(!ks.length)return "";
  ks.sort(function(a,b){return m[b][1]-m[a][1];});
  var byKey={};
  O.forEach(function(o,ix){var k=clKey(o); if(byKey[k]==null)byKey[k]=ix;});
  /* RECENT already lists whatever was opened this session, directly above.
     Skip those so a driver looked at a minute ago is not printed twice. */
  var rows=ks.map(function(k){
    var ix=byKey[k];
    return (ix==null||RECENT.indexOf("o/"+ix)>-1)?"":opCard(O[ix]);
  }).filter(Boolean).slice(0,6).join("");
  if(!rows)return "";
  return '<div class="seclabel">Checked recently</div>'+rows+
    '<div class="hint"><span class="link" onclick="clearCheckLog()">Clear check history</span></div>';
}

function opCard(o,extra){
  const init=((o["First Name"]||" ")[0]+(o["Last Name"]||" ")[0]).toUpperCase();
  const flag=(o["In Active"]===true||o["Cancelled"]===true);
  const ownop=(o._veh&&o._veh.length>0);
  return '<div class="card" onclick="go(\'o/'+o._i+'\')">'+
    '<div class="opdot">'+esc(init)+'</div>'+
    '<div class="cmain"><div class="cname">'+esc(opName(o))+'</div>'+
    '<div class="csub">'+dash(o["Business Name"])+'</div>'+
    '<div class="crow2"><span class="chip t-'+esc(o["Operator Type"])+'">'+esc(o["Operator Type"])+' operator</span>'+
    (o["Licence Number"]?'<span class="plate">'+esc(o["Licence Number"])+'</span>':'')+
    (ownop?'<span class="chip" style="background:rgba(95,174,255,.16);color:var(--accent)">OWNER-OPERATOR</span>':'')+
    (flag?'<span class="badge b-bad">INACTIVE/CANCELLED</span>':'')+
    clChips(o)+
    '</div>'+(extra||'')+'</div><div class="chev">&#8250;</div></div>';
}
function ownerCard(w,extra){
  const init=((w["Owner First Name"]||" ")[0]+(w["Owner Last Name"]||" ")[0]).toUpperCase();
  const ownop=(w._op!=null);
  const n=(w._veh||[]).length;
  return '<div class="card" onclick="go(\'w/'+w._i+'\')">'+
    '<div class="opdot">'+esc(init)+'</div>'+
    '<div class="cmain"><div class="cname">'+esc(ownerName(w))+'</div>'+
    '<div class="csub">'+dash(w["Business Name"])+'</div>'+
    '<div class="crow2"><span class="chip" style="background:var(--panel2);color:var(--dim)">OWNER</span>'+
    '<span class="chip" style="background:var(--panel2);color:var(--dim)">'+n+' vehicle'+(n===1?"":"s")+'</span>'+
    (ownop?'<span class="chip" style="background:rgba(95,174,255,.16);color:var(--accent)">OWNER-OPERATOR</span>':'')+
    '</div>'+(extra||'')+'</div><div class="chev">&#8250;</div></div>';
}

function daysPast(s){
  const d=parseD(s); if(!d) return null;
  const n=Math.floor((new Date().setHours(0,0,0,0)-d)/864e5);
  return n>0?n:null;
}
function daysUntil(s){
  const d=parseD(s); if(!d) return null;
  const n=Math.floor((d-new Date().setHours(0,0,0,0))/864e5);
  return n>=0?n:null;
}
function addYear(s){
  const d=parseD(s); if(!d) return null;
  const due=new Date(d.getFullYear()+1,d.getMonth(),d.getDate());
  return String(due.getMonth()+1).padStart(2,"0")+"/"+String(due.getDate()).padStart(2,"0")+"/"+due.getFullYear();
}
function chkInto(out,soon,k,label,date){
  if(date==null||parseD(date)==null){out.push({k:k,label:label,date:null,miss:1,w:1e9});return;}
  const p=daysPast(date);
  if(p){out.push({k:k,label:label,date:date,over:p,w:p});return;}
  if(soon){const u=daysUntil(date); if(u!=null&&u<=30)out.push({k:k,label:label,date:date,soon:u,w:-1-u});}
}
function vehFlagList(v,soon){
  const out=[];
  chkInto(out,soon,"mvi","MVI",addYear(v["First MVIDate"]));
  chkInto(out,soon,"permit","PVH License",v["Expiry Date"]);
  chkInto(out,soon,"ins","Insurance",v["Insurance Expiry"]);
  chkInto(out,soon,"vlic","NS Permit",v["NSVehicle Permit Expiry"]);
  return out;
}
function opFlagList(o,soon){
  const out=[];
  if(o["In Active"]===true||o["Cancelled"]===true){
    out.push({k:"inactive",label:o["Cancelled"]===true?"Cancelled":"Inactive",date:null,over:0,w:0});
    return out;
  }
  chkInto(out,soon,"olic","Op licence",o["Renewal Date"]);
  chkInto(out,soon,"nsdl","NS DL",o["NSDLExpired Date"]);
  return out;
}
const FILTERS=[["all","All"],["permit","PVH License"],["ins","Insurance"],["vlic","NS Permit"],
  ["mvi","MVI"],["olic","Op licence"],["nsdl","NS DL"],["inactive","Inactive"]];
function buildFlagged(soon){
  const vi=[],oi=[];
  for(const v of V){const f=vehFlagList(v,soon); if(f.length)vi.push([v,f]);}
  for(const o of O){const f=opFlagList(o,soon); if(f.length)oi.push([o,f]);}
  return {vi,oi};
}
function flagBadges(f){
  return '<div class="cardbadges">'+f.map(function(x){
    if(x.miss)return '<span class="badge b-warn">'+x.label+' NONE ON FILE</span>';
    if(x.soon!=null)return '<span class="badge b-warn">'+x.label+' in '+x.soon+'d</span>';
    return '<span class="badge b-bad">'+x.label+(x.over>0?' '+x.over+'d over':'')+'</span>';
  }).join("")+'</div>';
}
function renderFlags(fk,soon){
  fk=fk||"all";
  const all=buildFlagged(soon);
  const cnt={all:all.vi.length+all.oi.length};
  for(const[,f]of all.vi)for(const x of f)cnt[x.k]=(cnt[x.k]||0)+1;
  for(const[,f]of all.oi)for(const x of f)cnt[x.k]=(cnt[x.k]||0)+1;
  let vshow=all.vi,oshow=all.oi;
  if(fk!=="all"){
    vshow=all.vi.filter(([,f])=>f.some(x=>x.k===fk)).map(([v,f])=>[v,f.filter(x=>x.k===fk)]);
    oshow=all.oi.filter(([,f])=>f.some(x=>x.k===fk)).map(([o,f])=>[o,f.filter(x=>x.k===fk)]);
  }
  const worst=f=>Math.max(...f.map(x=>x.w));
  vshow.sort((a,b)=>worst(b[1])-worst(a[1]));
  oshow.sort((a,b)=>worst(b[1])-worst(a[1]));
  let h='<div class="fchips">'+FILTERS.map(([k,l])=>
    '<div class="fchip'+(k===fk?" on":"")+'" onclick="go(\'flags/'+k+(soon?"/soon":"")+'\')">'+l+
    ' \u00B7 '+(cnt[k]||0)+'</div>').join("")+
    '<div class="fchip'+(soon?" on":"")+'" onclick="go(\'flags/'+fk+(soon?"":"/soon")+'\')">\u226430d '+(soon?"ON":"OFF")+'</div></div>';
  if(vshow.length)h+='<div class="seclabel">Vehicles ('+vshow.length+')</div>'+
    vshow.map(([v,f])=>vehCard(v,flagBadges(f))).join("");
  if(oshow.length)h+='<div class="seclabel">Operators ('+oshow.length+')</div>'+
    oshow.map(([o,f])=>opCard(o,flagBadges(f))).join("");
  if(!vshow.length&&!oshow.length)h+='<div class="hint">Nothing flagged in this category.</div>';
  main.innerHTML=h;
}
const RECENT=[];
function noteRecent(id){
  const ix=RECENT.indexOf(id); if(ix>-1)RECENT.splice(ix,1);
  RECENT.unshift(id); if(RECENT.length>5)RECENT.length=5;
}
function statusText(label,date){
  if(date==null||parseD(date)==null)return label+": none on file";
  const p=daysPast(date); if(p)return label+": EXPIRED "+date+" ("+p+"d over)";
  const u=daysUntil(date);
  return label+": valid to "+date+(u!=null&&u<=30?" (in "+u+"d)":"");
}
function vehSummary(v){
  return ["PVH VEHICLE \u2014 Deck "+(v["Deck No"]==null?"\u2014":v["Deck No"])+" / Plate "+(v["Plate No"]||"\u2014"),
    [v["Vehicle Type"],v["v Year"],v["Make Model"],v["Vehicle Color"]].filter(Boolean).join(" "),
    "Owner: "+ownerName(v)+(v["Business Name"]?" \u00B7 "+v["Business Name"]:""),
    v["Owner Address"]?"Owner address: "+v["Owner Address"]:null,
    (v._op!=null&&(O[v._op]["Cell Phone"]||O[v._op]["Phone"]))?"Owner phone (op. record): "+(O[v._op]["Cell Phone"]||O[v._op]["Phone"]):null,
    "VIN: "+(v["VIN"]||"\u2014")+" \u00B7 Veh lic no: "+(v["Licence No"]||"\u2014"),
    statusText("PVH License",v["Expiry Date"]),
    statusText("Insurance",v["Insurance Expiry"]),
    statusText("NS Permit",v["NSVehicle Permit Expiry"]),
    statusText("MVI due",addYear(v["First MVIDate"])),
    "Data as of "+DB.built+" \u00B7 copied "+new Date().toLocaleString()].filter(x=>x!=null).join("\n");
}
function opSummary(o){
  const inact=(o["In Active"]===true||o["Cancelled"]===true);
  return ["PVH OPERATOR \u2014 "+opName(o)+" ("+(o["Operator Type"]||"")+")",
    "Business: "+(o["Business Name"]||"\u2014"),
    "Licence no: "+(o["Licence Number"]||"\u2014")+" \u00B7 Master no: "+(o["Master Number"]||"\u2014")+" \u00B7 ID: "+(o["ID"]==null?"\u2014":o["ID"]),
    "Address: "+[o["Address1"],o["Address2"],o["City"],o["Postal Code"]].filter(Boolean).join(", "),
    "Phone: "+(o["Phone"]||"\u2014")+" \u00B7 Cell: "+(o["Cell Phone"]||"\u2014"),
    inact?("STATUS: "+(o["Cancelled"]===true?"CANCELLED":"INACTIVE")+" in system"):
      statusText("Op licence",o["Renewal Date"])+"\n"+statusText("NS DL",o["NSDLExpired Date"]),
    "Data as of "+DB.built+" \u00B7 copied "+new Date().toLocaleString()].join("\n");
}
function ownerSummary(w){
  const op=(w._op!=null)?O[w._op]:null;
  return ["PVH OWNER — "+ownerName(w),
    "Business: "+(w["Business Name"]||"—"),
    "Owner ID: "+(w["Owner ID"]==null?"—":w["Owner ID"]),
    "Address: "+(w["Owner Address"]||"—"),
    "Vehicles: "+(w._veh||[]).length,
    op?("OWNER-OPERATOR — also licensed as "+opName(op)):"Not separately licensed as an operator",
    "Data as of "+DB.built+" · copied "+new Date().toLocaleString()].join("\n");
}
function copyRec(kind,i,btn){
  const t=kind==="v"?vehSummary(V[i]):(kind==="w"?ownerSummary(W[i]):opSummary(O[i]));
  const done=()=>{btn.textContent="Copied";setTimeout(()=>{btn.textContent="Copy record summary";},1200);};
  const fallback=()=>{
    const ta=document.createElement("textarea");ta.value=t;ta.style.position="fixed";ta.style.opacity="0";
    document.body.appendChild(ta);ta.focus();ta.select();
    try{document.execCommand("copy");done();}catch(e){btn.textContent="Copy failed";}
    document.body.removeChild(ta);};
  if(navigator.clipboard&&navigator.clipboard.writeText){
    navigator.clipboard.writeText(t).then(done).catch(fallback);
  }else fallback();
}
function recentCards(){
  if(!RECENT.length)return "";
  return '<div class="seclabel">Recent</div>'+RECENT.map(id=>{
    const p=id.split("/");
    return p[0]==="v"?vehCard(V[+p[1]]):(p[0]==="w"?ownerCard(W[+p[1]]):opCard(O[+p[1]]));
  }).join("");
}
var HFILTER={type:"",stat:""};
const TYPE_OPTS=[["","All types"],["Taxi","Taxi"],["Limo","Limo"],["Tour","Tour"],["Shuttle","Shuttle"]];
const STAT_OPTS=[["","Any status"],["expired","Any expired"],["permit","PVH License exp"],["ins","Insurance exp"],
  ["mvi","MVI overdue"],["vlic","NS Permit exp"],["ok","All valid"]];
function vehMatchesStat(v,s){
  if(!s)return true;
  const fl=vehFlagList(v,false).filter(x=>!x.miss);
  const keys=fl.map(x=>x.k);
  if(s==="expired")return keys.length>0;
  if(s==="ok")return keys.length===0;
  return keys.indexOf(s)>-1;
}
function applyBrowseFilters(list){
  return list.filter(v=>
    (!HFILTER.type||v["Vehicle Type"]===HFILTER.type)&&
    vehMatchesStat(v,HFILTER.stat));
}
function selectEl(id,opts,val){
  return '<select id="'+id+'">'+opts.map(([k,l])=>
    '<option value="'+k+'"'+(k===val?" selected":"")+'>'+l+'</option>').join("")+'</select>';
}
function renderHome(q){
  const r=search(q||"");
  if(!r){
    if(HFILTER.type||HFILTER.stat){renderBrowse();return;}
    const fl=buildFlagged(false), n=fl.vi.length+fl.oi.length;
    const cnt={};
    for(const v of V)for(const x of vehFlagList(v,false).filter(y=>!y.miss))cnt[x.k]=(cnt[x.k]||0)+1;
    const shortcut=(k,label)=>cnt[k]?'<div class="fchip" onclick="go(\'flags/'+k+'\')">'+label+' \u00B7 '+cnt[k]+'</div>':'';
    main.innerHTML='<div class="hint">Search by deck light number, plate, owner or operator name, business or licence number.</div>'+
      '<div class="browsebar">'+selectEl("fType",TYPE_OPTS,HFILTER.type)+selectEl("fStat",STAT_OPTS,HFILTER.stat)+'</div>'+
      (n?'<div class="card" onclick="go(\'flags\')">'+
        '<div class="opdot" style="color:var(--bad);border-color:var(--bad)">&#9888;</div>'+
        '<div class="cmain"><div class="cname">'+n+' flagged records</div>'+
        '<div class="csub">Expired or missing: PVH licences, insurance, NS Permits, MVI \u00B7 inactive operators</div></div>'+
        '<div class="chev">&#8250;</div></div>':'')+
      '<div class="fchips" style="margin-top:8px">'+shortcut("permit","PVH Licenses")+shortcut("ins","Insurance")+
        shortcut("mvi","MVI")+shortcut("vlic","NS Permit")+'</div>'+
      recentCards()+
      checkedCards()+
      '<div class="counts">'+DB.counts.vehicles+' active vehicles \u00B7 '+DB.counts.operators+' active operators \u00B7 '+(DB.counts.owners||0)+' owners</div>'+
      '<div class="stamp">Built '+esc(DB.built)+' from '+esc(DB.sources.vehicles)+' + '+esc(DB.sources.operators)+
        '<br>App version '+APP_VERSION+'</div>'+
      (window.__SHELL__?'<div class="hint">'+(window.__syncLine?window.__syncLine():'')+
        '<span class="link" onclick="window.__checkNow()">Check for new data</span> \u00B7 '+
        '<span class="link" onclick="window.__datasrc()">Data source\u2026</span> \u00B7 '+
        '<span class="link" onclick="window.__reimport()">Load a file\u2026</span></div>':'');
    wireFilters();
    return;
  }
  let h='<div class="counts">'+r.v.length+' vehicle'+(r.v.length===1?"":"s")+', '+r.o.length+' operator'+(r.o.length===1?"":"s")+', '+r.w.length+' owner'+(r.w.length===1?"":"s")+'</div>';
  if(r.v.length){h+='<div class="seclabel">Vehicles ('+r.v.length+')</div>'+r.v.map(v=>vehCard(v)).join("");}
  if(r.w.length){h+='<div class="seclabel">Owners ('+r.w.length+')</div>'+r.w.map(w=>ownerCard(w)).join("");}
  if(r.o.length){h+='<div class="seclabel">Operators ('+r.o.length+')</div>'+r.o.map(o=>opCard(o)).join("");}
  if(!r.v.length&&!r.o.length&&!r.w.length){
    // U4: absence-as-signal
    const aq=alnum(q); const looksPlate=aq.length>=2&&aq.length<=8&&/[0-9]/.test(aq)&&/^[A-Z0-9]+$/.test(aq);
    h='<div class="alertbar" style="background:rgba(245,190,74,.14);color:var(--warn);border-color:rgba(245,190,74,.4)">'+
      'No active record matches \u201C'+esc(q)+'\u201D as of '+esc(DB.built)+'.</div>'+
      '<div class="hint">'+(looksPlate?
        'A plate or deck not in the active list may be suspended, expired, transferred or never licensed \u2014 verify licensing status before clearing.':
        'Check spelling, or try the plate / deck number instead.')+'</div>';
  }
  main.innerHTML=h;
}
function wireFilters(){
  const t=document.getElementById("fType"), s=document.getElementById("fStat");
  if(t)t.addEventListener("change",()=>{HFILTER.type=t.value;renderHome("");});
  if(s)s.addEventListener("change",()=>{HFILTER.stat=s.value;renderHome("");});
}
var BSORT="over";
function renderBrowse(){
  let list=applyBrowseFilters(V.slice());
  if(BSORT==="over")list.sort((a,b)=>{
    const wa=Math.max(0,...vehFlagList(a,false).filter(x=>!x.miss).map(x=>x.w));
    const wb=Math.max(0,...vehFlagList(b,false).filter(x=>!x.miss).map(x=>x.w));
    return wb-wa;});
  else if(BSORT==="deck")list.sort((a,b)=>String(a["Deck No"]||"~").localeCompare(String(b["Deck No"]||"~"),undefined,{numeric:true}));
  else if(BSORT==="owner")list.sort((a,b)=>ownerName(a).localeCompare(ownerName(b)));
  const tl=TYPE_OPTS.find(x=>x[0]===HFILTER.type), sl=STAT_OPTS.find(x=>x[0]===HFILTER.stat);
  let h='<div class="browsebar">'+selectEl("fType",TYPE_OPTS,HFILTER.type)+selectEl("fStat",STAT_OPTS,HFILTER.stat)+'</div>'+
    '<div class="fchips"><div class="fchip'+(BSORT==="over"?" on":"")+'" onclick="setSort(\'over\')">Most overdue</div>'+
    '<div class="fchip'+(BSORT==="deck"?" on":"")+'" onclick="setSort(\'deck\')">Deck #</div>'+
    '<div class="fchip'+(BSORT==="owner"?" on":"")+'" onclick="setSort(\'owner\')">Owner A\u2013Z</div></div>'+
    '<div class="counts">'+list.length+' vehicle'+(list.length===1?"":"s")+
      ' \u00B7 '+(tl?tl[1]:"")+(sl&&sl[0]?" \u00B7 "+sl[1]:"")+'</div>';
  h+=list.length?list.map(v=>vehCard(v,statBadgesInline(v))).join(""):'<div class="hint">No vehicles match this filter.</div>';
  main.innerHTML=h;
  wireFilters();
}
function setSort(s){BSORT=s;renderBrowse();}
function statBadgesInline(v){
  const fl=vehFlagList(v,false).filter(x=>!x.miss);
  if(!fl.length)return "";
  return '<div class="cardbadges">'+fl.map(x=>'<span class="badge b-bad">'+x.label+(x.over>0?" "+x.over+"d":"")+'</span>').join("")+'</div>';
}

function row(k,v,mono){
  return '<div class="row"><div class="k">'+k+'</div><div class="v'+(mono?' mono':'')+'">'+v+'</div></div>';
}
function copyField(el,text){
  const done=()=>{const o=el.textContent;el.classList.add("copied");el.textContent=text+" \u2713";
    setTimeout(()=>{el.classList.remove("copied");el.textContent=o;},1000);};
  if(navigator.clipboard&&navigator.clipboard.writeText)navigator.clipboard.writeText(text).then(done).catch(done);
  else done();
}
function tapField(text){
  const t=esc(text);
  return '<span class="tap" onclick="copyField(this,\''+t.replace(/'/g,"\\'")+'\')">'+t+'</span>';
}
function telLink(p){
  if(!p) return "\u2014";
  const digits=String(p).replace(/[^0-9]/g,"");
  return digits.length>=7?'<a class="tel" href="tel:'+digits+'">'+esc(p)+'</a>':esc(p);
}

function renderVehicle(i){
  const v=V[i]; if(!v){renderHome("");return;}
  noteRecent("v/"+i);
  const owner=(v._owner!=null)?W[v._owner]:null;
  const others=owner?(owner._veh||[]).map(ix=>V[ix]).filter(x=>x!==v):[];
  const op=(v._op!=null)?O[v._op]:null;
  let h='<div class="dhead"><div class="dtop">'+deckHTML(v,true)+
    '<div><div class="dtitle">'+esc(v["Make Model"]||"Vehicle")+'</div>'+
    '<div class="dsub">'+dash(v["Vehicle Color"])+' \u00B7 '+dash(v["v Year"])+
    ' \u00B7 <span class="chip t-'+esc(v["Vehicle Type"])+'">'+esc(v["Vehicle Type"])+'</span></div>'+
    (v["Plate No"]?'<div class="bigplate tap" onclick="copyField(this,\''+esc(v["Plate No"]).replace(/'/g,"\\'")+'\')">'+esc(v["Plate No"])+'</div>':'')+
    '</div></div><div class="badges">'+
    status("PVH License",v["Expiry Date"])+
    status("Insurance",v["Insurance Expiry"])+
    status("NS Permit",v["NSVehicle Permit Expiry"])+
    status("MVI due",addYear(v["First MVIDate"]))+
    '</div></div>';
  h+='<button class="copybtn" onclick="copyRec(\'v\','+i+',this)">Copy record summary</button>';
  h+='<div class="grid">'+
    row("Owner",owner?('<span class="link" onclick="go(\'w/'+owner._i+'\')">'+esc(ownerName(v))+' &#8250;</span>'):dash(ownerName(v)))+
    (owner&&owner._op!=null?row("Owner-Operator",'<span class="link" onclick="go(\'o/'+owner._op+'\')">'+esc(opName(O[owner._op]))+' &#8250;</span>'):"")+
    (op&&(op["Cell Phone"]||op["Phone"])?row("Owner phone (op. record)",telLink(op["Cell Phone"]||op["Phone"])):"")+
    row("Business",dash(v["Business Name"]))+
    (v["Owner Address"]?row("Owner address",esc(v["Owner Address"])):"")+
    row("Owner ID",dash(v["Owner ID"]))+
    row("District",dash(v["District"]))+
    row("VIN",v["VIN"]?tapField(v["VIN"]):"\u2014",1)+
    row("Vehicle licence no",v["Licence No"]?tapField(v["Licence No"]):"\u2014",1)+
    row("Last inspection",dash(v["Inspection Date"]))+
    row("First MVI",dash(v["First MVIDate"]))+
    row("MVI due (MVI + 1 yr)",dash(addYear(v["First MVIDate"])))+
    row("Vehicle ID",dash(v["Vehicle ID"]))+
    '</div>';
  if(v["Notes"]) h+='<div class="seclabel">Notes</div><div class="notes">'+esc(v["Notes"])+'</div>';
  if(others.length){
    h+='<div class="seclabel">Other vehicles, same owner ('+others.length+')</div>'+others.map(vehCard).join("");
  }
  main.innerHTML=h;
}

function renderOperator(i){
  const o=O[i]; if(!o){renderHome("");return;}
  noteRecent("o/"+i);
  /* Read the log before recording this visit, otherwise the page would only
     ever report the look being taken right now. */
  const prior=clGet(o);
  clMark(o,false);
  const flag=(o["In Active"]===true||o["Cancelled"]===true);
  const ownedW=W.filter(w=>w._op===i);
  let h="";
  if(ownedW.length) h+='<div class="alertbar" style="background:rgba(95,174,255,.14);color:var(--accent);border-color:rgba(95,174,255,.4)">'+
    '&#9733; OWNER-OPERATOR &mdash; also registered as vehicle owner</div>';
  if(flag) h+='<div class="alertbar">&#9888; Licence flagged '+
    (o["Cancelled"]===true?'CANCELLED':'INACTIVE')+' in system</div>';
  if(prior.checked&&clDays(prior.checked)<14)
    h+='<div class="alertbar" style="background:rgba(95,174,255,.14);color:var(--accent);border-color:rgba(95,174,255,.4)">'+
      '&#9202; Already checked '+esc(clAgo(prior.checked))+' on this device</div>';
  h+='<div class="dhead"><div class="dtop"><div class="opdot" style="flex:0 0 66px;height:66px;font-size:22px">'+
    esc(((o["First Name"]||" ")[0]+(o["Last Name"]||" ")[0]).toUpperCase())+'</div>'+
    '<div><div class="dtitle">'+esc(opName(o))+'</div>'+
    '<div class="dsub">'+dash(o["Business Name"])+' \u00B7 <span class="chip t-'+esc(o["Operator Type"])+'">'+
    esc(o["Operator Type"])+' operator</span></div>'+
    (o["Licence Number"]?'<div class="bigplate">'+esc(o["Licence Number"])+'</div>':'')+
    '</div></div><div class="badges">'+
    status("Op licence",o["Renewal Date"])+
    status("NS DL",o["NSDLExpired Date"])+
    '</div></div>';
  const addr=[o["Address1"],o["Address2"],[o["City"],o["Province"],o["Postal Code"]].filter(Boolean).join(", ")]
    .filter(Boolean).map(esc).join("<br>");
  h+='<button class="copybtn" onclick="markChecked('+i+')">'+esc(clBtnLabel(prior))+'</button>';
  h+='<button class="copybtn" onclick="copyRec(\'o\','+i+',this)">Copy record summary</button>';
  h+='<div class="grid">'+
    row("Last checked",prior.checked?esc(clAgo(prior.checked)):"\u2014")+
    row("Last viewed",prior.seen?esc(clAgo(prior.seen)):"\u2014")+
    row("Address",addr||"\u2014")+
    row("Phone",telLink(o["Phone"]))+
    row("Cell",telLink(o["Cell Phone"]))+
    row("Master number",dash(o["Master Number"]),1)+
    row("Operator ID",dash(o["ID"]))+
    row("Licence year",dash(o["v Year"]))+
    row("Approved",dash(o["Approval Date"]))+
    row("District",dash(o["District"]))+
    '</div>';
  if(o["Notes"]) h+='<div class="seclabel">Notes</div><div class="notes">'+esc(o["Notes"])+'</div>';
  const veh=(o._veh||[]).map(ix=>V[ix]).filter(Boolean);
  h+='<div class="seclabel">Vehicles in this name ('+veh.length+')</div>';
  h+=veh.length?veh.map(vehCard).join(""):'<div class="empty">No active vehicles registered under this exact name.</div>';
  const fz=(o._vfz||[]).map(ix=>V[ix]).filter(Boolean);
  if(fz.length)h+='<div class="seclabel">Possible matches \u2014 similar name, verify before relying on ('+fz.length+')</div>'+fz.map(vehCard).join("");
  main.innerHTML=h;
}

function renderOwner(i){
  const w=W[i]; if(!w){renderHome("");return;}
  noteRecent("w/"+i);
  const op=(w._op!=null)?O[w._op]:null;
  let h='<div class="dhead"><div class="dtop">'+
    '<div class="opdot" style="flex:0 0 66px;height:66px;font-size:22px">'+
    esc(((w["Owner First Name"]||" ")[0]+(w["Owner Last Name"]||" ")[0]).toUpperCase())+'</div>'+
    '<div><div class="dtitle">'+esc(ownerName(w))+'</div>'+
    '<div class="dsub">'+dash(w["Business Name"])+' &middot; <span class="chip" style="background:var(--panel2);color:var(--dim)">OWNER</span>'+
    (op?' <span class="chip" style="background:rgba(95,174,255,.16);color:var(--accent)">OWNER-OPERATOR</span>':'')+
    '</div></div></div></div>';
  h+='<button class="copybtn" onclick="copyRec(\'w\','+i+',this)">Copy record summary</button>';
  h+='<div class="grid">'+
    row("Owner ID",dash(w["Owner ID"]))+
    row("Address",dash(w["Owner Address"]))+
    (op?row("Operator link",'<span class="link" onclick="go(\'o/'+op._i+'\')">'+esc(opName(op))+' &#8250;</span>'):
       row("Operator link",'<span class="empty">Not separately licensed as an operator</span>'))+
    '</div>';
  const veh=(w._veh||[]).map(ix=>V[ix]).filter(Boolean);
  h+='<div class="seclabel">Vehicles ('+veh.length+')</div>';
  h+=veh.length?veh.map(v=>vehCard(v)).join(""):'<div class="empty">No active vehicles on file for this owner.</div>';
  main.innerHTML=h;
}

/* ---------- routing ---------- */
const q=document.getElementById("q"), clr=document.getElementById("clr");
function go(route){location.hash="#/"+route;}
function route(){
  const h=location.hash.replace(/^#\/?/,"");
  const m=/^(v|o|w)\/(\d+)$/.exec(h);
  const fm=/^flags(?:\/([a-z]+))?(\/soon)?$/.exec(h);
  window.scrollTo(0,0);
  document.getElementById("back").style.display=(m||fm)?"flex":"none";
  if(m){ if(m[1]==="v") renderVehicle(+m[2]); else if(m[1]==="o") renderOperator(+m[2]); else renderOwner(+m[2]); }
  else if(fm){ renderFlags(fm[1],!!fm[2]); }
  else { renderHome(q.value); }
}
document.body.classList.add("light");
document.getElementById("theme").addEventListener("click",()=>{
  const light=document.body.classList.toggle("light");
  document.getElementById("theme").innerHTML=light?"&#9789;":"&#9788;";
});
q.addEventListener("input",()=>{
  clr.style.display=q.value?"block":"none";
  if(location.hash && location.hash!=="#/") history.replaceState(null,"","#/");
  renderHome(q.value);
});
clr.addEventListener("click",()=>{q.value="";clr.style.display="none";q.focus();renderHome("");});
window.go=go;window.copyRec=copyRec;window.copyField=copyField;window.setSort=setSort;
window.markChecked=markChecked;window.clearCheckLog=clearCheckLog;
window.__route=route;
window.addEventListener("hashchange",route);
route();
}
__BOOT__
</script>
</body>
</html>
"""

if __name__ == "__main__":
    main()
