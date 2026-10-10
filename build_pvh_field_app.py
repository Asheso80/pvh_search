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
import io
import os
import re
import sys
import json
import time
import hashlib
import tempfile
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


def write_out(path, data, encoding="utf-8"):
    """Write a build output by rename, never by truncating the target.

    Every output lands in the OneDrive-synced project folder, where the files
    are Files-On-Demand placeholders. Opening one with "w" truncates it in
    place, and the sync filter rejects that at random with OSError 22
    (Invalid argument) -- that is what killed the Sep 2 build after the pull
    had already succeeded. Writing a fresh temp file alongside the target and
    renaming it over top never truncates anything, and means a build that dies
    halfway can't leave a half-written app on the handsets.
    """
    folder = os.path.dirname(os.path.abspath(path))
    fd, tmp = tempfile.mkstemp(dir=folder, prefix="._pvh_", suffix=".tmp")
    try:
        if isinstance(data, bytes):
            with os.fdopen(fd, "wb") as f:
                f.write(data)
        else:
            with os.fdopen(fd, "w", encoding=encoding) as f:
                f.write(data)
        # The rename can still lose a race with the sync client or a scanner
        # holding the old file open. That clears in moments, so wait it out
        # rather than failing a run that has already done all the real work.
        for attempt in range(4):
            try:
                os.replace(tmp, path)
                return
            except OSError:
                if attempt == 3:
                    raise
                time.sleep(1.5)
    except BaseException:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


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
# a live record: Expiry Date 01/31/2027, NSVehicle Permit
# Expiry 07/31/2026 -- its PVH licence is current. Do not transpose these.
VEH_DATES = ["NSVehicle Permit Expiry", "First MVIDate", "Insurance Expiry",
             "Inspection Date", "Expiry Date"]
OP_DATES = ["NSDLExpired Date", "Approval Date", "Renewal Date"]

DATE_WARNINGS = []
DEDUP_LOG = []
OWNER_NAME_DRIFT = []
LINK_AMBIGUOUS = []

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


STREET_ABBR = {
    "STREET": "ST", "AVENUE": "AVE", "AV": "AVE", "ROAD": "RD", "DRIVE": "DR",
    "LANE": "LN", "CRESCENT": "CRES", "COURT": "CT", "BOULEVARD": "BLVD",
    "PLACE": "PL", "HIGHWAY": "HWY", "TERRACE": "TERR", "CIRCLE": "CIR",
}
UNIT_WORDS = {"APT", "APARTMENT", "UNIT", "SUITE", "STE"}


def norm_street(s):
    """Normalize a street line for tie-breaking operators who share a name.

    Case, punctuation and spacing are ignored, common street-type words are
    abbreviated (STREET -> ST) and a unit suffix ("APT 2", "UNIT 5", "#3") is
    dropped, so "10 Sample Street" and "10 SAMPLE ST APT 1" compare equal.
    Only the street line is used -- the owner address carries city and postal
    code after the first " · ", and those are left out.
    """
    if not s:
        return None
    s = str(s).split("·")[0].upper()
    s = re.sub(r"#\s*\w+", " ", s)
    toks = re.findall(r"[A-Z0-9]+", s)
    out = []
    skip = False
    for t in toks:
        if skip:
            skip = False
            continue
        if t in UNIT_WORDS:
            skip = True
            continue
        out.append(STREET_ABBR.get(t, t))
    return " ".join(out) or None


def resolve_operator(name_key, street, ops_by_name, operators):
    """Pick the operator record that is the same person as a vehicle owner.

    Returns (index, candidates). A name held by exactly one operator links as
    before. When several operators share the name, the owner's street line
    breaks the tie; if that does not single out exactly one operator, nobody
    is linked (index None) and the candidates are returned so the app can
    say "possible match, verify" instead of guessing. Linking the first
    operator in the file is how one vehicle ended up shown under the wrong operator.
    """
    cands = ops_by_name.get(name_key, []) if name_key else []
    if len(cands) == 1:
        return cands[0], []
    if not cands:
        return None, []
    s = norm_street(street)
    hits = [j for j in cands if s and norm_street(operators[j].get("Address1")) == s]
    if len(hits) == 1:
        return hits[0], []
    return None, list(cands)


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


def build_owners(vehicles, ops_by_name, operators):
    """Group vehicles into first-class Owner records.

    Grouped primarily by Owner ID -- a fully-populated identity key on
    every vehicle row, more reliable than the name-string matching used
    elsewhere in this file as a display-only fallback. BUT Owner ID is not
    always a clean 1:1 key: the source data has at least one confirmed
    case of the same Owner ID reused across two unrelated people (Owner ID
    1: one person with a Limo vehicle, another with
    five Tour vehicles). A same-ID group is therefore split further by
    normalized owner name so unrelated people never get merged onto one
    Owner card; a split logs a warning, since a reused Owner ID is a real
    source-data problem worth flagging, not cosmetic spelling drift.

    Owner-to-Operator linkage goes through resolve_operator(), the same
    exact-name match (address tie-break when a name is shared) the vehicles
    use -- it deliberately does NOT use the fuzzy tier, so an
    "Owner-Operator" label is never a guess. An unresolved shared name
    leaves _op None and lists the candidates in _opc. Mutates vehicles in
    place, setting v["_owner"].
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
            addr = most_common(rows, "Owner Address")
            op_ix, op_cands = resolve_operator(nkey, addr, ops_by_name, operators)
            owners.append({
                "Owner ID": oid,
                "Owner Last Name": last,
                "Owner First Name": first,
                "Business Name": most_common(rows, "Business Name"),
                "Owner Address": addr,
                "_veh": sub_idxs,
                "_op": op_ix,
                "_opc": op_cands,
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


def shared_config():
    """Connection details for the shared notes backend, or None.

    Read from pvh_local_config.json (gitignored) and carried to devices inside
    PVH_data.json -- the private data file -- rather than baked into the
    published app, so the public repo never names the project. The key is the
    public "anon" key, which is meant to be visible to clients; the database's
    access rules are what protect the data. A service_role key here would hand
    every device full control, so anything that looks like one is refused.
    """
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pvh_local_config.json")
    try:
        with io.open(path, encoding="utf-8") as f:
            cfg = json.load(f)
    except OSError:
        return None
    except ValueError as exc:
        print("WARNING: pvh_local_config.json is not valid JSON (" + str(exc) + ") -- "
              "shared notes left OFF. Check for a missing or trailing comma.")
        return None
    url = str(cfg.get("supabase_url", "")).strip().rstrip("/")
    # The dashboard hands out the REST address (".../rest/v1/"); the app adds
    # that part itself, so keep only the project's base address.
    for tail in ("/rest/v1", "/auth/v1"):
        if url.endswith(tail):
            url = url[: -len(tail)]
    key = str(cfg.get("supabase_anon_key", "")).strip()
    if not url and not key:
        return None
    if not (url.startswith("https://") and key):
        print("WARNING: supabase_url / supabase_anon_key in pvh_local_config.json are "
              "incomplete (need an https:// URL and a key) -- shared notes left OFF.")
        return None
    try:
        import base64
        seg = key.split(".")[1]
        role = json.loads(base64.urlsafe_b64decode(seg + "=" * (-len(seg) % 4))).get("role")
    except Exception:
        role = None
    if role == "service_role" or key.startswith("sb_secret_"):
        sys.exit("supabase_anon_key is a SECRET key. That key bypasses every access rule and "
                 "must never reach a device. Use the anon / publishable key.")
    print("Shared notes: ON (" + url + ")")
    return {"url": url, "key": key}


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
    # A name can belong to several licensed operators (16 names in the Sep 2026
    # data, e.g. three operators with one name), so each vehicle is resolved to at most
    # one operator via resolve_operator(); unresolved shared names are listed
    # as "possible matches" on every candidate instead of linked to all of them.
    veh_by_name = {}
    for i, v in enumerate(vehicles):
        key = norm_name(v["Owner Last Name"], v["Owner First Name"])
        if key:
            veh_by_name.setdefault(key, []).append(i)
    owner_keys = list(veh_by_name.keys())
    ops_by_name = {}
    for j, op in enumerate(operators):
        k = norm_name(op["Last Name"], op["First Name"])
        if k:
            ops_by_name.setdefault(k, []).append(j)
    for op in operators:
        op["_veh"], op["_vsn"] = [], []
    for i, v in enumerate(vehicles):
        k = norm_name(v["Owner Last Name"], v["Owner First Name"])
        v["_op"], v["_opc"] = resolve_operator(k, v.get("Owner Address"), ops_by_name, operators)
        if v["_op"] is not None:
            operators[v["_op"]]["_veh"].append(i)
        for j in v["_opc"]:
            operators[j]["_vsn"].append(i)
        if v["_opc"]:
            LINK_AMBIGUOUS.append(
                f"Deck {v.get('Deck No') if v.get('Deck No') not in (None, '') else '?'} "
                f"(plate {v.get('Plate No') or '?'}) — owner {v.get('Owner First Name')} "
                f"{v.get('Owner Last Name')} matches {len(v['_opc'])} operators, no address match: "
                + ", ".join(str(operators[j].get("Licence Number") or "?") for j in v["_opc"]))
    for op in operators:
        key = norm_name(op["Last Name"], op["First Name"])
        fz = []
        if key and len(key[0]) >= 3:
            for ok in owner_keys:
                if ok == key or len(ok[0]) < 3:
                    continue
                if lev1(key[0], ok[0]) and lev1(key[1], ok[1]):
                    fz.extend(veh_by_name[ok])
        op["_vfz"] = [i for i in fz if i not in op["_veh"] and i not in op["_vsn"]][:6]

    owners = build_owners(vehicles, ops_by_name, operators)

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
    shared = shared_config()
    if shared:
        data["shared"] = shared
    payload_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    # Escape "</" so the JSON can never terminate the script tag.
    payload = payload_json.replace("</", "<\\/")

    html = (TEMPLATE
            .replace("__HEAD_EXTRA__", "")
            .replace("__DATA_SCRIPT__", '<script id="data" type="application/json">' + payload + "</script>")
            .replace("__BOOT__", SINGLE_BOOT)
            .replace("__APPVER__", shell_version()[0]))
    write_out(OUT_FILE, html)
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
    warn += ["Operator link unconfirmed: " + a for a in LINK_AMBIGUOUS]
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

SW_JS = r"""const C="pvh-shell-__BUILD__", G="pvh-geo-__GEO__";
/* geo.bin is public map data, held in its own cache so a shell update does not
   re-download it. It is fetched at install so location works offline. */
self.addEventListener("install",e=>{e.waitUntil(Promise.all([
  caches.open(C).then(c=>c.addAll(["./","index.html","manifest.webmanifest","icon-192.png","icon-512.png"])),
  caches.open(G).then(c=>c.match("geo.bin").then(h=>h||c.add(new Request("geo.bin",{cache:"reload"})))).catch(()=>{})
]).then(()=>self.skipWaiting()))});
self.addEventListener("activate",e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==C&&x!==G).map(x=>caches.delete(x)))).then(()=>self.clients.claim()))});
self.addEventListener("fetch",e=>{
  if(e.request.method!=="GET")return;
  const u=new URL(e.request.url);
  /* Only the app shell is cached. Data fetches (Dropbox and friends) always go
     to the network, so a new build is actually seen and no record data is left
     behind in the cache. */
  if(u.origin!==self.location.origin)return;
  if(/\.json(\?|$)/i.test(u.pathname+u.search))return;
  if(/(^|\/)geo\.bin$/.test(u.pathname)){
    e.respondWith(caches.open(G).then(c=>c.match(e.request,{ignoreSearch:true}).then(r=>r||fetch(e.request).then(res=>{if(res.ok)c.put(e.request,res.clone());return res;}))));
    return;
  }
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
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        write_out(f"{folder}/icon-{px}.png", buf.getvalue())
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
        write_out(VERSION_FILE, json.dumps(rec, indent=2) + "\n")

    _SHELL_VER = (f"v{rec['version']} · {rec['date']}", digest)
    return _SHELL_VER


def write_shell(payload_json):
    os.makedirs("docs", exist_ok=True)
    label, _ = shell_version()
    shell_html = (TEMPLATE
                  .replace("__HEAD_EXTRA__", SHELL_HEAD)
                  .replace("__DATA_SCRIPT__", "")
                  .replace("__BOOT__", SHELL_BOOT)
                  .replace("__APPVER__", label))
    write_out("docs/index.html", shell_html)
    # Key the cache on the bytes actually served, not on the pre-substitution
    # digest: the version label is baked into this file, so a build that changes
    # only the label still needs a new cache key. Deterministic either way --
    # the label is settled before this is computed.
    cache_key = hashlib.sha256(shell_html.encode("utf-8")).hexdigest()[:8]
    try:
        with open("docs/geo.bin", "rb") as f:
            geo_key = hashlib.sha256(f.read()).hexdigest()[:8]
    except OSError:
        geo_key = "none"
    write_out("docs/sw.js", SW_JS.replace("__BUILD__", cache_key).replace("__GEO__", geo_key))
    write_out("docs/manifest.webmanifest", MANIFEST)
    make_icons("docs")
    write_out("PVH_data.json", payload_json)
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
/* Search result groups. Bigger than seclabel on purpose: a long result list
   has to show at a glance where Vehicles end and Owners begin. */
.reshead{display:flex;align-items:center;gap:8px;margin:18px 0 8px;padding:9px 12px;
  background:var(--panel2);border-left:4px solid var(--accent);border-radius:8px;
  font-size:15px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;color:var(--text)}
.reshead .n{margin-left:auto;font-size:13px;font-weight:700;letter-spacing:0;
  background:var(--accent);color:var(--bg);border-radius:10px;padding:1px 9px}
.jump{display:flex;gap:6px;justify-content:center;flex-wrap:wrap;padding-top:6px}
.jump span{font-size:13px;font-weight:600;color:var(--accent);background:var(--panel);
  border:1px solid var(--line);border-radius:14px;padding:4px 11px;cursor:pointer}
.jump span:active{background:var(--panel2)}

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
.deck span{font-size:10px;letter-spacing:.08em;font-weight:600;margin-bottom:2px}
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
.chip{font-size:12px;font-weight:700;letter-spacing:.05em;border-radius:5px;
  padding:2px 7px;text-transform:uppercase}
.t-Taxi{background:rgba(245,185,66,.14);color:var(--taxi)}
.t-Limo{background:rgba(176,140,255,.14);color:var(--limo)}
.t-Tour{background:rgba(77,163,255,.14);color:var(--tour)}
.clchk{background:rgba(95,174,255,.16);color:var(--accent)}
.clseen{background:var(--panel2);color:var(--faint)}
.clwhen{text-transform:none;letter-spacing:0}
.pairnote{display:inline-flex;align-items:center;gap:6px;margin-top:7px;
  padding:4px 9px;border-radius:7px;font-size:12px;font-weight:600;
  color:var(--accent);background:rgba(95,174,255,.12);
  border:1px solid rgba(95,174,255,.34)}
.pairnote b{font-size:13px;line-height:1}
.clpanel{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin:12px 0;
  padding:10px 13px;border:1px solid var(--line);border-radius:12px;
  background:var(--panel);font-size:13px}
.clpanel.clrecent{border-color:rgba(95,174,255,.5);background:rgba(95,174,255,.10)}
.cllabel{font-size:11px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;
  color:var(--faint);flex:0 0 auto}
.clline{flex:1 1 230px;color:var(--text)}
.clpanel.clrecent .clline{color:var(--accent)}
/* shared notes */
.nt{background:rgba(247,190,74,.2);color:var(--warn)}
.ntact{background:var(--bad);color:#fff}
.ntdone{background:var(--panel2);color:var(--ok)}
.ntbar{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin:12px 0;
  padding:10px 13px;border:1px solid var(--line);border-radius:12px;
  background:var(--panel);font-size:13px}
.ntbar-has{border-color:rgba(247,190,74,.75);background:rgba(247,190,74,.13)}
.ntbar-act{border-color:var(--bad);background:rgba(255,99,99,.13)}
.ntopen{color:var(--bad)}
.chlist{flex:0 0 100%;border-top:1px solid var(--line);padding-top:8px;margin-top:2px}
.chrow{display:flex;justify-content:space-between;gap:10px;padding:3px 0;font-size:13px;color:var(--dim)}
.chrow b{color:var(--text)}
.chrow b.chstop{color:var(--accent)}
.clrow{display:block}
.clq{color:var(--dim);font-size:14px}
.chlink{white-space:nowrap;font-weight:600}
.ntgone{background:var(--panel2);border-style:dashed}
.ntgone .nthead b{color:var(--dim);font-weight:600}
.ntorig{margin-top:8px;padding:9px 11px;border:1px solid var(--line);border-radius:10px;background:var(--panel)}
.nted{margin-top:6px;font-size:12px}
.nted-btns{display:flex;gap:8px;margin-top:8px}
.nthist{margin-top:6px;padding:8px 10px;border-left:3px solid var(--line);background:var(--panel2);border-radius:0 8px 8px 0}
.nthrow{padding:4px 0}
.nthrow .ntbody{font-size:14px;color:var(--dim)}
.ntmine{display:flex;gap:16px;margin-top:9px;font-size:13px;font-weight:600}
/* touch: text links get a 44px-tall hit area without changing the layout */
.chlink,.nted .link,.ntfoot .link,.ntwarn .link,.ntmine .link,.ntshow{display:inline-block;padding:12px 8px;margin:-12px -8px}
.ntmine{gap:18px}
/* check history + notes read as one two-row block, not two stacked boxes */
.clpanel.joined{margin-bottom:0;border-bottom-left-radius:0;border-bottom-right-radius:0}
.ntbar.joined{margin-top:0;border-top-width:0;border-top-left-radius:0;border-top-right-radius:0;margin-bottom:10px}
.joined .cllabel{display:none}
.joined .clline{flex:1 1 230px;min-width:0}
.ntprevs{flex:0 0 100%;border-top:1px solid rgba(255,99,99,.35);padding-top:8px}
.ntprev{font-size:15px;line-height:1.35;color:var(--text);cursor:pointer;margin-bottom:6px;
  display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.ntprev b{color:var(--bad)}
.copybtn.quiet{display:block;width:auto;margin:0 0 6px auto;padding:0 12px;min-height:44px;
  font-size:13px;background:transparent;border:0;color:var(--accent)}
.opcd summary{cursor:pointer;color:var(--warn);font-weight:600;min-height:44px;display:flex;align-items:center}
.opcl{padding:0 0 8px}
.appinfo{margin:18px 0 6px;text-align:center}
.appinfo summary{cursor:pointer;color:var(--faint);font-size:12px;min-height:44px;
  display:flex;align-items:center;justify-content:center}
.ntform{margin-bottom:12px}
.ntext{display:block;width:100%;min-height:86px;resize:vertical;border:1px solid var(--line);
  border-radius:12px;background:var(--panel);color:var(--text);font:inherit;font-size:16px;
  line-height:1.4;padding:10px 12px;outline:none}
.ntext:focus{border-color:var(--accent)}
.ntchk{display:flex;align-items:center;gap:8px;margin:9px 2px 11px;font-size:14px}
.ntchk input{width:20px;height:20px}
.ntitem{background:var(--panel);border:1px solid var(--line);border-radius:12px;
  padding:11px 13px;margin-bottom:9px}
.ntitem-act{border-color:var(--bad)}
.nthead{display:flex;justify-content:space-between;gap:10px;font-size:12px;
  color:var(--faint);margin-bottom:5px}
.nthead b{color:var(--text);font-size:13px}
.ntbody{white-space:pre-wrap;word-break:break-word;font-size:15px;line-height:1.4}
.ntfoot{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:9px}
.ntfoot .clbtn{margin-left:auto}
.faintx{color:var(--faint);font-size:12px}
.ntwarn{margin:0 0 12px;padding:10px 13px;border:1px solid var(--warn);border-radius:12px;
  background:rgba(247,190,74,.13);font-size:13px;line-height:1.45}
.clbtn{flex:0 0 auto;border:1px solid var(--line);background:var(--panel2);
  color:var(--text);border-radius:10px;padding:0 14px;min-height:44px;font-size:14px;
  font-weight:600;cursor:pointer}
.clbtn:active{background:var(--line)}
.t-Shuttle{background:rgba(63,203,126,.14);color:var(--shuttle)}
.chev{color:var(--faint);font-size:18px}

/* status badges */
.badge{display:inline-flex;align-items:center;gap:6px;border-radius:7px;
  padding:6px 10px;font-size:14px;font-weight:700}
.b-ok{background:rgba(63,203,126,.12);color:var(--ok)}
.b-warn{background:rgba(245,185,66,.14);color:var(--warn)}
.b-bad{background:rgba(255,93,93,.14);color:var(--bad)}
.b-na{background:var(--panel2);color:var(--faint)}
.badge b{font-weight:800;font-size:15px;letter-spacing:.03em}

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
  font-weight:600;cursor:pointer;white-space:normal;text-align:center;min-height:44px;
  display:inline-flex;align-items:center;justify-content:center}
.fchip.on{background:var(--accent);border-color:var(--accent);color:#fff}
.browsebar{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:10px}
.browsebar select{flex:1 1 160px;min-width:0;height:42px;border:1px solid var(--line);
  border-radius:10px;background:var(--panel);color:var(--text);font-size:14px;
  padding:0 32px 0 10px;-webkit-appearance:none;appearance:none;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 12 8'%3E%3Cpath d='M1 2l5 5 5-5' fill='none' stroke='%2371808F' stroke-width='1.6' stroke-linecap='round'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right 11px center;background-size:11px 7px}
.cardbadges{display:flex;flex-wrap:wrap;gap:5px;margin-top:6px}
.cardbadges .badge{padding:3px 8px;font-size:12px}
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

/* due date sheet: slides over the current page, so the record underneath is kept */
.sheetbg{position:fixed;inset:0;z-index:30;background:rgba(0,0,0,.45);
  display:flex;align-items:flex-end;justify-content:center}
.sheet{width:100%;max-width:640px;max-height:92vh;overflow-y:auto;background:var(--bg);
  border-top:1px solid var(--line);border-radius:16px 16px 0 0;
  padding:16px 16px calc(16px + env(safe-area-inset-bottom))}
.dueval{font-size:28px;font-weight:700;line-height:1.2;margin:2px 0 4px}
.duesub{color:var(--dim);font-size:13px;margin-bottom:14px}
.duelbl{display:block;font-size:12px;font-weight:700;letter-spacing:.06em;
  text-transform:uppercase;color:var(--faint);margin-bottom:14px}
/* Phones draw a date field with their own styling (centred text, extra height,
   a different width). Reset it so it matches the search box, with the date centred. */
.duelbl input{display:block;width:100%;min-width:0;max-width:100%;height:48px;margin-top:6px;
  -webkit-appearance:none;appearance:none;border:1px solid var(--line);border-radius:12px;
  background:var(--panel);color:var(--text);font:inherit;font-size:17px;font-weight:400;
  letter-spacing:normal;text-transform:none;text-align:center;line-height:46px;padding:0 12px}
.duelbl input::-webkit-date-and-time-value{text-align:center;margin:0;min-height:1.2em}
.duelbl input::-webkit-calendar-picker-indicator{margin:0}
.duenote{color:var(--faint);font-size:12px;text-align:center;padding:2px 0 10px}
.geolbl{font-size:11px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;
  color:var(--faint);margin-top:12px}
.geoval{font-size:20px;font-weight:700;line-height:1.25;white-space:pre-line}
.geosub{color:var(--dim);font-size:13px}
.geoacc{display:flex;align-items:center;gap:8px;font-size:14px;font-weight:600}
.geodot{width:10px;height:10px;border-radius:50%;background:var(--faint);flex:0 0 10px}
.geodot.good{background:var(--ok)}.geodot.fair{background:var(--warn)}.geodot.weak{background:var(--bad)}
.geostatus{color:var(--dim);font-size:13px;text-align:center;padding:2px 0 10px}
</style>
</head>
<body>
<div id="app">
  <header>
    <div class="hrow">
      <button class="navbtn" id="back" onclick="history.back()" style="display:none">&#8592;</button>
      <h1>PVH Field Lookup<small id="stamp"></small></h1>
      <button class="navbtn" id="duebtn" title="Due date" aria-label="Due date"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/></svg></button>
      <button class="navbtn" id="locbtn" title="Location" aria-label="Location"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="7"/><path d="M12 2v4M12 18v4M2 12h4M18 12h4"/><circle cx="12" cy="12" r="1.5" fill="currentColor"/></svg></button>
      <button class="navbtn" id="theme" title="Toggle dark mode">&#9789;</button>
    </div>
    <div id="searchwrap">
      <input id="q" type="search" placeholder="Deck # / plate / name / licence&#8230;"
        autocomplete="off" autocorrect="off" autocapitalize="characters" spellcheck="false">
      <button id="clr">&#10005;</button>
    </div>
  </header>
  <div id="toast" style="display:none"></div>
  <div class="sheetbg" id="duesheet" style="display:none">
    <div class="sheet" role="dialog" aria-label="Due date">
      <div class="seclabel" style="margin-top:0">Due date</div>
      <div class="dueval" id="dueval"></div>
      <div class="duesub" id="duesub"></div>
      <label class="duelbl">Ticket issued<input type="date" id="dueissued"></label>
      <div class="duenote">Check the courts' closure list if the date is close to a holiday.</div>
      <button class="copybtn" id="dueclose" style="color:var(--dim)">Close</button>
    </div>
  </div>
  <div class="sheetbg" id="geosheet" style="display:none">
    <div class="sheet" role="dialog" aria-label="Location">
      <div class="seclabel" style="margin-top:0">Location</div>
      <div class="geostatus" id="geostatus"></div>
      <button class="copybtn" id="geobtn" style="display:none">Try again</button>
      <div id="geobox" style="display:none">
        <div class="geoacc"><span class="geodot" id="geodot"></span><span id="geoacc"></span></div>
        <div class="geolbl">Street</div>
        <div class="geoval" id="geostreet"></div>
        <div class="geosub" id="geostreetsub"></div>
        <div class="geolbl">Nearest address</div>
        <div class="geoval" id="geocivic"></div>
        <div class="geosub" id="geocivicsub"></div>
        <div class="geolbl">Cross streets</div>
        <div class="geoval" id="geocs"></div>
      </div>
      <div class="duenote" id="geonote" style="display:none;padding-top:14px">Contains information licensed under the Open Government Licence &ndash; Nova Scotia. Field aid only, not a legal record.</div>
      <button class="copybtn" id="geoclose" style="color:var(--dim)">Close</button>
    </div>
  </div>
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
/* Several operators can share an owner's name (up to five on one name in
   the current data). The link carries the licence number so the one shown is
   identifiable, and an unresolved shared name lists every candidate rather
   than picking one -- the builder only links when the address confirms it. */
function opLink(j){
  const o=O[j];
  return '<span class="link" onclick="go(\'o/'+j+'\')">'+esc(opName(o))+
    (o["Licence Number"]?' ('+esc(o["Licence Number"])+')':'')+' &#8250;</span>';
}
function opcNote(c){
  return '<details class="opcd"><summary>Possible match \u2014 '+c.length+' operators share this name</summary>'+
    '<div class="opcl"><span class="empty">Verify manually:</span><br>'+c.map(opLink).join("<br>")+'</div></details>';
}
function opcText(c){
  return "POSSIBLE OWNER-OPERATOR \u2014 "+c.length+" operators share this name, verify: "+
    c.map(j=>O[j]["Licence Number"]||"?").join(", ");
}

function vehCard(v,extra){
  return '<div class="card" onclick="go(\'v/'+v._i+'\')">'+deckHTML(v)+
    '<div class="cmain"><div class="cname">'+esc(v["Make Model"]||"Vehicle")+
    (v["Vehicle Color"]?' \u00B7 '+esc(v["Vehicle Color"]):'')+'</div>'+
    '<div class="csub">'+esc(ownerName(v))+(v["Business Name"]?' \u00B7 '+esc(v["Business Name"]):'')+'</div>'+
    '<div class="crow2">'+ntChip(clKeyV(v))+'<span class="chip t-'+esc(v["Vehicle Type"])+'">'+esc(v["Vehicle Type"])+'</span>'+
    (v["Plate No"]?'<span class="plate">'+esc(v["Plate No"])+'</span>':'')+
    clChipsV(v)+
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

   This is the device's own copy. With shared notes signed in, stops and
   queries are also sent to the other officers (see clMarkK); this log is what
   keeps working offline and with sharing off. */
var CLOG_KEY="pvh_check_log", CLOG_DAYS=180, CLOG_MAX=900, CLOG=null;
function clKey(o){
  var s=String(o["Master Number"]==null?"":o["Master Number"]).trim().toUpperCase();
  s=s.replace(/^4A\b/,"").replace(/^0*4\s+/,"").replace(/\s+/g,"");
  return "o:"+(s||("#"+(o["ID"]==null?"?":o["ID"])));
}
/* Vehicles key on Licence No -- the CBRM-assigned TO/LTO/LO number, and the
   only vehicle field that holds up: present on 344 of 345 and never
   duplicated. Deck numbers are blank on a third of the fleet and 29 plates
   are shared between two records, so keying on either would file one
   vehicle's history under another. The o:/v: prefixes keep the two sets
   apart in one log. */
function clKeyV(v){
  var s=String(v["Licence No"]==null?"":v["Licence No"]).trim().toUpperCase().replace(/\s+/g,"");
  return "v:"+(s||("#"+(v["Vehicle ID"]==null?"?":v["Vehicle ID"])));
}
function clLoad(){
  if(CLOG)return CLOG;
  try{
    var r=localStorage.getItem(CLOG_KEY), o=r?JSON.parse(r):{};
    CLOG=(o&&typeof o==="object"&&!Array.isArray(o))?o:{};
  }catch(e){CLOG={};}
  return CLOG;
}
/* Entries are [lastQueried, lastStopped, lastQuerySent] in epoch seconds, 0
   for never. Pruned
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
/* Two levels, both shared when signed in. "Queried" (seen) is logged whenever
   a record is opened -- someone looked, for whatever reason. "Stopped"
   (checked) is the deliberate button for an actual stop. Each is the later of
   this device's own entry and the shared one, with the name of whoever made
   it. */
function clGetK(k){
  var e=clLoad()[k], sh=SC.checks[k], sq=SC.queries[k], me=(AU&&AU.name)||"";
  var lc=(e&&e[1])||0, st=(sh&&sh.t)||0, by="";
  if(st&&st>=lc)by=sh.by||"";
  else if(lc)by=me;
  var ls=(e&&e[0])||0, qt=(sq&&sq.t)||0, qby="";
  if(qt&&qt>=ls)qby=sq.by||"";
  else if(ls)qby=me||"you";
  return {k:k, seen:Math.max(ls,qt), qby:qby, checked:Math.max(lc,st), by:by};
}
/* A query is sent at most once per record every QUERY_GAP seconds from this
   device, so flicking back and forth during one encounter logs it once. The
   time of the last one sent is the entry's third slot. */
var QUERY_GAP=1800;
function clMarkK(k,deliberate,label){
  var m=clLoad(), e=m[k]||[0,0,0], now=Math.floor(Date.now()/1000);
  e[deliberate?1:0]=now;
  var sendQ=!deliberate&&shOn()&&!!AU&&now-(e[2]||0)>=QUERY_GAP;
  if(sendQ)e[2]=now;
  m[k]=e; clSave();
  if(!shOn()||!AU||!(deliberate||sendQ))return;
  OB.push({type:deliberate?"check":"query",row:{id:uid(),rec_key:k,rec_label:String(label||"").slice(0,200),
    written_at:new Date(now*1000).toISOString()}});
  obSave(); shSync(true);
}
/* Called by the record pages. The query panel reports who looked BEFORE this
   visit -- otherwise it would only ever show the look being taken now -- and
   that snapshot is held while the page is redrawn (a sync, a Stop tap), so the
   officer's own query does not replace it mid-visit. A visit is a fresh route,
   set by route(); a redraw of the same page is not one. */
var VS=null, FRESH=false;
function visitOpen(k,label){
  if(FRESH||!VS||VS.k!==k){
    var p=clGetK(k); VS={k:k,seen:p.seen,qby:p.qby};
    clMarkK(k,false,label);
  }
  var s=clGetK(k); s.seen=VS.seen; s.qby=VS.qby;
  return s;
}
/* 10/09/26 @ 1535Hrs -- the format officers write in their notebooks. */
function fmtWhen(t){
  var d=new Date(t*1000), p=function(n){return (n<10?"0":"")+n;};
  return p(d.getMonth()+1)+"/"+p(d.getDate())+"/"+p(d.getFullYear()%100)+" @ "+p(d.getHours())+p(d.getMinutes())+"Hrs";
}
/* An owner and an operator record can be the same human -- 164 of 200 owners
   link to one. Where they do, the owner's history IS the driver's history, so
   the key is the operator's: mark either page and both agree. The remaining 36
   are companies bar one blank artifact, and get their own key. Owner ID alone
   is not enough for those -- ID 1 is reused across two identities, which the
   build already warns about -- so the name goes into the key too.

   This is not the vehicle case. A car may be out with a different driver, so
   vehicle history stays separate from everyone's. */
function clKeyW(w){
  if(w._op!=null&&O[w._op])return clKey(O[w._op]);
  return nkW(w);
}
/* The owner's own key, never the operator's. Notes use this directly so a note
   on an owner stays on the owner even when the same person is also an operator. */
function nkW(w){
  var id=String(w["Owner ID"]==null?"":w["Owner ID"]).trim();
  var nm=String((w["Owner Last Name"]||"")+(w["Owner First Name"]||"")+(w["Business Name"]||""))
           .toUpperCase().replace(/[^A-Z0-9]/g,"");
  return "w:"+id+"|"+nm;
}
function clGet(o){return clGetK(clKey(o));}
function clMark(o,deliberate){clMarkK(clKey(o),deliberate,opName(o));}
function clGetV(v){return clGetK(clKeyV(v));}
function clMarkV(v,deliberate){clMarkK(clKeyV(v),deliberate,vehLabel(v));}
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
/* Both show: a recent query does not replace an older stop. Actual date and
   time, and who, rather than "1d" -- an officer needs to know whether it was
   this shift or the one before. */
function clChipsFor(s){
  return (s.checked?'<span class="chip clchk clwhen">Stopped '+esc(fmtWhen(s.checked))+(s.by?' \u00B7 '+esc(s.by):'')+'</span>':'')+
    (s.seen?'<span class="chip clseen clwhen">Queried '+esc(fmtWhen(s.seen))+(s.qby?' \u00B7 '+esc(s.qby):'')+'</span>':'');
}
function clChips(o){return clChipsFor(clGet(o));}
function clChipsV(v){return clChipsFor(clGetV(v));}
function clGetW(w){return clGetK(clKeyW(w));}
function clMarkW(w,deliberate){clMarkK(clKeyW(w),deliberate,ownerName(w));}
function clChipsW(w){return clChipsFor(clGetW(w));}

/* ---------- Shared notes and checks ----------
   Notes written on a record, "Stopped" marks and "Queried" looks are shared
   between officers through a Supabase project. Everything else --
   search and the records themselves -- stays on the device.

   Offline first, same rule as the rest of the app. Reads come from a copy kept
   on the device; anything an officer writes goes into an outbox on the device
   first and is sent whenever there is a signal and a valid sign-in. Nothing the
   officer does waits on the network.

   What is trusted where. The database stamps the author and the receive time
   itself (see supabase/schema.sql); the app only supplies when the officer
   wrote it, which the database clamps to "not in the future". So a note always
   carries the signed-in officer's real name, whatever this code sends.

   Notes are keyed per record and are NOT merged across an owner and the
   operator who is the same person -- a note appears only where it was written.
   Check history is the opposite on purpose (see clKeyW). The keys are the same
   stable identifiers the check history uses, never array positions, which
   shift on every rebuild. */
var SH=(DB.shared&&DB.shared.url&&DB.shared.key)?DB.shared:null;
var SC_KEY="pvh_shared_cache", OB_KEY="pvh_outbox", AU_KEY="pvh_auth";
function jget(k,d){try{var r=localStorage.getItem(k);return r?JSON.parse(r):d;}catch(e){return d;}}
function jset(k,v){try{localStorage.setItem(k,JSON.stringify(v));return true;}catch(e){return false;}}
function emptySC(){return {notes:{},checks:{},hist:{},ev:{},queries:{},qhist:{},nSince:"",cSince:"",qSince:"",at:0};}
var SC=jget(SC_KEY,null);
if(!SC||typeof SC!=="object"||!SC.notes||!SC.checks)SC=emptySC();
/* A cache written before the check list existed has only each record's latest
   check. Start its check pull from scratch so the list backfills. */
if(!SC.hist){SC.hist={}; SC.cSince="";}
if(!SC.ev){SC.ev={}; SC.eSince="";}
if(!SC.queries){SC.queries={}; SC.qhist={}; SC.qSince="";}
var OB=jget(OB_KEY,[]); if(!Array.isArray(OB))OB=[];
var AU=jget(AU_KEY,null);
var SYNC={busy:false,err:"",lastTry:0};
var NIDX=null;
function shOn(){return !!SH;}
function shSigned(){return !!(AU&&AU.refresh);}
function scSave(){
  var cut=Date.now()/1000-CLOG_DAYS*86400;
  [SC.checks,SC.queries].forEach(function(m){
    Object.keys(m).forEach(function(k){if((m[k].t||0)<cut)delete m[k];});
  });
  [SC.hist,SC.qhist].forEach(function(m){
    Object.keys(m).forEach(function(k){
      m[k]=m[k].filter(function(x){return x.t>=cut;});
      if(!m[k].length)delete m[k];
    });
  });
  NIDX=null; jset(SC_KEY,SC);
}
function obSave(){
  NIDX=null;
  if(!jset(OB_KEY,OB))alert("This device is out of storage, so that could not be saved. Free some space and try again.");
}
function auSave(){jset(AU_KEY,AU);}
function tsParse(s){
  /* Safari is strict about fractional seconds beyond milliseconds, which
     Postgres sends as microseconds. */
  var t=Date.parse(String(s||"").replace(/(\.\d{3})\d+/,"$1"));
  return isNaN(t)?0:t/1000;
}
function uid(){
  if(window.crypto&&crypto.randomUUID)return crypto.randomUUID();
  var b=new Uint8Array(16);
  if(window.crypto&&crypto.getRandomValues)crypto.getRandomValues(b);
  else for(var i=0;i<16;i++)b[i]=Math.random()*256;
  b[6]=(b[6]&15)|64; b[8]=(b[8]&63)|128;
  var h=Array.prototype.map.call(b,function(x){return (x+256).toString(16).slice(1);}).join("");
  return h.slice(0,8)+"-"+h.slice(8,12)+"-"+h.slice(12,16)+"-"+h.slice(16,20)+"-"+h.slice(20);
}

/* --- network. Only a signed-in user's token is ever sent as Authorization;
   the public key goes in apikey alone, which is what the newer publishable
   keys require. A reject with {net:true} means no usable connection. */
function shHttp(method,path,body,token,prefer){
  var h={"apikey":SH.key,"Content-Type":"application/json"};
  if(token)h.Authorization="Bearer "+token;
  if(prefer)h.Prefer=prefer;
  return fetch(SH.url+path,{method:method,headers:h,cache:"no-store",
      body:body==null?undefined:JSON.stringify(body)})
    .then(function(r){
      return r.text().then(function(t){
        var j=null; try{j=t?JSON.parse(t):null;}catch(e){}
        return {status:r.status,json:j};
      });
    },function(){throw {net:true};});
}
var REFRESHING=null;
function shRefresh(){
  if(REFRESHING)return REFRESHING;
  REFRESHING=shHttp("POST","/auth/v1/token?grant_type=refresh_token",{refresh_token:AU.refresh}).then(function(r){
    REFRESHING=null;
    if(r.status===200&&r.json&&r.json.access_token){
      AU.access=r.json.access_token; AU.refresh=r.json.refresh_token||AU.refresh;
      AU.exp=Math.floor(Date.now()/1000)+(r.json.expires_in||3600); AU.dead=false; auSave();
      return true;
    }
    if(r.status>=400&&r.status<500&&r.status!==408&&r.status!==429){AU.dead=true; auSave(); return false;}
    throw {net:true};
  },function(e){REFRESHING=null; throw e;});
  return REFRESHING;
}
function shApi(method,path,body,prefer){
  if(!AU||AU.dead)return Promise.reject({auth:true});
  function go(){return shHttp(method,path,body,AU.access,prefer);}
  var pre=((AU.exp||0)-60<Date.now()/1000)?shRefresh():Promise.resolve(true);
  return pre.then(function(ok){
    if(!ok)throw {auth:true};
    return go();
  }).then(function(r){
    if(r.status!==401)return r;
    return shRefresh().then(function(ok){if(!ok)throw {auth:true}; return go();});
  });
}
function shSignIn(email,pw){
  return shHttp("POST","/auth/v1/token?grant_type=password",{email:email,password:pw}).then(function(r){
    if(r.status!==200||!r.json||!r.json.access_token){
      var m=(r.json&&(r.json.error_description||r.json.msg||r.json.message))||("HTTP "+r.status);
      throw new Error(/invalid|credentials/i.test(m)?"Email or password not recognised.":m);
    }
    var a={access:r.json.access_token,refresh:r.json.refresh_token,
           exp:Math.floor(Date.now()/1000)+(r.json.expires_in||3600),
           uid:r.json.user&&r.json.user.id,email:email,name:""};
    return shHttp("GET","/rest/v1/profiles?select=display_name&id=eq."+encodeURIComponent(a.uid),null,a.access).then(function(p){
      if(p.status!==200||!p.json||!p.json.length)
        throw new Error("Signed in, but this account has not been enabled for PVH notes. Ask the administrator.");
      a.name=p.json[0].display_name;
      return a;
    });
  });
}

/* --- outbox. Ops: note, check, done. Safe to retry: every insert carries a
   client-made id, and a duplicate (409) means it already arrived. */
function shFlush(){
  var i=0;
  function next(){
    while(i<OB.length&&OB[i].failed)i++;
    if(i>=OB.length)return Promise.resolve();
    var op=OB[i], req;
    if(op.type==="note")req=shApi("POST","/rest/v1/notes",op.row,"return=minimal");
    else if(op.type==="check")req=shApi("POST","/rest/v1/checks",op.row,"return=minimal");
    else if(op.type==="query")req=shApi("POST","/rest/v1/queries",op.row,"return=minimal");
    else if(op.type==="edit")req=shApi("PATCH","/rest/v1/notes?id=eq."+encodeURIComponent(op.noteId),
      {body:op.body},"return=minimal");
    else if(op.type==="remove")req=shApi("PATCH","/rest/v1/notes?id=eq."+encodeURIComponent(op.noteId),
      {removed_at:new Date(op.at*1000).toISOString()},"return=minimal");
    else req=shApi("PATCH","/rest/v1/notes?id=eq."+encodeURIComponent(op.noteId),
      {done_at:op.done?new Date(op.at*1000).toISOString():null},"return=minimal");
    return req.then(function(r){
      if((r.status>=200&&r.status<300)||r.status===409){OB.splice(i,1); obSave(); return next();}
      if(r.status===401||r.status===403)throw {auth:true,refused:true};
      if(r.status>=500||r.status===408||r.status===429)throw {net:true};
      /* A query is a passing record of a look, not something an officer wrote.
         One the server will not take (say the queries table has not been
         added yet) is dropped rather than parked as "refused". */
      if(op.type==="query"){OB.splice(i,1); obSave(); return next();}
      op.failed=(r.json&&(r.json.message||r.json.hint))||("HTTP "+r.status);
      obSave(); CHG=true; i++; return next();
    });
  }
  return next();
}

/* --- pull. Incremental by a server timestamp, a page at a time. The boundary
   row is fetched again each round (gte), which is harmless because rows are
   merged by id / record key. */
var CHG=false, CH_KEEP=50, CH_SHOW=5;
function pullPages(table,col,since,take){
  function page(s){
    return shApi("GET","/rest/v1/"+table+"?select=*&order="+col+".asc&limit=1000&"+col+"=gte."+encodeURIComponent(s))
      .then(function(r){
        if(r.status===401||r.status===403)throw {auth:true,refused:true};
        if(r.status!==200||!Array.isArray(r.json))throw new Error("The server answered HTTP "+r.status+".");
        var rows=r.json;
        if(!rows.length)return s;
        take(rows);
        var last=rows[rows.length-1][col];
        if(rows.length<1000||last===s)return last;
        return page(last);
      });
  }
  return page(since);
}
function pullNotes(){
  return pullPages("notes","updated_at",SC.nSince||"1970-01-01T00:00:00Z",function(rows){
    rows.forEach(function(n){
      var c=SC.notes[n.id];
      if(!c||c.updated_at!==n.updated_at)CHG=true;
      SC.notes[n.id]=n;
    });
  }).then(function(s){SC.nSince=s;});
}
/* Checks (stops) and queries arrive the same way: the latest per record for
   the chips, and a capped newest-first list per record for the history. */
function mergeMarks(rows,latest,hist){
  rows.forEach(function(c){
    var t=tsParse(c.written_at), cur=latest[c.rec_key];
    if(!cur||t>cur.t){latest[c.rec_key]={t:t,by:c.author_name||""}; CHG=true;}
    var h=hist[c.rec_key]||(hist[c.rec_key]=[]);
    if(!h.some(function(x){return x.i===c.id;})){
      h.push({i:c.id,t:t,by:c.author_name||""});
      h.sort(function(a,b){return b.t-a.t;});
      if(h.length>CH_KEEP)h.length=CH_KEEP;
      CHG=true;
    }
  });
}
function pullChecks(){
  var since=SC.cSince||new Date(Date.now()-CLOG_DAYS*864e5).toISOString();
  return pullPages("checks","created_at",since,function(rows){mergeMarks(rows,SC.checks,SC.hist);})
    .then(function(s){SC.cSince=s;});
}
/* Never allowed to fail a sync: until 003_queries.sql has been run there is
   no such table, and everything else still works. */
function pullQueries(){
  var since=SC.qSince||new Date(Date.now()-CLOG_DAYS*864e5).toISOString();
  return pullPages("queries","created_at",since,function(rows){mergeMarks(rows,SC.queries,SC.qhist);})
    .then(function(s){SC.qSince=s;},function(e){if(e&&(e.net||e.auth))throw e;});
}
/* Earlier text of edited notes. Never allowed to fail a sync: until the
   database has been upgraded there is no such table, and notes still work. */
function pullEvents(){
  return pullPages("note_events","created_at",SC.eSince||"1970-01-01T00:00:00Z",function(rows){
    rows.forEach(function(e){
      if(e.kind!=="edit")return;
      var l=SC.ev[e.note_id]||(SC.ev[e.note_id]=[]);
      if(l.some(function(x){return x.i===e.id;}))return;
      l.push({i:e.id,t:tsParse(e.created_at),by:e.by_name||"",o:e.old_body||""});
      l.sort(function(a,b){return a.t-b.t;});
      CHG=true;
    });
  }).then(function(s){SC.eSince=s;},function(e){if(e&&(e.net||e.auth))throw e;});
}
function typing(){
  var a=document.activeElement, id=a&&a.id;
  if(id==="ntbody"||id==="nteditbody"||id==="shem"||id==="shpw")return true;
  if(document.getElementById("nteditbody"))return true;
  var t=document.getElementById("ntbody");
  return !!(t&&t.value);
}
function keepScroll(){var y=window.scrollY; route(); window.scrollTo(0,y);}
function shSync(force){
  if(!shOn()||!shSigned()||SYNC.busy)return Promise.resolve();
  if(!force&&Date.now()-SYNC.lastTry<20000)return Promise.resolve();
  SYNC.busy=true; SYNC.lastTry=Date.now(); CHG=false;
  return shFlush().then(pullNotes).then(pullChecks).then(pullQueries).then(pullEvents).then(function(){
    SYNC.err=""; SC.at=Math.floor(Date.now()/1000); scSave();
  }).catch(function(e){
    if(e&&e.auth)SYNC.err=(AU&&AU.dead)?"Signed out — sign in again to send and receive.":
      "The server refused this account. Ask the administrator.";
    else if(e&&e.net)SYNC.err="No signal — will retry.";
    else SYNC.err=(e&&e.message)||"Sync failed.";
    scSave();
  }).then(function(){
    SYNC.busy=false;
    var h=location.hash.replace(/^#\/?/,"");
    if(h==="shared")renderShared();
    else if(CHG&&!typing())keepScroll();
    CHG=false;
  });
}

/* --- notes, indexed by record key. Pending notes (still in the outbox) are
   folded in so an officer sees their own note the moment they save it, and a
   done-tick is applied optimistically over the server's copy. */
function nIndex(){
  if(NIDX)return NIDX;
  var idx={}, over={}, me=(AU&&AU.name)||"", myId=(AU&&AU.uid)||"";
  OB.forEach(function(op){
    if(op.type==="done"||op.type==="edit"||op.type==="remove")(over[op.noteId]=over[op.noteId]||[]).push(op);
  });
  Object.keys(SC.notes).forEach(function(id){
    var n=SC.notes[id], v={id:id,body:n.body,action:!!n.action_needed,by:n.author_name,
      mine:!!myId&&n.author_id===myId,t:tsParse(n.written_at),
      done:n.done_at?tsParse(n.done_at):0,doneBy:n.done_by||"",
      edited:n.edited_at?tsParse(n.edited_at):0,editedBy:n.edited_by||"",
      removed:n.removed_at?tsParse(n.removed_at):0,removedBy:n.removed_by||"",
      ev:SC.ev[id]||[]};
    /* This device's unsent changes show straight away; one the server refused
       is not shown as if it had happened. */
    (over[id]||[]).forEach(function(o){
      if(o.failed){v.opFailed=o.failed; return;}
      if(o.type==="done"){v.done=o.done?o.at:0; v.doneBy=o.done?me:"";}
      else if(o.type==="edit"&&!v.removed){v.body=o.body; v.edited=o.at; v.editedBy=me; v.sending=true;}
      else if(o.type==="remove"&&!v.removed){v.removed=o.at; v.removedBy=me; v.sending=true;}
    });
    (idx[n.rec_key]=idx[n.rec_key]||[]).push(v);
  });
  OB.forEach(function(op){
    if(op.type!=="note")return;
    var r=op.row;
    (idx[r.rec_key]=idx[r.rec_key]||[]).push({id:r.id,body:r.body,action:!!r.action_needed,by:me,mine:true,
      t:tsParse(r.written_at),done:0,doneBy:"",edited:0,removed:0,ev:[],pending:true,failed:op.failed||""});
  });
  Object.keys(idx).forEach(function(k){idx[k].sort(function(a,b){return b.t-a.t;});});
  NIDX=idx; return idx;
}
function ntStat(key){
  var l=nIndex()[key]||[], n=0, open=0;
  l.forEach(function(x){
    if(x.removed)return;
    n++;
    if(x.action&&!x.done)open++;
  });
  return {n:n,open:open};
}
function ntChip(key){
  if(!shOn())return "";
  var s=ntStat(key);
  if(!s.n)return "";
  return s.open?'<span class="chip ntact">ACTION NEEDED</span>':'<span class="chip nt">NOTES '+s.n+'</span>';
}
function vehLabel(v){return ((v["Make Model"]||"Vehicle")+" "+(v["Plate No"]||v["Licence No"]||"")).trim();}
function recInfo(t,i){
  if(t==="v"){var v=V[i]; return v&&{key:clKeyV(v),type:"vehicle",label:vehLabel(v)};}
  if(t==="o"){var o=O[i]; return o&&{key:clKey(o),type:"operator",label:(opName(o)+" "+(o["Licence Number"]||"")).trim()};}
  var w=W[i]; return w&&{key:nkW(w),type:"owner",label:ownerName(w)};
}
function ntAdd(t,i){
  if(!shSigned())return;
  var ta=document.getElementById("ntbody"), body=((ta&&ta.value)||"").trim();
  if(!body){if(ta)ta.focus(); return;}
  var r=recInfo(t,i); if(!r)return;
  OB.push({type:"note",row:{id:uid(),rec_key:r.key,rec_type:r.type,rec_label:r.label.slice(0,200),
    body:body.slice(0,2000),action_needed:!!document.getElementById("ntact").checked,
    written_at:new Date().toISOString()}});
  obSave(); ta.value="";
  keepScroll(); shSync(true);
}
function ntDone(id,flag){
  if(!shSigned())return;
  OB.push({type:"done",noteId:id,done:!!flag,at:Math.floor(Date.now()/1000)});
  obSave(); keepScroll(); shSync(true);
}
/* "Read" lands on the section; "Add note" also puts the cursor in the box, so
   adding one is a single tap whether or not the record already has notes. */
function ntJump(write){
  var e=document.getElementById("ntsec"); if(!e)return;
  e.scrollIntoView({behavior:"smooth",block:"start"});
  var ta=document.getElementById("ntbody");
  if(ta&&write)setTimeout(function(){ta.focus();},350);
}
function ntWhen(t){
  var d=new Date(t*1000), o={month:"short",day:"numeric",hour:"numeric",minute:"2-digit"};
  if(d.getFullYear()!==new Date().getFullYear())o.year="numeric";
  try{return d.toLocaleString("en-CA",o);}catch(e){return d.toLocaleString();}
}
/* The strip under a record's header. Always rendered when sharing is on, for
   the same reason the check panel is: "no notes" is itself an answer, and a
   strip that comes and goes is one you have to hunt for. Turns amber with
   notes and red when one still needs action, so it reads at a glance. */
function ntBar(key){
  if(!shOn())return "";
  var s=ntStat(key), cls="ntbar"+(s.open?" ntbar-act":(s.n?" ntbar-has":""));
  var line=s.n?('<b>'+s.n+' note'+(s.n===1?'':'s')+'</b>'+
      (s.open?' · <b class="ntopen">'+s.open+' need'+(s.open===1?'s':'')+' action</b>':'')):'No notes yet';
  var openN=(nIndex()[key]||[]).filter(function(n){return n.action&&!n.done&&!n.removed;});
  var prev=openN.length?'<div class="ntprevs" onclick="ntJump(0)">'+openN.slice(0,2).map(function(n){
      return '<div class="ntprev"><b>'+esc(n.by||"")+':</b> '+esc(n.body)+'</div>';}).join("")+
    (openN.length>2?'<div class="faintx">+'+(openN.length-2)+' more needing action</div>':'')+'</div>':'';
  return '<div class="'+cls+' joined"><span class="cllabel">Notes</span><span class="clline">'+line+'</span>'+
    (s.n?'<button class="clbtn" onclick="ntJump(0)">Read</button>':'')+
    '<button class="clbtn" onclick="ntJump(1)">Add note</button>'+prev+'</div>';
}
var NT_EDIT="", NT_SHOW={}, NT_HIST={};
function ntFind(id){
  var idx=nIndex(), ks=Object.keys(idx);
  for(var i=0;i<ks.length;i++){
    for(var j=0;j<idx[ks[i]].length;j++)if(idx[ks[i]][j].id===id)return idx[ks[i]][j];
  }
  return null;
}
/* Earlier wordings of an edited note, newest first. Each is the text as it
   stood before the named person changed it. */
function ntHistBlock(n){
  var ev=(n.ev||[]).slice().reverse();
  if(!ev.length)return '<div class="nthist"><div class="faintx">The earlier wording is not on this device yet.</div></div>';
  return '<div class="nthist">'+ev.map(function(e){
    return '<div class="nthrow"><div class="faintx">Before the edit by '+esc(e.by||"?")+' · '+esc(ntWhen(e.t))+
      '</div><div class="ntbody">'+esc(e.o)+'</div></div>';}).join("")+'</div>';
}
function ntEditedLine(n){
  if(!n.edited)return "";
  return '<div class="nted"><span class="link" onclick="ntHist(\''+n.id+'\')">Edited by '+esc(n.editedBy||"?")+
    ' · '+esc(ntWhen(n.edited))+(n.sending?' · waiting to send':'')+
    (NT_HIST[n.id]?' · hide':' · see earlier')+'</span></div>'+(NT_HIST[n.id]?ntHistBlock(n):'');
}
function ntItem(n){
  if(n.removed){
    var open=NT_SHOW[n.id];
    return '<div class="ntitem ntgone"><div class="nthead"><b>Note removed</b><span>'+esc(n.removedBy||"?")+
      ' · '+esc(ntWhen(n.removed))+(n.sending?' · waiting to send':'')+'</span></div>'+
      '<div class="ntfoot"><span class="link" onclick="ntShow(\''+n.id+'\')">'+(open?'Hide original':'Show original')+'</span></div>'+
      (open?'<div class="ntorig"><div class="nthead"><b>'+esc(n.by||"")+'</b><span>'+esc(ntWhen(n.t))+'</span></div>'+
        '<div class="ntbody">'+esc(n.body)+'</div>'+ntEditedLine(n)+'</div>':'')+'</div>';
  }
  var st=n.failed?'<span class="badge b-bad">NOT SENT</span>':(n.pending?'<span class="chip clseen">waiting to send</span>':'');
  var act=n.action?(n.done?'<span class="chip ntdone">DONE'+(n.doneBy?' · '+esc(n.doneBy):'')+'</span>'
    :'<span class="chip ntact">ACTION NEEDED</span>'):'';
  var btn=(n.action&&!n.pending)?'<button class="clbtn" onclick="ntDone(\''+n.id+'\','+(n.done?'false':'true')+')">'+
    (n.done?'Reopen':'Mark done')+'</button>':'';
  var body=(NT_EDIT===n.id)
    ?'<textarea id="nteditbody" class="ntext" rows="4" maxlength="2000">'+esc(n.body)+'</textarea>'+
      '<div class="nted-btns"><button class="clbtn" onclick="ntEditSave(\''+n.id+'\')">Save change</button>'+
      '<button class="clbtn" onclick="ntEditCancel()">Cancel</button></div>'
    :'<div class="ntbody">'+esc(n.body)+'</div>';
  var mine=(n.mine&&NT_EDIT!==n.id)?'<div class="ntmine"><span class="link" onclick="ntEdit(\''+n.id+'\')">Edit</span>'+
    '<span class="link" onclick="ntRemove(\''+n.id+'\')">Remove</span></div>':'';
  return '<div class="ntitem'+((n.action&&!n.done)?' ntitem-act':'')+'">'+
    '<div class="nthead"><b>'+esc(n.by||"")+'</b><span>'+esc(ntWhen(n.t))+'</span></div>'+
    body+ntEditedLine(n)+
    (n.opFailed?'<div class="ntwarn">A change to this note was not accepted: '+esc(n.opFailed)+
      ' <span class="link" onclick="ntDismiss(\''+n.id+'\')">Dismiss</span></div>':'')+
    ((act||st||btn||n.failed)?'<div class="ntfoot">'+act+st+(n.failed?'<span class="faintx">'+esc(n.failed)+'</span>':'')+btn+'</div>':'')+
    mine+'</div>';
}
function ntShow(id){NT_SHOW[id]=!NT_SHOW[id]; keepScroll();}
/* Clears a refused change from this device so its notice goes away. The note
   itself is untouched. */
function ntDismiss(id){
  OB=OB.filter(function(o){return !(o.failed&&o.noteId===id);});
  obSave(); keepScroll();
}
function ntHist(id){NT_HIST[id]=!NT_HIST[id]; keepScroll();}
function ntEdit(id){
  NT_EDIT=id; keepScroll();
  setTimeout(function(){var t=document.getElementById("nteditbody"); if(t){t.focus(); t.setSelectionRange(t.value.length,t.value.length);}},60);
}
function ntEditCancel(){NT_EDIT=""; keepScroll();}
/* A note still in the outbox has not reached anyone, so editing or removing it
   just changes what will be sent. Only a note that has been sent leaves a
   trail -- there is nothing to audit before anyone could have seen it. */
function ntEditSave(id){
  var ta=document.getElementById("nteditbody"), body=((ta&&ta.value)||"").trim().slice(0,2000);
  if(!body){if(ta)ta.focus(); return;}
  var pend=OB.filter(function(o){return o.type==="note"&&o.row.id===id;})[0];
  if(pend)pend.row.body=body;
  else{
    var cur=ntFind(id);
    if(cur&&cur.body===body){NT_EDIT=""; keepScroll(); return;}
    OB.push({type:"edit",noteId:id,body:body,at:Math.floor(Date.now()/1000)});
  }
  obSave(); NT_EDIT=""; keepScroll(); shSync(true);
}
function ntRemove(id){
  if(!confirm("Remove this note?\n\nIt stays on file. Everyone can still see that it was removed, by whom and when, and can read the original."))return;
  var isPending=OB.some(function(o){return o.type==="note"&&o.row.id===id;});
  if(isPending)OB=OB.filter(function(o){return !(o.type==="note"&&o.row.id===id);});
  else OB.push({type:"remove",noteId:id,at:Math.floor(Date.now()/1000)});
  obSave(); keepScroll(); if(!isPending)shSync(true);
}
function ntSection(t,i,key){
  if(!shOn())return "";
  var l=nIndex()[key]||[];
  var h='<div class="seclabel" id="ntsec">Notes · shared with other officers ('+ntStat(key).n+')</div>';
  if(shSigned()){
    h+=(AU.dead?'<div class="ntwarn">Signed out — what you write is kept and will send once you '+
        '<span class="link" onclick="go(\'shared\')">sign in again</span>.</div>':'')+
      '<div class="ntform"><textarea id="ntbody" class="ntext" rows="3" maxlength="2000" '+
        'placeholder="Add a note to this record… what happened, what needs doing"></textarea>'+
      '<label class="ntchk"><input type="checkbox" id="ntact"> Action needed</label>'+
      '<button class="copybtn" onclick="ntAdd(\''+t+'\','+i+')">Save note</button></div>';
  }else{
    h+='<div class="notes">Sign in to read and add shared notes. '+
      '<span class="link" onclick="go(\'shared\')">Sign in…</span></div>';
  }
  return h+(l.length?l.map(ntItem).join(""):'<div class="empty">No notes on this record yet.</div>');
}

/* --- records that still have an action open, for the home screen. */
function actionRecs(){
  var idx=nIndex(), keys=Object.keys(idx).filter(function(k){
    return idx[k].some(function(n){return n.action&&!n.done&&!n.removed;});});
  if(!keys.length)return [];
  var by={};
  W.forEach(function(w,ix){by[nkW(w)]={t:"w",i:ix};});
  V.forEach(function(v,ix){by[clKeyV(v)]={t:"v",i:ix};});
  O.forEach(function(o,ix){by[clKey(o)]={t:"o",i:ix};});
  return keys.map(function(k){return by[k];}).filter(Boolean);
}
function actionCard(){
  if(!shOn())return "";
  var n=actionRecs().length;
  if(!n)return "";
  return '<div class="card" onclick="go(\'actions\')"><div class="opdot" style="color:var(--bad);border-color:var(--bad)">!</div>'+
    '<div class="cmain"><div class="cname">'+n+' record'+(n===1?'':'s')+' with open actions</div>'+
    '<div class="csub">Notes marked "action needed" that nobody has marked done</div></div>'+
    '<div class="chev">&#8250;</div></div>';
}
function renderActions(){
  var r=actionRecs();
  main.innerHTML='<div class="seclabel">Open actions ('+r.length+')</div>'+
    (r.length?r.map(function(e){return e.t==="o"?opCard(O[e.i]):(e.t==="v"?vehCard(V[e.i]):ownerCard(W[e.i]));}).join(""):
      '<div class="empty">Nothing is waiting on an action.</div>');
}

/* --- sign-in and status page. */
function agoShort(ts){
  var s=Math.floor(Date.now()/1000-ts);
  if(s<60)return "just now";
  if(s<3600)return Math.floor(s/60)+" min ago";
  return clAgo(ts);
}
function shStatus(){
  var p=[];
  if(SYNC.busy)p.push("Syncing…");
  else if(SYNC.err)p.push(esc(SYNC.err));
  else if(SC.at)p.push("Synced "+esc(agoShort(SC.at)));
  else p.push("Not synced yet");
  if(OB.length)p.push(OB.length+" waiting to send");
  return p.join(" · ");
}
function shLine(){
  if(!shOn())return "";
  var t=!shSigned()?'Shared notes: <span class="link" onclick="go(\'shared\')">sign in…</span>':
    (AU.dead?'Shared notes: <span class="link" onclick="go(\'shared\')">sign in again</span>'+(OB.length?' · '+OB.length+' waiting to send':''):
      'Shared notes · '+esc(AU.name)+(OB.length?' · '+OB.length+' waiting to send':'')+
      ' · <span class="link" onclick="go(\'shared\')">details</span>');
  return '<div class="hint" style="padding:10px 20px 0">'+t+'</div>';
}
function renderShared(){
  var h='<div class="seclabel">Shared notes and checks</div>';
  if(!shOn()){
    main.innerHTML=h+'<div class="notes">This data file has no shared-notes connection, so notes and shared checks are off. '+
      'The administrator turns them on in the build settings.</div>';
    return;
  }
  if(shSigned()&&!AU.dead){
    var failed=OB.filter(function(o){return o.failed;}).length;
    main.innerHTML=h+'<div class="srccard">Signed in as <b>'+esc(AU.name)+'</b><br>'+esc(AU.email||"")+'</div>'+
      '<div class="stamp">'+shStatus()+'</div>'+
      '<button class="copybtn" onclick="shNow()">Sync now</button>'+
      (failed?'<button class="copybtn" style="color:var(--bad)" onclick="shDiscard()">Discard '+failed+
        ' item'+(failed===1?'':'s')+' the server refused</button>':'')+
      '<button class="copybtn" style="color:var(--bad)" onclick="shOut()">Sign out of this device</button>';
    return;
  }
  main.innerHTML=h+(AU&&AU.dead
    ?'<div class="ntwarn">The saved sign-in no longer works. Sign in again'+(OB.length?' — the '+OB.length+
      ' item'+(OB.length===1?'':'s')+' waiting to send will go out afterwards':'')+'.</div>'
    :'<div class="notes">Sign in with the account the administrator set up for you. Notes, stops and queries are then '+
      'shared with the other officers. Signing in needs a signal once; after that the app works offline and sends '+
      'anything pending when it reconnects.</div>')+
    '<input id="shem" class="dsinput" type="email" inputmode="email" autocomplete="username" spellcheck="false" '+
      'placeholder="Email" value="'+esc((AU&&AU.email)||"")+'">'+
    '<input id="shpw" class="dsinput" type="password" autocomplete="current-password" placeholder="Password">'+
    '<button class="copybtn" id="shgo">Sign in</button><div class="stamp" id="shmsg"></div>';
  var msg=document.getElementById("shmsg"), pw=document.getElementById("shpw");
  function go2(){
    var em=document.getElementById("shem").value.trim();
    if(!em||!pw.value){msg.textContent="Enter your email and password."; return;}
    msg.textContent="Signing in…";
    shSignIn(em,pw.value).then(function(a){
      if(AU&&AU.uid&&AU.uid!==a.uid){OB=[]; obSave(); SC=emptySC(); scSave();}
      AU=a; auSave(); NIDX=null; pw.value="";
      shSync(true); renderShared();
    }).catch(function(e){
      msg.textContent=(e&&e.net)?"No connection. Signing in needs a signal.":((e&&e.message)||"Could not sign in.");
    });
  }
  document.getElementById("shgo").addEventListener("click",go2);
  pw.addEventListener("keydown",function(ev){if(ev.key==="Enter")go2();});
}
function shNow(){shSync(true).then(function(){if(location.hash.replace(/^#\/?/,"")==="shared")renderShared();}); renderShared();}
function shOut(){
  var n=OB.length;
  if(!confirm(n?(n+" item"+(n===1?"":"s")+" not yet sent will be lost. Sign out anyway?"):
      "Sign out of shared notes on this device? The shared notes kept here are removed."))return;
  AU=null; OB=[]; SC=emptySC(); NIDX=null;
  try{localStorage.removeItem(AU_KEY); localStorage.removeItem(OB_KEY); localStorage.removeItem(SC_KEY);}catch(e){}
  renderShared();
}
function shDiscard(){
  OB=OB.filter(function(o){return !o.failed;}); obSave(); renderShared();
}
/* One set of listeners for the life of the page, however many times boot()
   runs, always pointing at the latest closure. */
window.__shTick=function(){shSync(false);};
window.__shForce=function(){shSync(true);};
if(!window.__shWired){
  window.__shWired=true;
  document.addEventListener("visibilitychange",function(){
    if(document.visibilityState==="visible"&&window.__shTick)window.__shTick();});
  /* A regained signal always syncs: the throttle exists to stop chatter, and
     a failed attempt seconds ago is exactly when this event matters. */
  window.addEventListener("online",function(){if(window.__shForce)window.__shForce();});
  setInterval(function(){if(document.visibilityState==="visible"&&window.__shTick)window.__shTick();},180000);
}
setTimeout(function(){shSync(true);},400);
/* One panel, one position, every record type. Always rendered -- "not checked"
   is itself the answer to the question the officer is asking, and a panel that
   comes and goes is one you have to hunt for. */
/* Shown on both halves of an owner-operator's pair of search hits, so two
   cards for one human read as one person with two records rather than two
   separate people. */
function pairNote(arrow,where){
  return '<div class="pairnote"><b>'+arrow+'</b> Same person \u00B7 see '+where+'</div>';
}
/* Every shared stop and query on a record, newest first, plus this device's
   own that have not been sent yet. Nothing here is drawn until someone taps. */
var CH_OPEN={}, CH_KEYS=[];
function chList(k){
  var out=(SC.hist[k]||[]).map(function(x){return {t:x.t,by:x.by,stop:true,pending:false};})
    .concat((SC.qhist[k]||[]).map(function(x){return {t:x.t,by:x.by,stop:false,pending:false};}));
  OB.forEach(function(op){
    if((op.type==="check"||op.type==="query")&&op.row.rec_key===k)
      out.push({t:tsParse(op.row.written_at),by:(AU&&AU.name)||"",stop:op.type==="check",pending:true});
  });
  out.sort(function(a,b){return b.t-a.t;});
  return out;
}
function chToggle(n){
  var k=CH_KEYS[n]; if(k==null)return;
  CH_OPEN[k]=CH_OPEN[k]?0:CH_SHOW;
  keepScroll();
}
function chMore(n){
  var k=CH_KEYS[n]; if(k==null)return;
  CH_OPEN[k]=CH_KEEP;
  keepScroll();
}
function chBlock(s,n,l){
  var lim=CH_OPEN[s.k]; if(!lim)return "";
  var rows=l.slice(0,lim).map(function(x){
    return '<div class="chrow"><span><b class="'+(x.stop?'chstop':'')+'">'+(x.stop?'Stopped':'Queried')+'</b> '+esc(x.by||"?")+'</span>'+
      '<span>'+esc(fmtWhen(x.t))+(x.pending?' \u00B7 waiting to send':'')+'</span></div>';}).join("");
  return '<div class="chlist">'+rows+(l.length>lim?
    '<div class="chrow"><span class="link" onclick="chMore('+n+')">Show all ('+l.length+')</span></div>':'')+'</div>';
}
function clPanel(s,handler){
  var local=shOn()?"":" on this device";
  var line='<span class="clrow">'+(s.checked
      ?'<b>Stopped '+esc(fmtWhen(s.checked))+(s.by?' \u00B7 '+esc(s.by):'')+'</b>'
      :'No stop recorded'+local)+'</span>'+
    '<span class="clrow clq">'+(s.seen
      ?'Queried '+esc(fmtWhen(s.seen))+(s.qby?' \u00B7 '+esc(s.qby):'')
      :'Not queried before'+local)+'</span>';
  /* The list only exists with sharing on, and only earns a link once there is
     more than the latest stop and query the lines already state. */
  var l=shOn()?chList(s.k):[], n=-1, link="", blk="";
  if(l.length>(s.checked?1:0)+(s.seen?1:0)){
    n=CH_KEYS.indexOf(s.k); if(n<0){CH_KEYS.push(s.k); n=CH_KEYS.length-1;}
    link=' <span class="link chlink" onclick="chToggle('+n+')">'+(CH_OPEN[s.k]?'Hide':'History ('+l.length+')')+'</span>';
    blk=chBlock(s,n,l);
  }
  return '<div class="clpanel'+((s.checked&&clDays(s.checked)<14)?' clrecent':'')+(shOn()?' joined':'')+'">'+
    '<span class="cllabel">Stops &amp; queries</span>'+
    '<span class="clline">'+line+link+'</span>'+
    '<button class="clbtn" onclick="'+handler+'">'+(s.checked?'Record another stop':'Record a stop')+'</button>'+
    blk+
    '</div>';
}
function markChecked(i){
  var o=O[i]; if(!o)return;
  clMark(o,true);
  renderOperator(i);
}
function markCheckedV(i){
  var v=V[i]; if(!v)return;
  clMarkV(v,true);
  renderVehicle(i);
}
function markCheckedW(i){
  var w=W[i]; if(!w)return;
  clMarkW(w,true);
  renderOwner(i);
}
function clearCheckLog(){
  if(!confirm("Clear this device's own stop and query history? Anything already shared with other officers stays shared."))return;
  try{localStorage.removeItem(CLOG_KEY);}catch(e){}
  CLOG=null;
  renderHome(q.value);
}
function checkedCards(){
  var m=clLoad(), byKey={};
  O.forEach(function(o,ix){var k=clKey(o); if(byKey[k]==null)byKey[k]={t:"o",i:ix};});
  V.forEach(function(v,ix){var k=clKeyV(v); if(byKey[k]==null)byKey[k]={t:"v",i:ix};});
  /* Owners last: an owner-operator shares the driver's key, and the operator
     record is the more useful of the two to land on. */
  W.forEach(function(w,ix){var k=clKeyW(w); if(byKey[k]==null)byKey[k]={t:"w",i:ix};});
  var seenK={}, ks=Object.keys(m).concat(Object.keys(SC.checks)).filter(function(k){
    if(seenK[k]||!byKey[k]||!clGetK(k).checked)return false;
    seenK[k]=1; return true;});
  if(!ks.length)return "";
  ks.sort(function(a,b){return clGetK(b).checked-clGetK(a).checked;});
  /* RECENT already lists whatever was opened this session, directly above.
     Skip those so a record looked at a minute ago is not printed twice. */
  var rows=ks.map(function(k){
    var e=byKey[k];
    if(RECENT.indexOf(e.t+"/"+e.i)>-1)return "";
    return e.t==="o"?opCard(O[e.i]):(e.t==="v"?vehCard(V[e.i]):ownerCard(W[e.i]));
  }).filter(Boolean).slice(0,6).join("");
  if(!rows)return "";
  return '<div class="seclabel">Stopped recently</div>'+rows+
    '<div class="hint"><span class="link" onclick="clearCheckLog()">Clear this device\'s history</span></div>';
}

function opCard(o,extra){
  const init=((o["First Name"]||" ")[0]+(o["Last Name"]||" ")[0]).toUpperCase();
  const flag=(o["In Active"]===true||o["Cancelled"]===true);
  const ownop=(o._veh&&o._veh.length>0);
  return '<div class="card" onclick="go(\'o/'+o._i+'\')">'+
    '<div class="opdot">'+esc(init)+'</div>'+
    '<div class="cmain"><div class="cname">'+esc(opName(o))+'</div>'+
    '<div class="csub">'+dash(o["Business Name"])+'</div>'+
    '<div class="crow2">'+ntChip(clKey(o))+'<span class="chip t-'+esc(o["Operator Type"])+'">'+esc(o["Operator Type"])+' operator</span>'+
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
    '<div class="crow2">'+ntChip(nkW(w))+'<span class="chip" style="background:var(--panel2);color:var(--dim)">OWNER</span>'+
    '<span class="chip" style="background:var(--panel2);color:var(--dim)">'+n+' vehicle'+(n===1?"":"s")+'</span>'+
    (ownop?'<span class="chip" style="background:rgba(95,174,255,.16);color:var(--accent)">OWNER-OPERATOR</span>':'')+
    clChipsW(w)+
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
function chkInto(out,k,label,date){
  if(date==null||parseD(date)==null){out.push({k:k,label:label,date:null,miss:1,w:1e9});return;}
  const p=daysPast(date);
  if(p){out.push({k:k,label:label,date:date,over:p,w:p});return;}
}
function vehFlagList(v){
  const out=[];
  chkInto(out,"mvi","MVI",addYear(v["First MVIDate"]));
  chkInto(out,"permit","PVH License",v["Expiry Date"]);
  chkInto(out,"ins","Insurance",v["Insurance Expiry"]);
  chkInto(out,"vlic","NS Permit",v["NSVehicle Permit Expiry"]);
  return out;
}
/* Long lists draw a page at a time. A common surname can be
   hundreds of cards, which is slow to draw and a long way to scroll. */
var LIM={}, PAGE=25;
function pg(id,arr,fn){
  var n=LIM[id]||PAGE, h=arr.slice(0,n).map(fn).join("");
  if(arr.length>n){
    var left=arr.length-n, step=Math.min(50,left);
    h+='<button class="copybtn" onclick="pgMore(\''+id+'\')">Show '+step+' more'+(left>step?' ('+left+' left)':'')+'</button>';
  }
  return h;
}
function pgMore(id){LIM[id]=(LIM[id]||PAGE)+50; keepScroll();}
/* The header is sticky, so land the group heading just below it, not under it. */
function jumpTo(k){
  var el=document.getElementById("rh-"+k); if(!el)return;
  var hd=document.querySelector("header"), off=(hd?hd.offsetHeight:0)+6;
  window.scrollTo(0,el.getBoundingClientRect().top+window.scrollY-off);
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
    op?("OWNER-OPERATOR — also licensed as "+opName(op)+(op["Licence Number"]?" ("+op["Licence Number"]+")":"")):
      ((w._opc||[]).length?opcText(w._opc):"Not separately licensed as an operator"),
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
  const fl=vehFlagList(v).filter(x=>!x.miss);
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
    main.innerHTML='<div class="hint">Search by deck light number, plate, owner or operator name, business or licence number.</div>'+
      '<div class="browsebar">'+selectEl("fType",TYPE_OPTS,HFILTER.type)+selectEl("fStat",STAT_OPTS,HFILTER.stat)+'</div>'+
      actionCard()+
      recentCards()+
      checkedCards()+
      shLine()+
      '<details class="appinfo"><summary>App '+esc(APP_VERSION)+' · data built '+esc(DB.built)+'</summary>'+
        '<div class="counts">'+DB.counts.vehicles+' active vehicles · '+DB.counts.operators+' active operators · '+(DB.counts.owners||0)+' owners</div>'+
        '<div class="stamp">Built '+esc(DB.built)+' from '+esc(DB.sources.vehicles)+' + '+esc(DB.sources.operators)+
          '<br>App version '+APP_VERSION+'</div>'+
        (window.__SHELL__?'<div class="hint">'+(window.__syncLine?window.__syncLine():'')+
          '<span class="link" onclick="window.__checkNow()">Check for new data</span> · '+
          '<span class="link" onclick="window.__datasrc()">Data source…</span> · '+
          '<span class="link" onclick="window.__reimport()">Load a file…</span></div>':'')+
      '</details>';
    wireFilters();
    return;
  }
  /* With more than one group in the result, a row of jump links so Owners or
     Operators are one tap away instead of a scroll past every vehicle. */
  const grps=[["v","Vehicles",r.v.length],["w","Owners",r.w.length],["o","Operators",r.o.length]].filter(g=>g[2]);
  let h=grps.length>1
    ?'<div class="jump">'+grps.map(g=>'<span onclick="jumpTo(\''+g[0]+'\')">'+g[1]+' '+g[2]+' ↓</span>').join("")+'</div>'
    :'<div class="counts">'+r.v.length+' vehicle'+(r.v.length===1?"":"s")+', '+r.o.length+' operator'+(r.o.length===1?"":"s")+', '+r.w.length+' owner'+(r.w.length===1?"":"s")+'</div>';
  const rh=(k,label,n)=>'<div class="reshead" id="rh-'+k+'">'+label+'<span class="n">'+n+'</span></div>';
  /* An owner-operator surfaces once as an owner and once as an operator.
     Flag the pair only when both sides actually made the result set. */
  const oInResult=new Set(r.o.map(o=>o._i));
  const pairedOps=new Set();
  r.w.forEach(w=>{if(w._op!=null&&oInResult.has(w._op))pairedOps.add(w._op);});
  if(r.v.length){h+=rh("v","Vehicles",r.v.length)+pg("sv",r.v,v=>vehCard(v));}
  if(r.w.length){h+=rh("w","Owners",r.w.length)+
    pg("sw",r.w,w=>ownerCard(w,(w._op!=null&&oInResult.has(w._op))?pairNote("\u2193","Operators"):""));}
  if(r.o.length){h+=rh("o","Operators",r.o.length)+
    pg("so",r.o,o=>opCard(o,pairedOps.has(o._i)?pairNote("\u2191","Owners"):""));}
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
    const wa=Math.max(0,...vehFlagList(a).filter(x=>!x.miss).map(x=>x.w));
    const wb=Math.max(0,...vehFlagList(b).filter(x=>!x.miss).map(x=>x.w));
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
  const fl=vehFlagList(v).filter(x=>!x.miss);
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
  /* Read before recording this visit, so the page reports the last time the
     vehicle was looked at rather than the look being taken now. Kept separate
     from the driver's history on purpose -- the same car may be out with a
     different operator, so a vehicle check says nothing about the driver. */
  const priorV=visitOpen(clKeyV(v),vehLabel(v));
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
  h+=clPanel(priorV,"markCheckedV("+i+")");
  h+=ntBar(clKeyV(v));
  h+='<button class="copybtn quiet" onclick="copyRec(\'v\','+i+',this)">Copy record summary</button>';
  h+='<div class="grid">'+
    row("Owner",owner?('<span class="link" onclick="go(\'w/'+owner._i+'\')">'+esc(ownerName(v))+' &#8250;</span>'):dash(ownerName(v)))+
    (owner&&owner._op!=null?row("Owner-Operator",opLink(owner._op)):"")+
    ((owner?owner._op==null&&(owner._opc||[]).length:(v._opc||[]).length)?row("Owner-Operator",opcNote(owner?owner._opc:v._opc)):"")+
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
  h+=ntSection("v",i,clKeyV(v));
  if(v["Notes"]) h+='<div class="seclabel">Notes from the system record</div><div class="notes">'+esc(v["Notes"])+'</div>';
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
  const prior=visitOpen(clKey(o),opName(o));
  const flag=(o["In Active"]===true||o["Cancelled"]===true);
  const ownedW=W.filter(w=>w._op===i);
  let h="";
  if(ownedW.length) h+='<div class="alertbar" style="background:rgba(95,174,255,.14);color:var(--accent);border-color:rgba(95,174,255,.4)">'+
    '&#9733; OWNER-OPERATOR &mdash; also registered as vehicle owner</div>';
  if(flag) h+='<div class="alertbar">&#9888; Licence flagged '+
    (o["Cancelled"]===true?'CANCELLED':'INACTIVE')+' in system</div>';
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
  h+=clPanel(prior,"markChecked("+i+")");
  h+=ntBar(clKey(o));
  h+='<button class="copybtn quiet" onclick="copyRec(\'o\','+i+',this)">Copy record summary</button>';
  h+='<div class="grid">'+
    row("Address",addr||"\u2014")+
    row("Phone",telLink(o["Phone"]))+
    row("Cell",telLink(o["Cell Phone"]))+
    row("Master number",dash(o["Master Number"]),1)+
    row("Operator ID",dash(o["ID"]))+
    row("Licence year",dash(o["v Year"]))+
    row("Approved",dash(o["Approval Date"]))+
    row("District",dash(o["District"]))+
    '</div>';
  h+=ntSection("o",i,clKey(o));
  if(o["Notes"]) h+='<div class="seclabel">Notes from the system record</div><div class="notes">'+esc(o["Notes"])+'</div>';
  const veh=(o._veh||[]).map(ix=>V[ix]).filter(Boolean);
  h+='<div class="seclabel">Vehicles in this name ('+veh.length+')</div>';
  const sn=(o._vsn||[]).map(ix=>V[ix]).filter(Boolean);
  h+=veh.length?veh.map(vehCard).join(""):'<div class="empty">'+(sn.length?
    'None confirmed for this operator — see possible matches below.':
    'No active vehicles registered under this exact name.')+'</div>';
  const fz=(o._vfz||[]).map(ix=>V[ix]).filter(Boolean);
  if(sn.length)h+='<div class="seclabel">Possible matches \u2014 other operators share this name, verify before relying on ('+sn.length+')</div>'+sn.map(vehCard).join("");
  if(fz.length)h+='<div class="seclabel">Possible matches \u2014 similar name, verify before relying on ('+fz.length+')</div>'+fz.map(vehCard).join("");
  main.innerHTML=h;
}

function renderOwner(i){
  const w=W[i]; if(!w){renderHome("");return;}
  noteRecent("w/"+i);
  const priorW=visitOpen(clKeyW(w),ownerName(w));
  const op=(w._op!=null)?O[w._op]:null;
  let h='<div class="dhead"><div class="dtop">'+
    '<div class="opdot" style="flex:0 0 66px;height:66px;font-size:22px">'+
    esc(((w["Owner First Name"]||" ")[0]+(w["Owner Last Name"]||" ")[0]).toUpperCase())+'</div>'+
    '<div><div class="dtitle">'+esc(ownerName(w))+'</div>'+
    '<div class="dsub">'+dash(w["Business Name"])+' &middot; <span class="chip" style="background:var(--panel2);color:var(--dim)">OWNER</span>'+
    (op?' <span class="chip" style="background:rgba(95,174,255,.16);color:var(--accent)">OWNER-OPERATOR</span>':'')+
    '</div></div></div></div>';
  h+=clPanel(priorW,"markCheckedW("+i+")");
  h+=ntBar(nkW(w));
  h+='<button class="copybtn quiet" onclick="copyRec(\'w\','+i+',this)">Copy record summary</button>';
  h+='<div class="grid">'+
    row("Owner ID",dash(w["Owner ID"]))+
    row("Address",dash(w["Owner Address"]))+
    (op?row("Operator link",opLink(w._op)):
       (w._opc||[]).length?row("Operator link",opcNote(w._opc)):
       row("Operator link",'<span class="empty">Not separately licensed as an operator</span>'))+
    '</div>';
  h+=ntSection("w",i,nkW(w));
  const veh=(w._veh||[]).map(ix=>V[ix]).filter(Boolean);
  h+='<div class="seclabel">Vehicles ('+veh.length+')</div>';
  h+=veh.length?veh.map(v=>vehCard(v)).join(""):'<div class="empty">No active vehicles on file for this owner.</div>';
  main.innerHTML=h;
}

/* ---------- routing ---------- */
const q=document.getElementById("q"), clr=document.getElementById("clr");
function go(route){location.hash="#/"+route;}
var ROUTE_LAST=null;
function route(){
  const h=location.hash.replace(/^#\/?/,"");
  /* A new page is a visit (and logs a query); the same page drawn again by a
     sync or keepScroll is not. */
  FRESH=(h!==ROUTE_LAST); ROUTE_LAST=h;
  const m=/^(v|o|w)\/(\d+)$/.exec(h);
  const sp=(h==="shared"||h==="actions");
  window.scrollTo(0,0);
  document.getElementById("back").style.display=(m||sp)?"flex":"none";
  if(h==="shared"){ renderShared(); }
  else if(h==="actions"){ renderActions(); }
  else if(m){ if(m[1]==="v") renderVehicle(+m[2]); else if(m[1]==="o") renderOperator(+m[2]); else renderOwner(+m[2]); }
  else { renderHome(q.value); }
  FRESH=false;
}
document.body.classList.add("light");
document.getElementById("theme").addEventListener("click",()=>{
  const light=document.body.classList.toggle("light");
  document.getElementById("theme").innerHTML=light?"&#9789;":"&#9788;";
});
/* ---------- due date ---------- */
/* Day 30 from the issue date, moved forward to the next Friday the courts are
   open. Only a closure that lands on a Friday can change the result, so the
   Monday holidays are listed for completeness. Review against the Courts of
   Nova Scotia closure list each year. Provincial closure dates that are not
   fixed holidays go in EXTRA_CLOSURES as "YYYY-MM-DD". */
const EXTRA_CLOSURES=[];
function localDay(y,m,d){return new Date(y,m,d,12,0,0,0);}
function ymd(d){return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");}
function nthWeekday(y,m,wd,n){
  const first=localDay(y,m,1);
  return localDay(y,m,1+(wd-first.getDay()+7)%7+7*(n-1));
}
function easterSunday(y){
  const a=y%19,b=Math.floor(y/100),c=y%100,d=Math.floor(b/4),e=b%4,
    f=Math.floor((b+8)/25),g=Math.floor((b-f+1)/3),h=(19*a+b-d-g+15)%30,
    i=Math.floor(c/4),k=c%4,l=(32+2*e+2*i-h-k)%7,m=Math.floor((a+11*h+22*l)/451),
    t=h+l-7*m+114;
  return localDay(y,Math.floor(t/31)-1,t%31+1);
}
function nsHolidays(y){
  const easter=easterSunday(y), vic=localDay(y,4,24);
  vic.setDate(24-(vic.getDay()+6)%7);
  return [
    localDay(y,0,1),                                   // New Year's Day
    nthWeekday(y,1,1,3),                               // Heritage Day
    localDay(y,easter.getMonth(),easter.getDate()-2),  // Good Friday
    vic,                                               // Victoria Day
    localDay(y,6,1),                                   // Canada Day
    nthWeekday(y,7,1,1),                               // Natal Day
    nthWeekday(y,8,1,1),                               // Labour Day
    localDay(y,8,30),                                  // Truth and Reconciliation
    nthWeekday(y,9,1,2),                               // Thanksgiving
    localDay(y,10,11),                                 // Remembrance Day
    localDay(y,11,25),                                 // Christmas Day
    localDay(y,11,26)                                  // Boxing Day
  ];
}
function courtClosed(d){
  const k=ymd(d);
  return EXTRA_CLOSURES.indexOf(k)>-1||nsHolidays(d.getFullYear()).some(h=>ymd(h)===k);
}
function dueDate(issued){
  const day30=localDay(issued.getFullYear(),issued.getMonth(),issued.getDate()+30);
  const d=new Date(day30), skipped=[];
  while(d.getDay()!==5)d.setDate(d.getDate()+1);
  while(courtClosed(d)&&skipped.length<60){skipped.push(new Date(d));d.setDate(d.getDate()+7);}
  return {day30:day30,due:d,skipped:skipped};
}
function fmtDay(d){
  return ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"][d.getDay()]+" "+
    ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][d.getMonth()]+" "+
    String(d.getDate()).padStart(2,"0")+", "+d.getFullYear();
}
function showDue(){
  const v=document.getElementById("dueissued").value, el=document.getElementById("dueval"),
    sub=document.getElementById("duesub");
  const m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(v);
  if(!m){el.textContent="—";sub.textContent="Enter the date the ticket was issued.";return;}
  const r=dueDate(localDay(+m[1],+m[2]-1,+m[3]));
  el.textContent=fmtDay(r.due);
  sub.textContent="Day 30 is "+fmtDay(r.day30)+(r.skipped.length?". Court closed "+
    r.skipped.map(fmtDay).join(", ")+".":".");
}
function openDue(){
  const t=new Date();
  document.getElementById("dueissued").value=ymd(t);
  showDue();
  document.getElementById("duesheet").style.display="flex";
}
function closeDue(){document.getElementById("duesheet").style.display="none";}
document.getElementById("duebtn").addEventListener("click",openDue);
document.getElementById("dueclose").addEventListener("click",closeDue);
document.getElementById("duesheet").addEventListener("click",e=>{if(e.target.id==="duesheet")closeDue();});
document.addEventListener("keydown",e=>{if(e.key==="Escape"){closeDue();closeGeo();}});
document.getElementById("dueissued").addEventListener("input",showDue);
/* ---------- location ---------- */
/* GPS starts when the location button is tapped and stops when its sheet
   closes. Map data is the public NS Civic Address File for CBRM (docs/geo.bin,
   gzipped JSON), fetched on first use and then held by the service worker. */
let GEO=null, GCELL=0, geoWatch=null, geoOn=false, geoHeading=null, geoCur=null, geoFilter=null;
const R_EARTH=6371000, D2R=Math.PI/180;
function hav(lo1,la1,lo2,la2){
  const dla=(la2-la1)*D2R, dlo=(lo2-lo1)*D2R;
  const a=Math.sin(dla/2)**2+Math.cos(la1*D2R)*Math.cos(la2*D2R)*Math.sin(dlo/2)**2;
  return 2*R_EARTH*Math.asin(Math.sqrt(a));
}
function geoCells(lon,lat){
  const cx=Math.floor(lon/GCELL), cy=Math.floor(lat/GCELL), out=[];
  for(let dx=-1;dx<=1;dx++)for(let dy=-1;dy<=1;dy++)out.push((cx+dx)+":"+(cy+dy));
  return out;
}
function ptSeg(plon,plat,alon,alat,blon,blat){
  const kx=Math.cos(plat*D2R)*111320, ky=111320;
  const px=plon*kx, py=plat*ky, ax=alon*kx, ay=alat*ky, bx=blon*kx, by=blat*ky;
  const dx=bx-ax, dy=by-ay, L2=dx*dx+dy*dy;
  let t=L2?((px-ax)*dx+(py-ay)*dy)/L2:0;
  t=Math.max(0,Math.min(1,t));
  return Math.hypot(px-(ax+t*dx),py-(ay+t*dy));
}
function nearestCivic(lon,lat){
  let best=null, bd=Infinity;
  for(const c of geoCells(lon,lat)){
    const ids=GEO.pt_grid[c]; if(!ids)continue;
    for(const i of ids){const p=GEO.points[i], d=hav(lon,lat,p.lon,p.lat); if(d<bd){bd=d;best=p;}}
  }
  return best?{p:best,d:bd}:null;
}
/* Distance is to the line between road vertices, not to the vertices
   themselves; vertex distance puts you on the wrong street. */
function nearestRoad(lon,lat){
  let best=null, bd=Infinity;
  const seen=new Set();
  for(const c of geoCells(lon,lat)){
    const ids=GEO.rd_grid[c]; if(!ids)continue;
    for(const i of ids){
      if(seen.has(i))continue; seen.add(i);
      const r=GEO.roads[i], g=r.g;
      for(let k=0;k<g.length-1;k++){
        const d=ptSeg(lon,lat,g[k][0],g[k][1],g[k+1][0],g[k+1][1]);
        if(d<bd){bd=d;best=r;}
      }
    }
  }
  return best?{r:best,d:bd}:null;
}
function bearing(la1,lo1,la2,lo2){
  const y=Math.sin((lo2-lo1)*D2R)*Math.cos(la2*D2R);
  const x=Math.cos(la1*D2R)*Math.sin(la2*D2R)-Math.sin(la1*D2R)*Math.cos(la2*D2R)*Math.cos((lo2-lo1)*D2R);
  return (Math.atan2(y,x)/D2R+360)%360;
}
/* Intersections on the current street within 800 m. With a travel heading they
   split into ahead and behind; without one, the two nearest are returned. */
function crossStreets(lon,lat,street,heading){
  if(!street)return {ahead:null,behind:null};
  const cand=[];
  for(const n of GEO.nodes){
    if(n.st.indexOf(street)<0)continue;
    const d=hav(lon,lat,n.lon,n.lat); if(d>800)continue;
    const others=n.st.filter(x=>x!==street); if(!others.length)continue;
    cand.push({d:d,br:bearing(lat,lon,n.lat,n.lon),name:others[0]});
  }
  if(heading==null){
    cand.sort((a,b)=>a.d-b.d);
    return {ahead:cand[0]||null,behind:cand[1]||null};
  }
  let ahead=null, behind=null;
  for(const c of cand){
    const diff=Math.abs(((c.br-heading+540)%360)-180);
    if(diff<90){if(!ahead||c.d<ahead.d)ahead=c;}
    else if(!behind||c.d<behind.d)behind=c;
  }
  return {ahead:ahead,behind:behind};
}
/* Blends each fix by its reported accuracy against the filter's own confidence,
   so one noisy reading does not flip the street. Process noise follows GPS
   speed: quick while driving, smooth while walking or standing. */
function makeGeoFilter(){
  let lat=null, lon=null, variance=-1, ts=null;
  return function(nLat,nLon,nAcc,nTs,speed){
    if(nAcc==null||isNaN(nAcc))nAcc=50;
    nAcc=Math.max(1,nAcc);
    if(variance<0||ts==null||(nTs-ts)>30000){lat=nLat;lon=nLon;variance=nAcc*nAcc;ts=nTs;}
    else{
      const q=Math.max(speed||0,1.5), dt=Math.max(0,nTs-ts)/1000;
      variance+=dt*q*q; ts=nTs;
      const k=variance/(variance+nAcc*nAcc);
      lat+=k*(nLat-lat); lon+=k*(nLon-lon); variance=(1-k)*variance;
    }
    return {lat:lat,lon:lon,acc:Math.sqrt(variance)};
  };
}
async function loadGeo(){
  if(GEO)return;
  const res=await fetch("geo.bin");
  if(!res.ok)throw new Error("HTTP "+res.status);
  const text=await new Response(res.body.pipeThrough(new DecompressionStream("gzip"))).text();
  GEO=JSON.parse(text);
  GCELL=GEO.meta.grid_cell_deg;
}
function geoSay(t){document.getElementById("geostatus").textContent=t||"";}
function geoReset(){
  geoCur=null; geoHeading=null; geoFilter=null;
  document.getElementById("geobox").style.display="none";
  document.getElementById("geonote").style.display="none";
  document.getElementById("geobtn").style.display="none";
  geoSay("");
}
function stopGeo(){
  geoOn=false;
  if(geoWatch!=null){navigator.geolocation.clearWatch(geoWatch);geoWatch=null;}
  geoReset();
}
function csLine(c,prefix){return (prefix||"")+c.name+" ("+Math.round(c.d)+" m)";}
function renderGeo(){
  const g=geoCur, $=id=>document.getElementById(id);
  $("geobox").style.display="block"; $("geonote").style.display="block"; geoSay("");
  const a=Math.round(g.acc);
  $("geodot").className="geodot "+(a<=15?"good":a<=40?"fair":"weak");
  $("geoacc").textContent="GPS \u00B1"+a+" m \u00B7 "+(a<=15?"Good":a<=40?"Fair":"Weak, move to open sky");
  if(g.road){
    $("geostreet").textContent=g.road.r.s||"(unnamed road)";
    $("geostreetsub").textContent=Math.round(g.road.d)+" m from the road centreline";
  }else{
    $("geostreet").textContent="No CBRM road data here";
    $("geostreetsub").textContent="";
  }
  if(g.civic){
    $("geocivic").textContent=g.civic.p.n+" "+g.civic.p.s;
    $("geocivicsub").textContent=(g.civic.p.c?g.civic.p.c+" \u00B7 ":"")+Math.round(g.civic.d)+" m away";
  }else{
    $("geocivic").textContent="None in range"; $("geocivicsub").textContent="";
  }
  const cs=g.cs, known=geoHeading!=null, lines=[];
  if(known){
    if(cs.ahead)lines.push(csLine(cs.ahead,"Ahead: "));
    if(cs.behind)lines.push(csLine(cs.behind,"Behind: "));
  }else{
    if(cs.ahead)lines.push(csLine(cs.ahead));
    if(cs.behind)lines.push(csLine(cs.behind));
  }
  $("geocs").textContent=lines.length?lines.join("\n"):"\u2014";
}
function onGeoFix(pos){
  const c=pos.coords;
  if(c.heading!=null&&!isNaN(c.heading)&&c.speed!=null&&c.speed>1.5)geoHeading=c.heading;
  const f=geoFilter(c.latitude,c.longitude,c.accuracy,pos.timestamp,c.speed);
  const road=nearestRoad(f.lon,f.lat);
  geoCur={acc:f.acc,road:road,civic:nearestCivic(f.lon,f.lat),
    cs:crossStreets(f.lon,f.lat,road&&road.r.s?road.r.s:null,geoHeading)};
  renderGeo();
}
function onGeoErr(err){
  geoSay(err&&err.code===1
    ?"Location is blocked. Allow location for this site in your browser settings, then try again."
    :"Can't get a GPS fix. Try again with a clear view of the sky.");
  if(geoWatch!=null){navigator.geolocation.clearWatch(geoWatch);geoWatch=null;}
  geoOn=false;
  document.getElementById("geobtn").style.display="block";
}
async function startGeo(){
  if(geoOn)return;
  geoOn=true;
  document.getElementById("geobtn").style.display="none";
  if(!navigator.geolocation){geoSay("This device has no location support.");geoOn=false;document.getElementById("geobtn").style.display="block";return;}
  geoSay("Loading map data\u2026");
  try{await loadGeo();}
  catch(e){
    if(!geoOn)return;
    geoSay("Map data is not available here. Open the installed app once with a signal to download it.");
    geoOn=false; document.getElementById("geobtn").style.display="block"; return;
  }
  if(!geoOn)return;
  geoSay("Finding your position\u2026");
  geoFilter=makeGeoFilter();
  geoWatch=navigator.geolocation.watchPosition(onGeoFix,onGeoErr,{enableHighAccuracy:true,maximumAge:1000,timeout:15000});
}
function openGeo(){
  document.getElementById("geosheet").style.display="flex";
  startGeo();
}
function closeGeo(){
  document.getElementById("geosheet").style.display="none";
  stopGeo();
}
document.getElementById("locbtn").addEventListener("click",openGeo);
document.getElementById("geoclose").addEventListener("click",closeGeo);
document.getElementById("geosheet").addEventListener("click",e=>{if(e.target.id==="geosheet")closeGeo();});
document.getElementById("geobtn").addEventListener("click",startGeo);
q.addEventListener("input",()=>{
  LIM={};
  clr.style.display=q.value?"block":"none";
  if(location.hash && location.hash!=="#/") history.replaceState(null,"","#/");
  renderHome(q.value);
});
clr.addEventListener("click",()=>{q.value="";clr.style.display="none";q.focus();renderHome("");});
window.go=go;window.copyRec=copyRec;window.copyField=copyField;window.setSort=setSort;
window.markChecked=markChecked;window.markCheckedV=markCheckedV;
window.markCheckedW=markCheckedW;window.clearCheckLog=clearCheckLog;
window.pgMore=pgMore;window.jumpTo=jumpTo;window.ntShow=ntShow;window.ntDismiss=ntDismiss;window.ntHist=ntHist;window.ntEdit=ntEdit;window.ntEditCancel=ntEditCancel;
window.ntEditSave=ntEditSave;window.ntRemove=ntRemove;window.chToggle=chToggle;window.chMore=chMore;window.ntAdd=ntAdd;window.ntDone=ntDone;window.ntJump=ntJump;
window.shNow=shNow;window.shOut=shOut;window.shDiscard=shDiscard;
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
