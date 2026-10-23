#!/usr/bin/env python3
r"""
test_server.py -- try the app with made-up records and a pretend shared-notes
server, so testing never touches the real data or the real database.

    py test_server.py [port]          (default 8795; Start_Test_Server.bat)

Open http://127.0.0.1:8795/ and sign in under Shared notes as
officer1@test or officer2@test, password "test" (two accounts, so editing
someone else's note or unlinking their link can be tried).

Operator photos: the first three made-up operators have drawn cartoon faces, served
from a pretend photo bucket to a signed-in officer only; the fourth has none (initials
stay). Needs Pillow (py -m pip install pillow). /__test/photo_hits counts requests.

What it serves is the single-file app built from build_pvh_field_app.py as it
is in the folder right now -- saved but uncommitted edits included -- rebuilt
on every reload. The records are invented and the shared-notes server is this
script: notes, stops, queries and links live in memory and are gone when the
window is closed. Nothing is read from PVH_data.json or pvh_local_config.json,
and nothing is sent anywhere else.

Why not docs/ (the installed app)? docs/ is built to talk only to the real
shared-notes server -- every record opened there logs a real "Queried" mark
under your name, and every link or note is real.

Phones: only this PC can reach it (127.0.0.1). To try it on a phone in your
tailnet, run once in an ADMINISTRATOR terminal:
    tailscale serve --bg --https=8447 http://127.0.0.1:8795
and open https://<this PC's tailnet name>:8447/ . Remove it afterwards with
    tailscale serve --https=8447 off
"""
import io
import os
import sys
import json
import uuid
import hashlib
import datetime
import threading
import importlib.util
import urllib.parse
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

HERE = os.path.dirname(os.path.abspath(__file__))
PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8795
PASSWORD = "test"
USERS = {"officer1@test": ("Officer One", "00000000-0000-4000-8000-000000000001"),
         "officer2@test": ("Officer Two", "00000000-0000-4000-8000-000000000002")}
TOKENS = {uid: name for name, uid in USERS.values()}


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def ago(days):
    return (datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=days)).isoformat()


def load_builder():
    """The build script as saved right now, without running a build."""
    argv = sys.argv
    sys.argv = [argv[0], "--data-only"]   # never touches app_version.json
    try:
        spec = importlib.util.spec_from_file_location("pvh_build_for_test", os.path.join(HERE, "build_pvh_field_app.py"))
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod
    finally:
        sys.argv = argv


# ------------------------------------------------------------- made-up records
def made_up_records(B):
    far, soon, gone = "12/31/2030", (datetime.date.today() + datetime.timedelta(days=12)).strftime("%m/%d/%Y"), "01/15/2025"
    veh = []
    for deck, mm, plate, lic, first, last, biz, permit in [
            (901, "Ford/Escape", "TEST-901", "ZZT901", "Avery", "Testowner", "", far),
            (902, "Toyota/Sienna", "TEST-902", "ZZT902", "Blake", "Sampleton", "Example Cabs", soon),
            (903, "Honda/Civic", "TEST-903", "ZZT903", "Blake", "Sampleton", "Example Cabs", gone),
            ("", "Kia/Carnival", "TEST-904", "ZZT904", "Morgan", "Placeholder", "", far),
            # Quotes, an apostrophe and a backslash, to show text is escaped everywhere.
            (905, "Dodge/Grand \"Caravan\"", "TEST'905", "ZZT905", "Riley", "O'Testly", "<b>Quote's</b> \"Cabs\"", far)]:
        veh.append({"Vehicle Type": "Taxi", "Business Name": biz, "Owner Last Name": last, "Owner First Name": first,
                    "Owner ID": "T-" + last, "VIN": ("ZZ\\TEST'VIN" if "'" in plate else "ZZTESTVIN0000") + lic[-3:], "Make Model": mm, "Vehicle Color": "White",
                    "v Year": 2021, "Deck No": deck, "Plate No": plate, "NSVehicle Permit Expiry": permit,
                    "District": "Test", "First MVIDate": "06/01/2026", "Insurance Expiry": far,
                    "Notes": "Made-up record for testing.", "Vehicle ID": lic, "Inspection Date": "06/01/2026",
                    "Licence No": lic, "Expiry Date": far, "Owner Address": "1 Example Street"})
    ops = []
    for i, (first, last, lic, renew) in enumerate([
            ("Avery", "Testowner", "ZZD001", far), ("Casey", "Testdriver", "ZZD002", soon),
            ("Devon", "Mockley", "ZZD003", gone), ("Blake", "Sampleton", "ZZD004", far)]):
        ops.append({"Operator Type": "Taxi", "ID": 9000 + i, "Business Name": "", "Last Name": last, "First Name": first,
                    "Middle": "", "Address1": "1 Example Street", "Address2": "", "City": "Testville", "Province": "NS",
                    "Postal Code": "A0A 0A0", "Phone": "555-0100", "Cell Phone": "555-0101",
                    "Master Number": "ZZTEST00000%d" % i, "Licence Number": lic, "v Year": 2026,
                    "NSDLExpired Date": far, "Approval Date": "01/01/2020", "Renewal Date": renew, "District": "Test",
                    "In Active": None, "Cancelled": None, "Notes": None})
    # The same name linking the real build does (main() in the builder).
    by_name = {}
    for j, o in enumerate(ops):
        by_name.setdefault(B.norm_name(o["Last Name"], o["First Name"]), []).append(j)
    for o in ops:
        o["_veh"], o["_vsn"], o["_vfz"] = [], [], []
    for i, v in enumerate(veh):
        v["_op"], v["_opc"] = B.resolve_operator(B.norm_name(v["Owner Last Name"], v["Owner First Name"]),
                                                 v["Owner Address"], by_name, ops)
        if v["_op"] is not None:
            ops[v["_op"]]["_veh"].append(i)
    owners = B.build_owners(veh, by_name, ops)
    return {"built": "TEST DATA -- made up", "sources": {"vehicles": "made up", "operators": "made up", "owners": "made up"},
            "counts": {"vehicles": len(veh), "operators": len(ops), "owners": len(owners)},
            "vehicles": veh, "operators": ops, "owners": owners,
            "shared": {"url": "set by the page", "key": "test-key"}}


def app_html():
    B = load_builder()
    payload = json.dumps(made_up_records(B), separators=(",", ":")).replace("</", "<\\/")
    # The pretend server is wherever the page came from, so this also works
    # through a tailnet address.
    boot = ('var d=JSON.parse(document.getElementById("data").textContent);'
            'd.shared.url=location.origin;boot(d);')
    return (B.TEMPLATE.replace("__HEAD_EXTRA__", "")
            .replace("__DATA_SCRIPT__", '<script id="data" type="application/json">' + payload + "</script>")
            .replace("__BOOT__", boot).replace("__SHPIN__", "").replace("__APPVER__", "TEST BUILD"))


# ----------------------------------------------------------- pretend photo bucket
# Drawn cartoon faces for the first three made-up operators (the fourth has none, to
# show the initials fallback). Files are named as the real bucket's are: the first 24
# hex characters of SHA-256("pvh-photo:" + the operator's record key).
PHOTO_HITS = {}   # file name -> times the "bucket" was asked for it (shown at /__test/photo_hits)


def photo_id(key):
    return hashlib.sha256(("pvh-photo:" + key).encode("utf-8")).hexdigest()[:24]


def cartoon_face(i):
    from PIL import Image, ImageDraw
    bg = [(150, 175, 205), (190, 170, 150), (160, 190, 170)][i % 3]
    skin = [(224, 172, 130), (141, 85, 54), (245, 205, 170)][i % 3]
    hair = [(40, 30, 25), (20, 20, 20), (150, 90, 40)][i % 3]
    shirt = [(60, 70, 90), (120, 40, 40), (50, 90, 70)][i % 3]
    w, h = 120, 158
    im = Image.new("RGB", (w, h), bg)
    d = ImageDraw.Draw(im)
    d.ellipse((5, 118, 115, 220), fill=shirt)                  # shoulders
    d.rectangle((48, 98, 72, 128), fill=skin)                  # neck
    d.ellipse((24, 22, 96, 112), fill=hair)                    # hair behind
    d.ellipse((28, 32, 92, 108), fill=skin)                    # face
    d.pieslice((26, 26, 94, 82), 180, 360, fill=hair)          # fringe
    for x in (46, 74):
        d.ellipse((x - 5, 62, x + 5, 70), fill="white")
        d.ellipse((x - 2, 63, x + 2, 69), fill=(30, 30, 30))   # eyes
    d.arc((46, 76, 74, 96), 20, 160, fill=(120, 40, 40), width=2)   # smile
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=85)
    return buf.getvalue()


PHOTOS = {}


def make_photos():
    for i in range(3):
        PHOTOS[photo_id("ZZTEST00000%d" % i) + ".jpg"] = cartoon_face(i)


# ------------------------------------------------------- pretend shared notes
DB = {"notes": [], "checks": [], "queries": [], "note_events": [], "links": [], "link_events": []}
LOCK = threading.Lock()   # the browser sends requests side by side


def seed():
    two = USERS["officer2@test"]
    DB["notes"].append({"id": str(uuid.uuid4()), "rec_key": "v:ZZT902", "rec_type": "vehicle",
                        "rec_label": "Toyota/Sienna TEST-902", "body": "Made-up note: check the roof light.",
                        "action_needed": True, "written_at": ago(1), "created_at": ago(1), "updated_at": ago(1),
                        "author_id": two[1], "author_name": two[0], "done_at": None, "done_by": None,
                        "edited_at": None, "edited_by": None, "removed_at": None, "removed_by": None})
    DB["checks"].append({"id": str(uuid.uuid4()), "rec_key": "o:ZZTEST000001", "rec_label": "Casey Testdriver ZZD002",
                         "written_at": ago(2), "created_at": ago(2), "author_id": two[1], "author_name": two[0]})
    DB["links"].append({"id": "l|o:ZZTEST000001|v:ZZT901", "op_key": "o:ZZTEST000001", "veh_key": "v:ZZT901",
                        "op_label": "Casey Testdriver ZZD002", "veh_label": "Deck 901 Ford/Escape TEST-901",
                        "active": True, "created_at": ago(3), "created_by": two[1], "created_name": two[0],
                        "updated_at": ago(3), "updated_by": two[1], "updated_name": two[0]})


class Refused(Exception):
    def __init__(self, code, message):
        super().__init__(message)
        self.code, self.message = code, message


def eq(query, field):
    v = query.get(field, [""])[0]
    return v[3:] if v.startswith("eq.") else None


def write(table, uid, row, query):
    """Inserts, with the database's own rules from supabase/schema.sql."""
    name = TOKENS[uid]
    rows = DB[table]
    if table == "links":
        if not (row.get("op_key", "").startswith("o:") and row.get("veh_key", "").startswith("v:")
                and row.get("id") == "l|%s|%s" % (row.get("op_key"), row.get("veh_key"))):
            raise Refused(400, "new row violates check constraint \"links_id_shape\"")
        old = next((r for r in rows if r["id"] == row["id"]), None)
        if old is None:
            rows.append(dict(row, active=True, created_at=now(), created_by=uid, created_name=name,
                             updated_at=now(), updated_by=uid, updated_name=name))
            DB["link_events"].append({"link_id": row["id"], "kind": "add", "by_name": name, "created_at": now()})
        elif not old["active"]:
            old.update(active=True, op_label=row.get("op_label"), veh_label=row.get("veh_label"),
                       updated_at=now(), updated_by=uid, updated_name=name)
            DB["link_events"].append({"link_id": row["id"], "kind": "add", "by_name": name, "created_at": now()})
        return
    if any(r["id"] == row.get("id") for r in rows):
        raise Refused(409, "duplicate key value violates unique constraint")
    new = dict(row, id=row.get("id") or str(uuid.uuid4()), author_id=uid, author_name=name, created_at=now())
    new["written_at"] = min(new.get("written_at") or now(), now())
    if table == "notes":
        new.update(updated_at=now(), done_at=None, done_by=None, edited_at=None, edited_by=None,
                   removed_at=None, removed_by=None)
    rows.append(new)


def change(table, uid, patch, query):
    """Updates, with the same rules the database triggers apply."""
    name = TOKENS[uid]
    rid = eq(query, "id")
    row = next((r for r in DB[table] if r["id"] == rid), None)
    if row is None:
        return
    if table == "links":
        if "active" in patch and patch["active"] != row["active"]:
            row.update(active=bool(patch["active"]), updated_at=now(), updated_by=uid, updated_name=name)
            DB["link_events"].append({"link_id": rid, "kind": "add" if patch["active"] else "remove",
                                      "by_name": name, "created_at": now()})
        return
    if table != "notes":
        raise Refused(403, "permission denied")
    if "body" in patch and patch["body"] != row["body"]:
        if row["removed_at"]:
            raise Refused(400, "A removed note cannot be edited")
        if row["author_id"] != uid:
            raise Refused(400, "Only the author can edit a note")
        DB["note_events"].append({"id": str(uuid.uuid4()), "note_id": rid, "kind": "edit", "old_body": row["body"],
                                  "new_body": patch["body"], "by_name": name, "created_at": now()})
        row.update(body=patch["body"], edited_at=now(), edited_by=name)
    if patch.get("removed_at"):
        if row["removed_at"]:
            raise Refused(400, "A removal cannot be undone from the app")
        if row["author_id"] != uid:
            raise Refused(400, "Only the author can remove a note")
        DB["note_events"].append({"id": str(uuid.uuid4()), "note_id": rid, "kind": "remove", "old_body": row["body"],
                                  "by_name": name, "created_at": now()})
        row.update(removed_at=now(), removed_by=name)
    if "done_at" in patch:
        row.update(done_at=now() if patch["done_at"] else None, done_by=name if patch["done_at"] else None)
    row["updated_at"] = now()


def read(table, query):
    rows = list(DB.get(table, []))
    for field, vals in query.items():
        v = vals[0]
        if v.startswith("gte."):
            rows = [r for r in rows if str(r.get(field) or "") >= v[4:]]
        elif v.startswith("eq."):
            rows = [r for r in rows if str(r.get(field)) == v[3:]]
    order = query.get("order", [""])[0].split(".")[0]
    if order:
        rows.sort(key=lambda r: str(r.get(order) or ""))
    return rows[:int(query.get("limit", ["1000"])[0])]


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass

    def send(self, code, obj=None, ctype="application/json"):
        body = b"" if obj is None else (obj if isinstance(obj, bytes) else json.dumps(obj).encode("utf-8"))
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def parts(self):
        u = urllib.parse.urlparse(self.path)
        return u.path, urllib.parse.parse_qs(u.query)

    def body(self):
        n = int(self.headers.get("Content-Length") or 0)
        try:
            return json.loads(self.rfile.read(n) or b"null")
        except ValueError:
            return None

    def signed_in(self):
        a = self.headers.get("Authorization", "")
        uid = a[7:] if a.startswith("Bearer ") else ""
        return uid if uid in TOKENS else None

    def table(self, path):
        t = path[len("/rest/v1/"):] if path.startswith("/rest/v1/") else ""
        return t if t in DB or t == "profiles" else None

    def do_GET(self):
        path, query = self.parts()
        if path in ("/", "/index.html"):
            return self.send(200, app_html().encode("utf-8"), "text/html; charset=utf-8")
        if path == "/__test/photo_hits":
            return self.send(200, PHOTO_HITS)
        if path.startswith("/storage/v1/object/authenticated/operator-photos/"):
            if not self.signed_in():
                return self.send(400, {"statusCode": "400", "error": "Unauthorized", "message": "Invalid JWT"})
            name = path.rsplit("/", 1)[1]
            PHOTO_HITS[name] = PHOTO_HITS.get(name, 0) + 1
            if name in PHOTOS:
                return self.send(200, PHOTOS[name], "image/jpeg")
            # The real storage API answers a missing object with HTTP 400 and this body.
            return self.send(400, {"statusCode": "404", "error": "not_found", "message": "Object not found"})
        t = self.table(path)
        if not t:
            return self.send(404, {"message": "not found"})
        uid = self.signed_in()
        if not uid:
            return self.send(401, {"message": "JWT expired"})
        if t == "profiles":
            return self.send(200, [{"display_name": TOKENS[uid]}] if eq(query, "id") == uid else [])
        with LOCK:
            rows = read(t, query)
        self.send(200, rows)

    def do_POST(self):
        path, query = self.parts()
        b = self.body()
        if path == "/auth/v1/token":
            grant = query.get("grant_type", [""])[0]
            if grant == "password" and isinstance(b, dict) and b.get("email") in USERS and b.get("password") == PASSWORD:
                uid = USERS[b["email"]][1]
                return self.send(200, {"access_token": uid, "refresh_token": uid, "expires_in": 3600, "user": {"id": uid}})
            if grant == "refresh_token" and isinstance(b, dict) and b.get("refresh_token") in TOKENS:
                return self.send(200, {"access_token": b["refresh_token"], "refresh_token": b["refresh_token"],
                                       "expires_in": 3600})
            return self.send(400, {"error_description": "Invalid login credentials"})
        if path == "/auth/v1/logout":
            return self.send(204)
        t = self.table(path)
        uid = self.signed_in()
        if not t or t in ("profiles", "note_events", "link_events"):
            return self.send(403, {"message": "permission denied"})
        if not uid:
            return self.send(401, {"message": "JWT expired"})
        try:
            with LOCK:
                write(t, uid, b if isinstance(b, dict) else {}, query)
        except Refused as r:
            return self.send(r.code, {"message": r.message})
        self.send(201)

    def do_PATCH(self):
        path, query = self.parts()
        t = self.table(path)
        uid = self.signed_in()
        if not t:
            return self.send(404, {"message": "not found"})
        if not uid:
            return self.send(401, {"message": "JWT expired"})
        try:
            b = self.body()
            with LOCK:
                change(t, uid, b if isinstance(b, dict) else {}, query)
        except Refused as r:
            return self.send(r.code, {"message": r.message})
        self.send(204)


def main():
    seed()
    make_photos()
    app_html()   # fail now, not in the browser, if the build script has an error
    httpd = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    print("PVH TEST server -- made-up records, pretend shared notes. Nothing real is used or sent.")
    print("Open  http://127.0.0.1:%d/   Sign in: officer1@test or officer2@test, password: %s" % (PORT, PASSWORD))
    print("Reload the page to pick up edits to build_pvh_field_app.py. Close this window to stop.", flush=True)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
