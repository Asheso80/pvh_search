"""Upload operator face thumbnails to the private Supabase bucket `operator-photos`.

PC only. Run it yourself; it is never part of the daily build.

    python pvh_photos_upload.py "D:\\path\\to\\Operator Thumbs"            # upload what is missing
    python pvh_photos_upload.py "D:\\path\\to\\Operator Thumbs" --changed  # missing, plus any whose picture has changed
    python pvh_photos_upload.py "D:\\path\\to\\Operator Thumbs" --check    # look at the folder, no sign-in, no upload
    python pvh_photos_upload.py "D:\\path\\to\\Operator Thumbs" --replace  # overwrite every photo, changed or not

--changed is the one to use after a renewal or a retake: the file name of a person's thumbnail never
changes, so the plain run skips it. This compares each local file with what the bucket holds (its
checksum where the bucket reports one, otherwise its size) and sends only the new and altered ones.
Phones notice the change by themselves the next time they are online and signed in.

Signs in as you (the account listed in `photo_admins`, see supabase/007_photos.sql) with the
project URL and anon key from pvh_local_config.json, exactly as pvh_intel.py does. The email is
remembered as "intel_email"; the password is asked every run and never stored. No service_role
key is used or needed: the storage rules let only a photo administrator write.

Files are `<24 hex characters>.jpg`, the one-way hash of each record key made by the thumbnail
step. Anything else in the folder (contact sheets, manifests) is ignored. The photos are personal
information: this script refuses a folder inside the repo, so they cannot be committed by accident.
Prints counts only, never names.
"""
import getpass
import hashlib
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
CFG_PATH = os.path.join(HERE, "pvh_local_config.json")
BUCKET = "operator-photos"
NAME_RE = re.compile(r"^[0-9a-f]{24}\.jpg$")
MAX_BYTES = 200 * 1024            # the bucket's own limit


def s_(v):
    return "" if v is None else str(v)


def http(method, url, key, token=None, body=None, raw=None, ctype="application/json", extra=None, timeout=60):
    h = {"apikey": key, "Content-Type": ctype}
    if token:
        h["Authorization"] = "Bearer " + token
    h.update(extra or {})
    data = raw if raw is not None else (json.dumps(body).encode() if body is not None else None)
    req = urllib.request.Request(url, method=method, headers=h, data=data)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, json.loads(r.read() or b"null")
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.loads(e.read() or b"null")
        except ValueError:
            return e.code, None
    except (urllib.error.URLError, TimeoutError) as e:
        return 0, {"message": str(e)}


def inside_repo(path):
    repo = os.path.normcase(os.path.realpath(HERE))
    p = os.path.normcase(os.path.realpath(path))
    return p == repo or p.startswith(repo + os.sep)


def local_files(folder):
    good, skipped = [], 0
    for n in sorted(os.listdir(folder)):
        full = os.path.join(folder, n)
        if not os.path.isfile(full) or n.startswith("_"):
            continue
        if NAME_RE.match(n) and 0 < os.path.getsize(full) <= MAX_BYTES:
            good.append(n)
        else:
            skipped += 1
    return good, skipped


def sign_in(cfg):
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
    return url, key, j["access_token"]


def remote_names(url, key, token):
    """Every object in the bucket: {name: the listing entry, with its checksum and size under 'metadata'}."""
    names, off = {}, 0
    while True:
        st, page = http("POST", url + "/storage/v1/object/list/" + BUCKET, key, token,
                        body={"prefix": "", "limit": 1000, "offset": off,
                              "sortBy": {"column": "name", "order": "asc"}})
        if st != 200 or not isinstance(page, list):
            sys.exit("Could not list the bucket: HTTP " + str(st) + " " + s_(page)
                     + "\n(Has supabase/007_photos.sql been run, and is this account in photo_admins?)")
        names.update({p["name"]: p for p in page if p.get("name")})
        if len(page) < 1000:
            return names
        off += 1000


def same_as_remote(folder, name, entry):
    """True when the bucket already holds exactly this file. Uses the checksum the bucket reports (the
    MD5 of a single-part upload) when there is one, else the size. With neither, assume it differs."""
    with open(os.path.join(folder, name), "rb") as fh:
        raw = fh.read()
    meta = (entry or {}).get("metadata") or {}
    etag = s_(meta.get("eTag") or meta.get("etag")).strip('"')
    if re.fullmatch(r"[0-9a-f]{32}", etag):
        return hashlib.md5(raw).hexdigest() == etag
    size = meta.get("size", meta.get("contentLength"))
    return size is not None and int(size) == len(raw)


def upload(url, key, token, folder, name, replace):
    with open(os.path.join(folder, name), "rb") as fh:
        raw = fh.read()
    for attempt in range(3):
        st, j = http("POST", url + "/storage/v1/object/" + BUCKET + "/" + name, key, token, raw=raw,
                     ctype="image/jpeg", extra={"x-upsert": "true" if replace else "false"})
        if st in (200, 201):
            return "uploaded"
        if st == 409 or (st == 400 and "already exists" in s_(j).lower()):
            return "exists"
        # Storage reports a refused write as HTTP 400 with "row-level security" in the message, so look at the text too.
        if st in (401, 403) or "row-level security" in s_(j).lower():
            sys.exit("Upload refused (HTTP %d): %s\nThis account is not a photo administrator. Add it in the SQL editor "
                     "(see supabase/007_photos.sql) and run this again. Nothing has been uploaded by this run."
                     % (st, s_((j or {}).get("message") or j)))
        if st == 0 or st >= 500:
            time.sleep(2 * (attempt + 1))
            continue
        return "failed: HTTP %d %s" % (st, s_((j or {}).get("message") or j))
    return "failed: no response"


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = {a for a in sys.argv[1:] if a.startswith("--")}
    if len(args) != 1 or flags - {"--check", "--replace", "--changed"} or {"--replace", "--changed"} <= flags:
        sys.exit(__doc__)
    folder = os.path.abspath(args[0])
    if not os.path.isdir(folder):
        sys.exit("Not a folder: " + folder)
    if inside_repo(folder):
        sys.exit("Refusing: that folder is inside the repo. The photos are personal information and must stay outside it.")
    files, skipped = local_files(folder)
    print("%d photo file(s) ready; %d other file(s) ignored." % (len(files), skipped))
    if "--check" in flags or not files:
        return
    try:
        cfg = json.load(open(CFG_PATH, encoding="utf-8"))
    except (OSError, ValueError):
        sys.exit("Could not read pvh_local_config.json.")
    url, key, token = sign_in(cfg)
    have = remote_names(url, key, token)
    if "--replace" in flags:
        todo = files
    elif "--changed" in flags:
        todo = [n for n in files if n not in have or not same_as_remote(folder, n, have[n])]
    else:
        todo = [n for n in files if n not in have]
    new = [n for n in todo if n not in have]
    print("In the bucket already: %d. New: %d. Changed (replacing): %d. Unchanged, skipped: %d."
          % (len(set(have) & set(files)), len(new), len(todo) - len(new), len(files) - len(todo)))
    done = done_new = failed = 0
    for i, n in enumerate(todo, 1):
        r = upload(url, key, token, folder, n, n in have)     # overwrite only a photo that is already there
        if r in ("uploaded", "exists"):
            done += 1
            done_new += n not in have
        else:
            failed += 1
            print("  " + n + " -> " + r)
        if i % 50 == 0:
            print("  %d / %d" % (i, len(todo)))
    orphans = len(set(have) - set(files))
    print("Done. In the bucket now: %d of %d. Sent this run: %d. Failed: %d."
          % (len(set(have) & set(files)) + done_new, len(files), done, failed))
    if orphans:
        print("%d photo(s) in the bucket have no file in this folder (left alone)." % orphans)
    if failed:
        sys.exit(1)


main()
