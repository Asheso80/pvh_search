#!/usr/bin/env python3
r"""
pull_and_build_pvh.py

Refreshes the PVH Field Lookup data pipeline in one step:
  1. Pulls the taxi reports from SSRS as Excel, using your Windows login
     (Integrated Auth -- no password stored).
  2. Runs build_pvh_field_app.py to regenerate PVH_data.json,
     PVH_Field_App.html, and docs/ -- same as double-clicking
     Build_PVH_App.bat, just without the manual export/download step first.
  3. Copies PVH_data.json to your Dropbox folder, overwriting the existing
     copy so the shared link stays current.

Place this file in the same folder as build_pvh_field_app.py:
  ...\PVH_Projects\PVH_Search\

Requirements:
  pip install requests requests_negotiate_sspi

Safety: each download is validated (must actually be an .xlsx, not an HTML
parameter-prompt or error page) BEFORE it overwrites the existing file. If
any report fails validation, the run stops before touching PVH_data.json,
before rebuilding, and before touching the Dropbox copy -- a bad pull will
not silently corrupt the live data the officer team uses or gets shared.

Note on the Dropbox copy: PVH_data.json is unredacted owner/operator data
(names, phone, address, licence numbers) for the whole PVH dataset --
build_pvh_field_app.py's own comments flag it "keep private" / "do NOT
commit". Confirmed with the folder owner that the Dropbox destination is
access-restricted appropriately before relying on this.

Every run appends to logs\pull_and_build_YYYY-MM-DD.log so a scheduled,
unattended run leaves a record of what happened.
"""

import json
import logging
import os
import subprocess
import sys
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import requests

try:
    from requests_negotiate_sspi import HttpNegotiateAuth
except ImportError:
    print("Missing dependency. Run: pip install requests_negotiate_sspi")
    sys.exit(1)

# --------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent  # PVH_Search folder
BUILD_SCRIPT = SCRIPT_DIR / "build_pvh_field_app.py"
LOG_DIR = SCRIPT_DIR / "logs"
CONFIG_FILE = SCRIPT_DIR / "pvh_local_config.json"

REQUIRED_KEYS = ("server", "folder_path", "dropbox_target")


def load_config():
    """Read the machine-specific settings, failing loudly rather than guessing.

    The SSRS hostname, report folder and Dropbox destination differ per
    machine and together describe internal infrastructure, so they live in
    pvh_local_config.json (gitignored) rather than in this file, which is
    published. Copy pvh_local_config.example.json to start one.

    Every problem here exits non-zero with an explicit message. A scheduled
    07:30 run has nobody watching, so stopping outright is safer than
    falling back to a default that would pull from the wrong server or
    write unredacted data to the wrong path.
    """
    if not CONFIG_FILE.exists():
        sys.exit(
            f"Missing {CONFIG_FILE.name}. Copy pvh_local_config.example.json to "
            f"{CONFIG_FILE.name} and fill in the SSRS server, report folder and "
            f"Dropbox path."
        )
    try:
        cfg = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except ValueError as exc:
        sys.exit(f"{CONFIG_FILE.name} is not valid JSON: {exc}")
    missing = [k for k in REQUIRED_KEYS if not str(cfg.get(k, "")).strip()]
    if missing:
        sys.exit(f"{CONFIG_FILE.name} is missing a value for: {', '.join(missing)}")
    return cfg


_CFG = load_config()
SERVER = _CFG["server"]
REPORT_SERVER_URL = f"http://{SERVER}/ReportServer"
FOLDER_PATH = _CFG["folder_path"]
DROPBOX_TARGET = Path(_CFG["dropbox_target"])

# SSRS report name -> local filename (matches what's already in this folder).
REPORTS = {
    "All_Active_Vehicles": "All_Active_Vehicles.xlsx",              # -> vehicles (required by build)
    "OperatorList": "OperatorList.xlsx",                            # -> operators (required by build)
    "All_Vehicle_LastInspection": "All_Vehicle_LastInspection.xlsx",  # -> addresses (optional, used if present)
    "OwnersAndVehicles": "OwnersAndVehicles.xlsx",                  # refreshed for archive -- NOT passed to the build
}
NOT_CONSUMED_BY_BUILD = {"OwnersAndVehicles"}

REQUEST_TIMEOUT = 180  # seconds per report
# --------------------------------------------------------------------------


def setup_logging():
    LOG_DIR.mkdir(exist_ok=True)
    log_file = LOG_DIR / f"pull_and_build_{datetime.now():%Y-%m-%d}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(message)s",
        datefmt="%H:%M:%S",
        handlers=[logging.FileHandler(log_file, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
    )
    return logging.getLogger("pull_and_build")


def get_session() -> requests.Session:
    s = requests.Session()
    s.auth = HttpNegotiateAuth()  # uses your current Windows login
    return s


def fetch_excel(session: requests.Session, report_name: str) -> bytes:
    item_path = f"{FOLDER_PATH}/{report_name}"
    url = f"{REPORT_SERVER_URL}?{quote(item_path)}&rs:Command=Render&rs:Format=EXCELOPENXML"
    resp = session.get(url, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    content = resp.content
    # A real .xlsx is a zip archive and always starts with "PK". A
    # parameter-prompt or error page comes back as HTML instead -- catch
    # that here rather than writing garbage into the pipeline.
    if not content.startswith(b"PK"):
        snippet = content[:200].decode("utf-8", errors="replace")
        raise ValueError(
            f"response wasn't a valid .xlsx (likely needs input parameters "
            f"with no default, or an auth/permission issue). "
            f"First 200 bytes: {snippet!r}"
        )
    return content


def write_source(path: Path, content: bytes) -> None:
    """Land a refreshed export by rename rather than truncating it in place.

    Same reason as write_out() in build_pvh_field_app.py: these files are
    OneDrive placeholders, and truncating one can fail mid-run with OSError 22.
    A half-written .xlsx would also be worse than no write at all.
    """
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), prefix="._pvh_", suffix=".tmp")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(content)
        os.replace(tmp, path)
    except BaseException:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


def main():
    log = setup_logging()
    log.info("=== Run started ===")
    log.info(f"Pulling {len(REPORTS)} report(s) from {SERVER}{FOLDER_PATH} ...")
    session = get_session()

    fetched, errors = {}, []
    for report_name in REPORTS:
        try:
            fetched[report_name] = fetch_excel(session, report_name)
            log.info(f"  OK   {report_name}")
        except (requests.exceptions.RequestException, ValueError) as exc:
            errors.append(f"{report_name}: {exc}")
            log.error(f"  FAIL {report_name}: {exc}")

    if errors:
        log.error(
            "One or more reports failed to pull. Stopping before touching any "
            "existing files, rebuilding, or updating the Dropbox copy -- the "
            "last known-good data stays intact and stays what's shared."
        )
        for e in errors:
            log.error(f"  - {e}")
        sys.exit(1)

    # All downloads validated -- safe to overwrite in place.
    for report_name, filename in REPORTS.items():
        write_source(SCRIPT_DIR / filename, fetched[report_name])
    log.info("All source files updated.")

    consumed = [n for n in REPORTS if n not in NOT_CONSUMED_BY_BUILD]
    log.info(
        f"Note: OwnersAndVehicles.xlsx was refreshed but build_pvh_field_app.py "
        f"doesn't read it -- only {', '.join(consumed)} feed the build."
    )

    log.info("Rebuilding PVH_data.json / PVH_Field_App.html / docs/ ...")
    result = subprocess.run(
        [
            sys.executable, str(BUILD_SCRIPT),
            "All_Active_Vehicles.xlsx", "OperatorList.xlsx",
            "All_Vehicle_LastInspection.xlsx", "PVH_Field_App.html",
        ],
        cwd=str(SCRIPT_DIR),
        capture_output=True, text=True,
    )
    for line in result.stdout.splitlines():
        log.info(f"  [build] {line}")
    for line in result.stderr.splitlines():
        log.error(f"  [build] {line}")
    if result.returncode != 0:
        log.error("Build step failed -- see [build] lines above. Dropbox copy skipped.")
        sys.exit(result.returncode)

    # Build succeeded -- push the fresh PVH_data.json to Dropbox, overwriting
    # in place so the existing shared link keeps pointing at current data.
    try:
        DROPBOX_TARGET.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(SCRIPT_DIR / "PVH_data.json", DROPBOX_TARGET)
        log.info(f"Copied PVH_data.json -> {DROPBOX_TARGET}")
    except OSError as exc:
        log.error(f"Build succeeded but Dropbox copy failed: {exc}")
        log.error("PVH_data.json and PVH_Field_App.html are current locally; the "
                   "shared Dropbox link was NOT updated this run.")
        sys.exit(1)

    log.info("=== Done. PVH_data.json, PVH_Field_App.html, docs/, and the Dropbox copy are all current. ===")


if __name__ == "__main__":
    main()
