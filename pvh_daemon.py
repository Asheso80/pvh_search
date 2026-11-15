#!/usr/bin/env python3
r"""
pvh_daemon.py

Workaround for restricted Task Scheduler access: runs the PVH pull/build
pipeline once daily at 07:30 without needing "Run whether logged on or
not," admin rights, or the "Log on as a batch job" right. Starts
automatically via your personal Startup folder (no IT approval needed)
and just waits from there.

One-time setup:
  1. Press Win+R, type: shell:startup , Enter -- opens your Startup folder.
  2. Right-click inside it -> New -> Shortcut. Target:
       pythonw.exe "<full path to this file>"
     (pythonw.exe, not python.exe -- runs silently, no console window.)
  3. Double-click the new shortcut once now to start it immediately rather
     than waiting for your next login. Check logs\daemon.log for a
     "Waiting until ..." line to confirm it's alive.

Behaviour:
  - Runs indefinitely once started -- no need to relaunch daily. It only
    stops if you log out, shut down, or reboot; the Startup shortcut
    restarts it automatically at your next login.
  - Only one copy runs at a time (file lock). Logging in again while it's
    already running, or the Startup shortcut firing twice, will not
    double-run the pull.
  - If you log out (not just lock the screen) or the machine restarts and
    you're not back before 07:30, that day's run happens a few minutes after
    you next sign in instead. The same catch-up covers a 07:30 run that
    failed, the next time the daemon starts that day. A day with no sign-in
    at all still has no run; the app warns officers once its data is three
    days old.
  - logs\daemon.log only says when each run happened and how it ended. What
    the run did -- including the build's data warnings, which name owners
    and operators -- is in that day's pull_and_build log, kept 90 days. This
    log starts a new file each day and keeps the same 90.
"""

import sys
import time
import logging
import logging.handlers
import subprocess
import msvcrt
from datetime import datetime, timedelta
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
PULL_SCRIPT = SCRIPT_DIR / "pull_and_build_pvh.py"
LOG_DIR = SCRIPT_DIR / "logs"
LOCK_FILE = SCRIPT_DIR / "daemon.lock"
RUN_AT_HOUR, RUN_AT_MINUTE = 7, 30  # 7:30 AM -- adjust if you want it closer to 8:00
LOG_KEEP_DAYS = 90  # same as pull_and_build_pvh.py
CATCH_UP_DELAY = 180  # seconds to wait after a late start before the catch-up run


def setup_logging():
    LOG_DIR.mkdir(exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.handlers.TimedRotatingFileHandler(
                LOG_DIR / "daemon.log", when="midnight", backupCount=LOG_KEEP_DAYS, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger("pvh_daemon")


def acquire_lock():
    """Exclusive, non-blocking lock held for the process's entire lifetime --
    released automatically on exit or crash (no stale-PID bookkeeping needed).
    A second instance starting up fails to acquire it and exits quietly."""
    fh = open(LOCK_FILE, "w")
    fh.write("locked\n")
    fh.flush()
    try:
        msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
    except OSError:
        fh.close()
        return None
    return fh  # caller must keep this open for as long as the daemon runs


def next_run_time(now):
    target = now.replace(hour=RUN_AT_HOUR, minute=RUN_AT_MINUTE, second=0, microsecond=0)
    if target <= now:
        target += timedelta(days=1)
    return target


def ran_today(now):
    """True when today's pull and build has already finished properly -- by the
    daemon or by Run_PVH_Pull_Now.bat. Goes by the last line a good run writes
    to its dated log, so a run that failed earlier today does not count."""
    path = LOG_DIR / f"pull_and_build_{now:%Y-%m-%d}.log"
    try:
        return "=== Done." in path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return False


def catch_up_due(now):
    """Started after today's run time with no good run today: the PC was off or
    signed out at 07:30, or the run failed. Without this the phones would keep
    yesterday's data until tomorrow morning."""
    target = now.replace(hour=RUN_AT_HOUR, minute=RUN_AT_MINUTE, second=0, microsecond=0)
    return now >= target and not ran_today(now)


def run_pipeline(log):
    python_exe = sys.executable.replace("pythonw.exe", "python.exe")
    log.info("Target time reached -- running pull_and_build_pvh.py")
    result = subprocess.run(
        [python_exe, str(PULL_SCRIPT)],
        cwd=str(SCRIPT_DIR), capture_output=True, text=True,
    )
    # The run's own output is already in its dated log; copying it here would
    # keep the names in it past that log's 90 days. Errors Python itself
    # reports (a crash before the run could log anything) are still kept.
    for line in result.stderr.splitlines():
        log.error(f"  [pull] {line}")
    log.info(f"Run finished with exit code {result.returncode} -- "
             f"details in logs\\pull_and_build_{datetime.now():%Y-%m-%d}.log")


def main():
    log = setup_logging()
    lock = acquire_lock()
    if lock is None:
        log.info("Another instance is already running -- exiting.")
        return
    log.info("=== PVH daemon started ===")

    if catch_up_due(datetime.now()):
        # Just signed in: give the network, OneDrive and Dropbox a few minutes
        # to come up before pulling.
        log.info(f"Started after {RUN_AT_HOUR:02d}:{RUN_AT_MINUTE:02d} and today's run has not finished -- "
                 f"running it in {CATCH_UP_DELAY // 60} minutes.")
        time.sleep(CATCH_UP_DELAY)
        if catch_up_due(datetime.now()):   # unless it was run by hand meanwhile
            run_pipeline(log)

    while True:
        target = next_run_time(datetime.now())
        log.info(f"Waiting until {target:%Y-%m-%d %H:%M} ...")
        while True:
            remaining = (target - datetime.now()).total_seconds()
            if remaining <= 0:
                break
            # Wake at least hourly rather than one long sleep, so the loop
            # stays responsive to clock changes (DST, manual clock edits).
            time.sleep(min(remaining, 3600))
        run_pipeline(log)


if __name__ == "__main__":
    main()
