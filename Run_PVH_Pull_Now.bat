@echo off
REM ============================================================
REM  PVH - run the daily 07:30 pull and build right now
REM  Same thing pvh_daemon.py does each morning: pulls the reports
REM  from SSRS, refreshes PVH_data.json (--data-only, docs\ untouched)
REM  and copies it to Dropbox. The daemon's schedule is not affected.
REM  Don't start this around 07:30 -- two runs at once would both
REM  write the same files.
REM ============================================================

REM Run from this file's own folder no matter where it's launched from
cd /d "%~dp0"

echo.
echo === PVH pull and build (manual run) ===
echo Folder: %cd%
echo.

REM --- locate Python ---
set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY (
  where python >nul 2>nul && set "PY=python"
)
if not defined PY (
  echo [X] Python was not found on this PC.
  echo     Install Python from python.org, then run this again.
  goto :end
)

if not exist "pull_and_build_pvh.py" (
  echo [X] Missing: pull_and_build_pvh.py
  goto :end
)

REM --- pull and build ---
%PY% "pull_and_build_pvh.py"
set "RC=%errorlevel%"
echo.
if not "%RC%"=="0" (
  echo [X] Run failed ^(exit code %RC%^). Read the message above.
  goto :end
)

echo === Done ===
echo Full record of this run: today's logs\pull_and_build_YYYY-MM-DD.log

:end
echo.
pause
