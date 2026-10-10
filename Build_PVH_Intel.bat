@echo off
REM ============================================================
REM  PVH intel export - notes, stops and queries from Supabase,
REM  linked to the PVH records.
REM  Double-click it. You will be asked for your Supabase password
REM  (it is not saved). Output: intel\ (keep it on this PC).
REM  Run the app build first so PVH_data.json is current.
REM ============================================================

cd /d "%~dp0"

set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY (
  where python >nul 2>nul && set "PY=python"
)
if not defined PY (
  echo [X] Python was not found on this PC.
  goto :end
)

%PY% "pvh_intel.py" %*
if errorlevel 1 goto :end

echo.
echo Opening the intel folder...
start "" "%~dp0intel"

:end
echo.
pause
