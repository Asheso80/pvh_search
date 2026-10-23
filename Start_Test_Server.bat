@echo off
rem Try the app with MADE-UP records and a pretend shared-notes server
rem (test_server.py), on this PC only: http://127.0.0.1:8795/
rem Nothing real is read or sent, so testing never puts marks, notes or links
rem in the real database. Phones: see the note at the top of test_server.py.
cd /d "%~dp0"
set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY (
  where python >nul 2>nul && set "PY=python"
)
if not defined PY (
  echo [X] Python was not found on this PC.
  pause
  exit /b 1
)
%PY% test_server.py %*
pause
