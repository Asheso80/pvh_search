@echo off
rem Serves the test build (docs\) on this PC only (127.0.0.1:8777).
rem Phones reach it through Tailscale: see the one-time command in the notes.
cd /d "%~dp0docs"
echo Test server for the operator-vehicle links branch.
echo Close this window to stop it.
py -m http.server 8777 --bind 127.0.0.1
