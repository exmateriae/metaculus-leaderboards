@echo off
REM Scrape Metaculus leaderboards into Excel, locally on Windows.
REM
REM Attaches to a dedicated, logged-in Chrome started with remote debugging, so
REM Cloudflare sees you (home IP, real login) and does not captcha-gate it the
REM way a fresh automated Chromium gets gated.
REM
REM FIRST-TIME SETUP: a Chrome window opens on the leaderboard. Log into
REM Metaculus once in that window and solve the captcha; the dedicated profile
REM then stays logged in. After that, runs are unattended and the window closes
REM on its own when done. The full run output is written to scrape.log.

setlocal
cd /d "%~dp0"

set "PORT=9222"

REM Workbooks default to this folder; set METACULUS_DATA_DIR to override.

REM NOTE: Chrome is started (and, when it owns it, closed) by scrap_metaculus.py
REM itself (see ensure_chrome() there) -- deliberately NOT here. If run.bat also
REM pre-started Chrome, the script would find the debug port already up, assume
REM it's attaching to your personal browser, and leave it running forever
REM instead of closing it when the run finishes.

REM Set up the virtualenv on first run.
if not exist ".venv\Scripts\python.exe" (
    python -m venv .venv
    ".venv\Scripts\python.exe" -m pip install --quiet --upgrade pip
    ".venv\Scripts\python.exe" -m pip install --quiet -r requirements.txt
)

REM Run the scraper. Output goes to scrape.log; the window closes when done.
set "CHROME_DEBUGGER_ADDRESS=127.0.0.1:%PORT%"
".venv\Scripts\python.exe" scrap_metaculus.py > scrape.log 2>&1

endlocal