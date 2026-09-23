# Metaculus leaderboards

Scrapes the Metaculus leaderboards (Baseline, Peer, Comments, Question writing — 1, 2, 5 and 10-year windows) into Excel workbooks, one sheet per day, highlighting rank changes versus the previous day.

## How it works

`scrap_metaculus.py` drives a dedicated Chrome profile through Selenium (remote debugging on port 9222). Using a real, logged-in browser avoids Cloudflare's captcha on automated sessions.

## Usage (Windows)

```bat
run.bat
```

The first run creates a `.venv`, installs `requirements.txt`, and opens Chrome on the leaderboard. Log into Metaculus and solve the captcha once in that window; later runs are unattended. Output goes to `scrape.log`.

Environment overrides: `METACULUS_DATA_DIR` (where workbooks are written), `CHROME_DEBUGGER_ADDRESS`, `CHROME_PROFILE_DIR`.

## Example output

See [`example/2026 Comments Ranking.xlsx`](example/).
