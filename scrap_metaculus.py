from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
from datetime import date, timedelta
from pathlib import Path
import sys, time, openpyxl, os, locale
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Log in UTF-8 so non-ASCII usernames don't crash printing on Windows (cp1252).
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# Force a point as the decimal separator. Guarded: the locale may be missing on
# some machines, in which case format_score still handles "," via its own logic.
try:
    locale.setlocale(locale.LC_NUMERIC, 'en_US.UTF-8')
except locale.Error:
    pass

# Where the .xlsx workbooks are written. Override with METACULUS_DATA_DIR;
# defaults to the folder this script lives in, so the workbooks always sit
# next to scrap_metaculus.py regardless of where it's launched from.
DATA_DIR = Path(os.environ.get("METACULUS_DATA_DIR", Path(__file__).resolve().parent))

# Address of an already-running Chrome started with --remote-debugging-port.
# Attaching to your real, logged-in browser is what gets us past the Cloudflare
# captcha: from its point of view the requests come from you, not a bot.
DEBUGGER_ADDRESS = os.environ.get("CHROME_DEBUGGER_ADDRESS", "127.0.0.1:9222")

# Dedicated Chrome profile to launch if none is already listening. Separate from
# your everyday browsing so automation never touches your normal windows. First
# run opens on the leaderboard so you can log in + solve the captcha once.
CHROME_PROFILE_DIR = os.environ.get(
    "CHROME_PROFILE_DIR", str(Path.home() / ".metaculus-chrome-profile")
)

def format_score(score_str: str) -> str:
    score_str = score_str.strip()
    try:
        raw = score_str.lower()
        raw = raw.replace(" ", "").replace("\xa0", "").replace(" ", "")
        multiplier = 1.0
        if raw.endswith("k"):
            multiplier = 1000.0
            raw = raw[:-1]
        elif raw.endswith("m"):
            multiplier = 1000000.0
            raw = raw[:-1]

        if "," in raw and "." in raw:
            # Treat the last separator as decimal, strip the other as thousands.
            if raw.rfind(",") > raw.rfind("."):
                raw = raw.replace(".", "")
                raw = raw.replace(",", ".")
            else:
                raw = raw.replace(",", "")
        elif "," in raw:
            # If comma is likely thousands separator (e.g. 12,345), strip it.
            if len(raw.split(",")[-1]) == 3:
                raw = raw.replace(",", "")
            else:
                raw = raw.replace(",", ".")

        num = float(raw) * multiplier
        return str(int(num)) if num == int(num) else f"{num:.2f}"
    except ValueError:
        return "0"

def scraping(category, year, duration, attempts=3):
    url = f"https://www.metaculus.com/leaderboard/?category={category}&year={year}&duration={duration}"
    # Retry with backoff: a single Cloudflare/network hiccup used to permanently
    # lose a leaderboard for the day (one bad moment wiped 9/10 on 2026-06-28).
    # Transient timeouts now self-heal by reloading before giving up.
    last_err = None
    for attempt in range(1, attempts + 1):
        try:
            driver.get(url)
            # Wait for actual rows to render, not just the <table> shell (the React
            # app paints the table element before the data arrives).
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
            )
            time.sleep(0.5)  # let any remaining rows settle
            return BeautifulSoup(driver.page_source, 'html.parser')
        except TimeoutException as e:
            last_err = e
            if attempt < attempts:
                backoff = 3 * attempt  # 3s, then 6s
                print(f"    rows didn't render (attempt {attempt}/{attempts}); "
                      f"retrying in {backoff}s...")
                time.sleep(backoff)
    raise last_err

def createSheet(file_path):
    try:
        workbook = load_workbook(file_path)
    except FileNotFoundError:
        workbook = openpyxl.Workbook()

    sheet_name = f"{time.strftime('%d-%m')}"
    # If we already wrote a sheet for today, replace it so a re-run picks up the
    # latest data instead of doing nothing.
    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]

    return workbook.create_sheet(title=sheet_name, index=0), workbook

def extract_table_data(soup, is_peer=False):
    table = soup.find('table')
    if not table:
        print("    Table not found on the page")
        return []

    # Diagnostic: show the actual column headers so we can confirm the layout.
    header_cells = [th.get_text(strip=True) for th in table.find_all('th')]
    rows = table.find('tbody').find_all('tr')
    print(f"    columns={header_cells} | tbody rows={len(rows)}")

    data = []

    for row in rows:
        cells = row.find_all('td')
        if len(cells) < 4:
            continue

        rank = cells[0].find('span').text.strip() if cells[0].find('span') else ""
        name = cells[1].find('a').text.strip() if cells[1].find('a') else ""

        if is_peer:
            # For peer leaderboard, we expect specific formats
            questions = cells[2].text.strip()
            coverage = cells[3].text.strip()
            score_str = cells[4].text.strip() if len(cells) > 4 else "0"
            score_display = format_score(score_str)
            data.append([rank, name, questions, coverage, score_display, ""])
        else:
            # Direct approach: assuming score is in the 3rd column
            score_val = cells[3].text.strip() if len(cells) > 2 else "0"
            # Questions/numeric value typically in 4th column
            numeric_val = cells[2].text.strip() if len(cells) > 3 else ""

            score_display = format_score(score_val)
            data.append([rank, name, score_display, numeric_val, ""])

    if data:
        print(f"    extracted {len(data)} rows; first: {data[0][:5]}")
    return data

def loadPreviousDayData(workbook, min_row=2):
    if len(workbook.sheetnames) < 2:
        return {}

    prev_sheet = workbook[workbook.sheetnames[1]]
    headers = {cell.value: idx for idx, cell in enumerate(prev_sheet[1], start=1)}

    prev_data = {}
    for row in prev_sheet.iter_rows(min_row=min_row, values_only=True):
        name = row[headers["Name"] - 1]

        # Adjust based on table type
        if "Score" in headers:
            score = row[headers["Score"] - 1]
        else:
            score = row[4] if len(row) > 4 else 0  # For peer tables

        prev_data[name] = score

    return prev_data

def compareDataAndUpdateChanges(worksheet, prev_data, min_row=2, is_peer=False):
    headers = {cell.value: idx for idx, cell in enumerate(worksheet[1], start=1)}
    changes_detected = False

    # Determine which column to use as score based on table type
    score_column_name = "Score" if "Score" in headers else None

    for row_idx, row in enumerate(worksheet.iter_rows(min_row=min_row, values_only=False), start=min_row):
        name_cell = row[headers["Name"] - 1]
        name = name_cell.value

        # Get score based on table type
        if is_peer:
            # For peer tables
            score_cell = row[4] if len(row) > 4 else None
        else:
            # For other tables
            score_cell = row[headers["Score"] - 1] if score_column_name else None

        score = score_cell.value if score_cell else 0

        # Check if name exists in prev_data
        if name in prev_data:
            prev_score = prev_data[name]

            # Calculate difference
            try:
                if isinstance(score, (int, float)) and isinstance(prev_score, (int, float)):
                    difference = score - prev_score
                else:
                    difference = float(str(score).replace(',', '.')) - float(str(prev_score).replace(',', '.'))

                # Determine appropriate emoji
                if difference > 0:
                    emoji = "📈"
                    difference_str = f"{difference:+.2f}" if not float(difference).is_integer() else f"{int(difference):+d}"
                    changes_detected = True
                elif difference < 0:
                    emoji = "📉"
                    difference_str = f"{difference:+.2f}" if not float(difference).is_integer() else f"{int(difference):+d}"
                    changes_detected = True
                else:
                    emoji = "➖"
                    difference_str = ""

                # Write emoji and difference in the new column
                change_cell = worksheet.cell(row=row_idx, column=headers["Change"])
                change_cell.value = f"{emoji} {difference_str}".strip()
                change_cell.number_format = '@'
            except (TypeError, ValueError):
                worksheet.cell(row=row_idx, column=headers["Change"]).value = ""
        else:
            # If name is not in prev_data, consider it a new participant
            worksheet.cell(row=row_idx, column=headers["Change"]).value = "🆕 New"
            changes_detected = True  # New participant, so change detected

    return changes_detected

def process_data(file_path, data, is_peer=False):
    # An empty leaderboard (no entries scraped) should not be saved at all,
    # otherwise a header-only sheet gets written every day.
    if not data:
        print(f"No entries for {os.path.basename(file_path)}; skipping (nothing saved).")
        return "empty (not saved)"

    try:
        worksheet, workbook = createSheet(file_path)
    except ValueError:
        print(f"Sheet for today already exists in {file_path}. Skipping.")
        return "skipped (already ran today)"

    # Write headers based on data type
    if is_peer:
        worksheet.append(["Rank", "Name", "Questions", "Coverage", "Score", "Change"])
    else:
        worksheet.append(["Rank", "Name", "Score", "Questions", "Change"])

    # Write data
    for row in data:
        worksheet.append(row)

    # Load previous data and compare
    prev_data = loadPreviousDayData(workbook)
    changes_detected = compareDataAndUpdateChanges(worksheet, prev_data, is_peer=is_peer)

    # Delete sheet if no changes
    if not changes_detected and prev_data:
        print(f"No changes detected on {worksheet.title} from {os.path.basename(file_path)}. Deleting sheet.")
        workbook.remove(worksheet)
        return "no change (not saved)"
    else:
        print(f"Changes detected or first sheet on {worksheet.title} from {os.path.basename(file_path)}. Keeping sheet.")

        # Highlight "exmateriae"
        highlight_fill = PatternFill(start_color='a2fa9d', end_color='a2fa9d', fill_type='solid')
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.value == "exmateriae":
                    cell.fill = highlight_fill

        # Save the workbook
        try:
            workbook.save(file_path)
        except PermissionError:
            print(f"    PermissionError saving {os.path.basename(file_path)} "
                  f"— is it open in Excel? Close it and re-run.")
            raise
        return f"SAVED ({len(data)} rows)"

def scrapLeaderboard(category, year, duration, is_peer=False):
    # Adjust year if needed
    if duration == 2:
        year = 2032 if year > 2031 else 2030 if year > 2029 else 2028 if year > 2027 else 2026 if year > 2025 else 2024
    elif duration == 5:
        year = 2031 if year > 2030 else 2026 if year > 2025 else 2021
    elif duration == 10:
        year = 2026 if year > 2025 else 2016

    # Create file path (Mac/Linux friendly, under DATA_DIR)
    type_name = category.capitalize()
    filename = f'{year} '
    filename += f'{duration} Years ' if duration > 1 else ''
    filename += f'{type_name} Ranking.xlsx'
    file_path = str(DATA_DIR / filename)

    print(f"\n[{category} {year} {duration}yr] -> {filename}")

    # Get data
    soup = scraping(category, year, duration)
    data = extract_table_data(soup, is_peer)

    # Process data
    return process_data(file_path, data, is_peer)

def get_active_years(today):
    active_years = {today.year}
    if today <= date(today.year, 1, 1) + timedelta(days=100):
        active_years.add(today.year - 1)
    return sorted(active_years)

def _debugger_up(address, timeout=2):
    # DevTools answers /json/version once the debug port is listening.
    import urllib.request
    try:
        with urllib.request.urlopen(f"http://{address}/json/version", timeout=timeout):
            return True
    except Exception:
        return False

def _find_chrome():
    import shutil
    for cand in (
        os.path.join(os.environ.get("ProgramFiles", ""),
                     r"Google\Chrome\Application\chrome.exe"),
        os.path.join(os.environ.get("ProgramFiles(x86)", ""),
                     r"Google\Chrome\Application\chrome.exe"),
        os.path.join(os.environ.get("LocalAppData", ""),
                     r"Google\Chrome\Application\chrome.exe"),
    ):
        if cand and os.path.exists(cand):
            return cand
    return shutil.which("chrome") or shutil.which("google-chrome")

def _kill_process_tree(proc):
    """Hard-kill the Chrome we launched, including its child processes. Chrome's
    launcher spawns a separate browser process, so killing just `proc` can leave
    windows open — on Windows we use taskkill /T to take the whole tree."""
    if proc is None or proc.poll() is not None:
        return
    import subprocess
    try:
        if os.name == "nt":
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
        else:
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except Exception:
                proc.kill()
    except Exception:
        pass

def ensure_chrome(address=DEBUGGER_ADDRESS, wait=30):
    """Make the script self-sufficient: if no debuggable Chrome is listening,
    launch our dedicated profile and wait for the port. Lets you run
    `python scrap_metaculus.py` directly without run.bat starting Chrome first.

    Returns the subprocess.Popen handle of the Chrome we launched (so the caller
    can kill it when done), or None if a debuggable Chrome was already running
    (leave it be)."""
    if _debugger_up(address):
        return None
    chrome = _find_chrome()
    if not chrome:
        print("Chrome not found; install it or set CHROME_DEBUGGER_ADDRESS to a "
              "running debug Chrome.")
        return None
    port = address.rsplit(":", 1)[-1]
    print(f"No debug Chrome on {address}; launching {CHROME_PROFILE_DIR} ...")
    import subprocess
    proc = subprocess.Popen(
        [chrome, f"--remote-debugging-port={port}",
         f"--user-data-dir={CHROME_PROFILE_DIR}",
         "https://www.metaculus.com/leaderboard/"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    for _ in range(wait):
        if _debugger_up(address):
            print("  Chrome debug port is up.")
            return proc
        time.sleep(1)
    print(f"  Chrome didn't open the debug port within {wait}s; will try to "
          f"attach anyway.")
    return proc

if __name__ == "__main__":
    print(time.strftime("%D %H:%M:%S", time.localtime()))

    DATA_DIR.mkdir(parents=True, exist_ok=True)

    # Self-bootstrap: start the dedicated debug Chrome if none is listening, so
    # this script can be run directly (run.bat is now optional convenience).
    chrome_proc = ensure_chrome(DEBUGGER_ADDRESS)

    # Attach to the running, logged-in Chrome (started with --remote-debugging-port).
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", DEBUGGER_ADDRESS)
    try:
        driver = webdriver.Chrome(options=options)
    except Exception as e:
        print(
            f"Could not attach to Chrome at {DEBUGGER_ADDRESS}.\n"
            f"Start it first (see run.bat / README), then re-run. Error: {e}"
        )
        raise SystemExit(1)

    # Work in a dedicated tab so we don't hijack whatever you're browsing.
    original_handles = driver.window_handles
    driver.switch_to.new_window('tab')
    scrape_handle = driver.current_window_handle

    summary = []  # (label, outcome) per leaderboard, for the end-of-run audit

    def run_job(category, year, duration, is_peer=False):
        label = f"{category} {year} {duration}yr"
        try:
            status = scrapLeaderboard(category, year, duration, is_peer=is_peer)
            summary.append((label, status or "no status"))
        except Exception as e:
            print(f"  ! {label} failed: {e}")
            summary.append((label, f"FAILED: {type(e).__name__}"))

    try:
        # Process active periods (current year + previous year within 100-day grace)
        for year in get_active_years(date.today()):
            for duration in [1, 2, 5, 10]:
                run_job("baseline", year, duration)
                run_job("peer", year, duration, is_peer=True)

            run_job("comments", year, 1)
            run_job("questionWriting", year, 1)
    finally:
        if chrome_proc is not None:
            # We opened this Chrome ourselves, so shut the whole browser down
            # when the run finishes. driver.quit() alone does NOT kill a Chrome
            # we merely attached to, so ask DevTools to close the browser, then
            # kill the process tree we spawned as a hard fallback.
            try:
                driver.execute_cdp_cmd("Browser.close", {})
            except Exception:
                pass
            try:
                driver.quit()
            except Exception:
                pass
            _kill_process_tree(chrome_proc)
        else:
            # Chrome was already running (you started it / logged in yourself):
            # close just our scraping tab and leave the browser + your tabs be.
            try:
                driver.switch_to.window(scrape_handle)
                driver.close()
                if original_handles:
                    driver.switch_to.window(original_handles[0])
            except Exception:
                pass

    # End-of-run audit: one line per leaderboard so scrape.log is easy to read.
    print("\n===== RUN SUMMARY =====")
    for label, outcome in summary:
        print(f"  {label:<28} {outcome}")
    saved = sum(1 for _, o in summary if o.startswith("SAVED"))
    print(f"  -> {saved} saved, {len(summary) - saved} not saved, of {len(summary)} leaderboards")

    print(time.strftime("%D %H:%M:%S", time.localtime()))