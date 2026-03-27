import os
import re
import json
import hashlib
import argparse
import pyperclip
import subprocess
import time
import pyautogui
import ctypes
import sys
from datetime import datetime, timedelta, timezone
from playwright.sync_api import sync_playwright

try:
    from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
except Exception:
    PlaywrightTimeoutError = Exception

try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except ImportError:
    Request = None
    Credentials = None
    InstalledAppFlow = None
    build = None
    HttpError = None

# --- 1. DYNAMIC CONFIGURATION ---
# Safely hide the browser profile in the user's hidden AppData folder
PROFILE_DIR = os.path.join(os.environ['LOCALAPPDATA'], "Classflow")

# Store generated files beside the Classflow folder (workspace root).
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.dirname(SCRIPT_DIR)

# Downloads remain on Desktop for easy access.
USER_DESKTOP = os.path.join(os.environ['USERPROFILE'], 'Desktop')

DOWNLOAD_DIR = os.path.join(USER_DESKTOP, "Assignments")
DEADLINE_FILE = os.path.join(OUTPUT_DIR, "deadlines.txt")
HISTORY_FILE = os.path.join(OUTPUT_DIR, "assignment_history.json")
GOOGLE_SETUP_FILE = os.path.join(PROFILE_DIR, "google_calendar_setup.flag")
TEAMS_SETUP_FILE = os.path.join(PROFILE_DIR, "teams_setup.flag")
GOOGLE_TOKEN_FILE = os.path.join(PROFILE_DIR, "google_token.json")
GOOGLE_SCOPES = ["https://www.googleapis.com/auth/calendar.events"]
GOOGLE_CLIENT_SECRET_CANDIDATES = [
    "google_client_secret.json",
    "credentials.json",
    "client_secret.json"
]
# Set this to your target Google Calendar ID to avoid using My Calendar.
# Example: "abc123@group.calendar.google.com"
GOOGLE_CALENDAR_ID = "primary"
TASK_NAME_NOON = "Classflow Daily 12PM"
TASK_NAME_EVENING = "Classflow Daily 6PM"

if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

# --- YOUR CUSTOM NAMING CONVENTION ---
COURSE_MAP = {
    "CS224": "FLAT",
    "CS272": "HCI",
    "CE222": "COAL",
    "CS232": "DBMS Lab"
}

def load_history():
    """Load assignment history from JSON file."""
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_history(history):
    """Save assignment history to JSON file."""
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, indent=2, ensure_ascii=False)

def normalize_text(value):
    """Collapse repeated whitespace and trim leading/trailing spaces."""
    return re.sub(r"\s+", " ", (value or "")).strip()

def strip_course_prefix(title):
    """Remove course codes and names from the beginning of a title."""
    # Remove patterns like "CS272 -" or "CS272:" or "[CS272]"
    title = re.sub(r"^\s*\[?[A-Z]+\d+\]?\s*[-:]\s*", "", title)
    # Also check for custom course names from COURSE_MAP
    for key, clean_name in COURSE_MAP.items():
        # Remove course code or clean name from the start: "CS272 - Title" or "FLAT - Title"
        title = re.sub(rf"^\s*{re.escape(key)}\s*[-:]\s*", "", title, flags=re.IGNORECASE)
        title = re.sub(rf"^\s*{re.escape(clean_name)}\s*[-:]\s*", "", title, flags=re.IGNORECASE)
    return normalize_text(title)

def is_teams_setup_complete():
    """Check whether one-time Teams setup has been completed."""
    return os.path.exists(TEAMS_SETUP_FILE)

def mark_teams_setup_complete():
    """Persist Teams setup completion marker."""
    os.makedirs(PROFILE_DIR, exist_ok=True)
    with open(TEAMS_SETUP_FILE, "w", encoding="utf-8") as f:
        f.write(datetime.now().isoformat())

def clean_date_string(raw_text):
    """Just strips the word 'Due ' from the text."""
    return raw_text.replace("Due ", "").strip()

def parse_due_date(raw_due_date):
    """Parse Teams due date text into a datetime when possible."""
    if not raw_due_date or raw_due_date == "No date specified":
        return None

    cleaned = raw_due_date.replace(" at ", " ").strip()
    cleaned = re.sub(r"\s+", " ", cleaned)

    date_formats = [
        "%A, %B %d, %Y %I:%M %p",
        "%a, %B %d, %Y %I:%M %p",
        "%A, %d %B %Y %I:%M %p",
        "%B %d, %Y %I:%M %p",
        "%A, %B %d, %Y",
        "%B %d, %Y"
    ]

    for fmt in date_formats:
        try:
            parsed = datetime.strptime(cleaned, fmt)
            return parsed
        except ValueError:
            continue

    return None

# Detect if running as frozen EXE
IS_FROZEN = getattr(sys, "frozen", False)
EXE_LOG_FILE = os.path.join(os.path.expanduser("~"), "Desktop", "Classflow_Log.txt") if IS_FROZEN else None

def show_windows_popup(title, message):
    """Show a Windows popup message and log to file if running as EXE."""
    ctypes.windll.user32.MessageBoxW(0, message, title, 0)
    # Log to file for exe mode for debugging
    if IS_FROZEN and EXE_LOG_FILE:
        try:
            with open(EXE_LOG_FILE, "a", encoding="utf-8") as f:
                f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {title}: {message}\n")
        except:
            pass

def log_output(message, show_popup=False, title="Classflow"):
    """Log message to console (IDE) or popup (EXE). Set show_popup=True for user-facing messages."""
    print(message)
    if IS_FROZEN:
        # Log to file when running as EXE
        try:
            os.makedirs(os.path.dirname(EXE_LOG_FILE), exist_ok=True)
            with open(EXE_LOG_FILE, "a", encoding="utf-8") as f:
                f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")
        except:
            pass
        # Show popup for important user-facing messages
        if show_popup:
            show_windows_popup(title, message)

def build_scheduler_action():
    """Return the Task Scheduler action string for script or frozen EXE mode."""
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    return f'"{sys.executable}" "{os.path.abspath(__file__)}"'

def create_windows_task(task_name, start_time):
    """Create or overwrite a daily Windows task for the provided time."""
    action = build_scheduler_action()
    command = [
        "schtasks",
        "/Create",
        "/F",
        "/SC",
        "DAILY",
        "/TN",
        task_name,
        "/ST",
        start_time,
        "/RL",
        "LIMITED",
        "/TR",
        action,
    ]
    subprocess.run(command, check=True, capture_output=True, text=True)

def delete_windows_task(task_name):
    """Delete a Windows scheduled task if it exists."""
    command = ["schtasks", "/Delete", "/F", "/TN", task_name]
    result = subprocess.run(command, capture_output=True, text=True)
    if result.returncode != 0 and "cannot find" not in result.stderr.lower():
        raise RuntimeError(result.stderr.strip() or result.stdout.strip())

def install_schedule():
    """Install two daily tasks at 12:00 and 18:00."""
    create_windows_task(TASK_NAME_NOON, "12:00")
    create_windows_task(TASK_NAME_EVENING, "18:00")

def remove_schedule():
    """Remove both daily Classflow tasks."""
    delete_windows_task(TASK_NAME_NOON)
    delete_windows_task(TASK_NAME_EVENING)

def is_google_setup_complete():
    """Check whether one-time Google Calendar setup has been completed."""
    return os.path.exists(GOOGLE_SETUP_FILE)

def mark_google_setup_complete():
    """Persist Google Calendar setup completion marker."""
    os.makedirs(PROFILE_DIR, exist_ok=True)
    with open(GOOGLE_SETUP_FILE, "w", encoding="utf-8") as f:
        f.write(datetime.now().isoformat())

def resolve_google_client_secret_path():
    """Find OAuth client JSON in script directory."""
    for name in GOOGLE_CLIENT_SECRET_CANDIDATES:
        candidate = os.path.join(SCRIPT_DIR, name)
        if os.path.exists(candidate):
            return candidate

    try:
        for file_name in os.listdir(SCRIPT_DIR):
            if not file_name.lower().endswith(".json"):
                continue
            full_path = os.path.join(SCRIPT_DIR, file_name)
            with open(full_path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            if isinstance(payload, dict) and ("installed" in payload or "web" in payload):
                return full_path
    except Exception:
        return None

    return None

def get_google_calendar_service(interactive_auth):
    """Create an authenticated Google Calendar service instance."""
    if not all([Request, Credentials, InstalledAppFlow, build]):
        show_windows_popup(
            "Classflow Google API Error",
            "Google API libraries are missing. Install:\n"
            "pip install google-api-python-client google-auth google-auth-oauthlib"
        )
        return None

    client_secret_path = resolve_google_client_secret_path()
    if not client_secret_path:
        show_windows_popup(
            "Classflow Google API Error",
            "Google OAuth client JSON was not found in the Classflow folder."
        )
        return None

    creds = None
    if os.path.exists(GOOGLE_TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_FILE, GOOGLE_SCOPES)
        except Exception:
            creds = None

    try:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        elif (not creds or not creds.valid) and interactive_auth:
            flow = InstalledAppFlow.from_client_secrets_file(client_secret_path, GOOGLE_SCOPES)
            creds = flow.run_local_server(port=0)

        if not creds or not creds.valid:
            return None

        os.makedirs(PROFILE_DIR, exist_ok=True)
        with open(GOOGLE_TOKEN_FILE, "w", encoding="utf-8") as token_file:
            token_file.write(creds.to_json())

        return build("calendar", "v3", credentials=creds, cache_discovery=False)
    except Exception as e:
        log_output(f"Google Calendar API authentication failed: {e}", show_popup=IS_FROZEN, title="Google Calendar Error")
        return None

def build_google_event(unique_name, due_text):
    """Build a Google Calendar event payload from one assignment deadline."""
    parsed_due = parse_due_date(due_text)
    if parsed_due is None:
        return None

    event = {
        "summary": unique_name,
        "description": f"Assignment deadline: {due_text}"
    }

    if parsed_due.hour == 0 and parsed_due.minute == 0 and "am" not in due_text.lower() and "pm" not in due_text.lower():
        event["start"] = {"date": parsed_due.strftime("%Y-%m-%d")}
        event["end"] = {"date": (parsed_due + timedelta(days=1)).strftime("%Y-%m-%d")}
    else:
        local_tz = datetime.now().astimezone().tzinfo
        if parsed_due.tzinfo is None:
            parsed_due = parsed_due.replace(tzinfo=local_tz)

        due_utc = parsed_due.astimezone(timezone.utc)
        start_utc = (parsed_due - timedelta(minutes=30)).astimezone(timezone.utc)
        event["start"] = {"dateTime": start_utc.strftime("%Y-%m-%dT%H:%M:%SZ")}
        event["end"] = {"dateTime": due_utc.strftime("%Y-%m-%dT%H:%M:%SZ")}

    return event

def make_google_event_id(unique_name):
    """Build a deterministic Google Calendar event id for upsert behavior."""
    digest = hashlib.sha1(unique_name.encode("utf-8")).hexdigest()
    return f"cf{digest}"

def sync_deadlines_to_google_calendar(deadlines, interactive_auth):
    """Insert/update changed deadlines in Google Calendar through API."""
    if not deadlines:
        return {"ok": True, "inserted": 0, "updated": 0, "skipped": 0, "failed": 0}

    service = get_google_calendar_service(interactive_auth=interactive_auth)
    if not service:
        return {"ok": False, "inserted": 0, "updated": 0, "skipped": 0, "failed": len(deadlines)}

    calendar_id = os.environ.get("CLASSFLOW_CALENDAR_ID", GOOGLE_CALENDAR_ID)
    inserted = 0
    updated = 0
    skipped = 0
    failed = 0

    for unique_name, due_text in sorted(deadlines.items()):
        event_body = build_google_event(unique_name, due_text)
        if not event_body:
            skipped += 1
            continue

        event_id = make_google_event_id(unique_name)
        event_body["id"] = event_id

        try:
            try:
                service.events().get(calendarId=calendar_id, eventId=event_id).execute()
                service.events().update(calendarId=calendar_id, eventId=event_id, body=event_body).execute()
                updated += 1
            except Exception as err:
                is_not_found = False
                if HttpError and isinstance(err, HttpError):
                    is_not_found = getattr(err, "status_code", None) == 404 or (getattr(err, "resp", None) and err.resp.status == 404)
                if is_not_found:
                    service.events().insert(calendarId=calendar_id, body=event_body).execute()
                    inserted += 1
                else:
                    raise
        except Exception as e:
            failed += 1
            log_output(f"Google Calendar API sync failed for '{unique_name}': {e}", show_popup=False)

    return {
        "ok": failed == 0,
        "inserted": inserted,
        "updated": updated,
        "skipped": skipped,
        "failed": failed
    }

def open_teams_and_wait_for_assignments(page, setup_mode=False):
    """Open Teams and wait until the Assignments app is available."""
    log_output("Opening Microsoft Teams...", show_popup=IS_FROZEN, title="Classflow")
    page.set_default_navigation_timeout(60000)
    page.set_default_timeout(600000 if setup_mode else 60000)

    try:
        page.goto("https://teams.microsoft.com/v2/", timeout=60000)
    except Exception as nav_err:
        log_output(f"Teams navigation timeout, retrying once: {nav_err}", show_popup=False)
        page.goto("https://teams.microsoft.com/v2/", timeout=60000)

    if setup_mode:
        show_windows_popup(
            "Teams Setup",
            "Please sign in to Microsoft Teams in the opened Chrome window and complete any prompts."
        )
        log_output("Teams setup required. Please complete login in the opened Chrome window.", show_popup=IS_FROZEN, title="Classflow Setup")

    log_output("Waiting for Teams Assignments to become available...", show_popup=False)
    total_timeout = 600000 if setup_mode else 120000
    max_probe_timeout = 10000
    selectors = [
        lambda: page.get_by_role("button", name="Assignments (Ctrl+Shift+4)").first,
        lambda: page.get_by_role("button", name=re.compile(r"Assignments", re.IGNORECASE)).first,
        lambda: page.locator("[data-tid='app-bar-edu-assignments']").first,
    ]

    deadline = time.time() + (total_timeout / 1000)
    attempt = 1

    while time.time() < deadline:
        if page.is_closed():
            log_output("Teams page was closed before assignments loaded.", show_popup=IS_FROZEN, title="Teams Error")
            return False

        for selector_factory in selectors:
            try:
                remaining_ms = max(1000, int((deadline - time.time()) * 1000))
                probe_timeout = min(max_probe_timeout, remaining_ms)
                assignments_btn = selector_factory()
                assignments_btn.wait_for(state="visible", timeout=probe_timeout)
                assignments_btn.click()
                return True
            except PlaywrightTimeoutError:
                continue
            except Exception as e:
                # Target/page closure can happen intermittently when Teams reloads.
                if "Target page, context or browser has been closed" in str(e):
                    log_output("Teams browser context closed while waiting for Assignments.", show_popup=IS_FROZEN, title="Teams Error")
                    return False
                continue

        if time.time() >= deadline:
            break

        log_output(f"Assignments button not ready yet (attempt {attempt}), retrying...", show_popup=False)
        attempt += 1
        try:
            page.wait_for_timeout(1500)
            page.reload(timeout=60000)
        except Exception:
            pass

    log_output("Could not open Teams Assignments. Please verify login and try again.", show_popup=IS_FROZEN, title="Teams Error")
    return False

def extract_assignment_title(iframe, fallback_title):
    """Prefer detail-page heading text to avoid card-title truncation."""
    fallback = normalize_text(fallback_title)
    candidates = [fallback]

    selectors = [
        "h1",
        "[role='heading'][aria-level='1']",
        "[id*='title']",
        "[data-tid*='title']",
        "[data-test*='title']"
    ]

    for selector in selectors:
        try:
            texts = iframe.locator(selector).all_inner_texts()
            for text in texts:
                # Only take the first line to avoid combining multiple lines
                first_line = text.split('\n')[0]
                cleaned = normalize_text(first_line)
                if cleaned and len(cleaned) > 3 and not cleaned.lower().startswith("due "):
                    candidates.append(cleaned)
        except Exception:
            continue

    best = fallback
    for candidate in candidates:
        if len(candidate) > len(best):
            best = candidate

    # Strip course prefix from extracted title to avoid duplication
    best = strip_course_prefix(best)
    return best

def format_assignment_name(course_name, assignment_title):
    """Use clean display naming without square brackets."""
    return f"{course_name} - {assignment_title}"

def first_time_setup():
    """Complete one-time Teams and Google setup, then ask user to rerun."""
    log_output("First-time setup detected. Starting account setup...", show_popup=IS_FROZEN, title="Classflow Setup")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            channel="chrome",
            headless=False,
            accept_downloads=True
        )
        page = context.pages[0] if context.pages else context.new_page()

        teams_ok = is_teams_setup_complete()
        if not teams_ok:
            teams_ok = open_teams_and_wait_for_assignments(page, setup_mode=True)
            if teams_ok:
                mark_teams_setup_complete()

        google_ok = is_google_setup_complete()
        if not google_ok:
            google_ok = get_google_calendar_service(interactive_auth=True) is not None
            if google_ok:
                mark_google_setup_complete()

        context.close()

    show_windows_popup(
        "Classflow Teams Login",
        "Teams login/setup successful." if teams_ok else "Teams login/setup failed. Please rerun and complete Teams login."
    )

    show_windows_popup(
        "Classflow Google Login",
        "Google login/setup successful." if google_ok else "Google login/setup failed. Please rerun and complete Google login."
    )

    if teams_ok and google_ok:
        show_windows_popup(
            "Classflow Setup Complete",
            "Teams and Google Calendar setup are complete.\n\nPlease run Classflow again to start assignment tracking."
        )
        return True

    show_windows_popup(
        "Classflow Setup Incomplete",
        "Setup did not finish successfully. Please run Classflow again and complete login prompts."
    )
    return False


def run():
    if not is_teams_setup_complete() or not is_google_setup_complete():
        first_time_setup()
        return
    
    master_deadlines = {}
    history = load_history()
    assignments_to_process = []
    changed_deadlines = {}

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            channel="chrome",
            headless=True,
            accept_downloads=True
        )
        page = context.pages[0] if context.pages else context.new_page()

        if not open_teams_and_wait_for_assignments(page, setup_mode=False):
            context.close()
            return

        log_output("Waiting for assignment list to fetch...", show_popup=False)
        iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame
        target_card_id = "[id*='CardHeader__headerEDUASSIGN']:visible"
        iframe.locator(target_card_id).first.wait_for(state="visible", timeout=30000)

        assignment_cards = iframe.locator(target_card_id)
        total_assignments = assignment_cards.count()
        
        log_output(f"Found {total_assignments} upcoming assignment(s).", show_popup=IS_FROZEN, title="Classflow")
        
        for i in range(total_assignments):
            iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame
            current_card = iframe.locator(target_card_id).nth(i)

            full_card_text = current_card.locator("xpath=..").inner_text()
            assignment_title = strip_course_prefix(normalize_text(full_card_text.split('\n')[0]))
            
            course_name = "Other" 
            for key, clean_name in COURSE_MAP.items():
                if key.upper() in full_card_text.upper():
                    course_name = clean_name
                    break
            
            unique_display_name = format_assignment_name(course_name, assignment_title)

            current_card.click()
            iframe.locator("[data-test=\"back-button\"]").first.wait_for(state="visible", timeout=15000)

            assignment_title = extract_assignment_title(iframe, assignment_title)
            unique_display_name = format_assignment_name(course_name, assignment_title)
            legacy_display_name = f"[{course_name}] {assignment_title}"
            
            due_date = "No date specified"
            try:
                date_element = iframe.get_by_text(re.compile(r"^Due ")).first
                if date_element.is_visible(timeout=3000):
                    raw_due_text = date_element.inner_text()
                    due_date = clean_date_string(raw_due_text)
            except:
                pass

            previous_due = history.get(unique_display_name)
            if previous_due is None and legacy_display_name in history:
                previous_due = history.get(legacy_display_name)

            if previous_due == due_date:
                log_output(f"Unchanged: {unique_display_name} | Due: {due_date}")
                master_deadlines[unique_display_name] = due_date
                try:
                    iframe.locator("[data-test=\"back-button\"]").first.click()
                except:
                    page.go_back()
                iframe.locator(target_card_id).first.wait_for(state="visible", timeout=15000)
                continue

            assignments_to_process.append({
                "index": i,
                "title": assignment_title,
                "course": course_name,
                "display_name": unique_display_name,
                "due_date": due_date
            })
            changed_deadlines[unique_display_name] = due_date

            try:
                iframe.locator("[data-test=\"back-button\"]").first.click()
            except:
                page.go_back()
            iframe.locator(target_card_id).first.wait_for(state="visible", timeout=15000)
        
        for assignment_info in assignments_to_process:
            log_output(f"--- Processing: {assignment_info['display_name']} ---")
            log_output(f"  Deadline: {assignment_info['due_date']}")
            
            iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame
            current_card = iframe.locator(target_card_id).nth(assignment_info['index'])
            
            current_card.click()
            iframe.locator("[data-test=\"back-button\"]").first.wait_for(state="visible", timeout=15000)

            try:
                attachment_menus = iframe.locator("[data-test=\"attachment-options-button\"]")
                attachment_count = attachment_menus.count()

                if attachment_count > 0:
                    for j in range(attachment_count):
                        attachment_menus.nth(j).click()
                        
                        try:
                            download_btn = iframe.get_by_role("menuitem", name="Download")
                            download_btn.wait_for(state="visible", timeout=5000)
                        except:
                            download_btn = iframe.get_by_text("Download").first
                            download_btn.wait_for(state="visible", timeout=5000)

                        with page.expect_download() as download_info:
                            download_btn.click()
                        
                        download = download_info.value
                        original_filename = download.suggested_filename
                        
                        safe_course = re.sub(r'[\\/*?:"<>|]', "", assignment_info['course'])
                        safe_title = re.sub(r'[\\/*?:"<>|]', "", assignment_info['title'])
                        
                        new_filename = f"{safe_course} - {safe_title} - {original_filename}"
                        final_path = os.path.join(DOWNLOAD_DIR, new_filename)

                        if os.path.exists(final_path):
                            log_output(f"    => Skipping: '{new_filename}' already exists.")
                            download.cancel() 
                        else:
                            log_output(f"    => Saving as: '{new_filename}'")
                            download.save_as(final_path)

                        page.wait_for_timeout(1000)
                else:
                    log_output("  -> No files to download.")

            except Exception as e:
                log_output(f"  -> Error: {e}")

            try:
                iframe.locator("[data-test=\"back-button\"]").first.click()
            except:
                page.go_back()

            master_deadlines[assignment_info['display_name']] = assignment_info['due_date']
            history[assignment_info['display_name']] = assignment_info['due_date']
            
            iframe.locator(target_card_id).first.wait_for(state="visible", timeout=15000)

        context.close()

    log_output("Generating output files...")
    
    sorted_deadlines = dict(sorted(master_deadlines.items()))

    formatted_list = "ASSIGNMENT TRACKER\n"
    formatted_list += "------------------\n"
    
    for unique_name, date in sorted_deadlines.items():
        formatted_list += f"{unique_name}\nDue: {date}\n\n"
        
    with open(DEADLINE_FILE, "w", encoding="utf-8") as f:
        f.write(formatted_list)

    save_history(history)
        
    pyperclip.copy(formatted_list)
    
    try:
        subprocess.Popen([
            "explorer.exe",
            "shell:appsFolder\\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App"
        ])
    except Exception as e:
        log_output(f"Could not open Sticky Notes: {e}", show_popup=False)

    time.sleep(3)
    pyautogui.hotkey('ctrl', 'a')  # Select all
    time.sleep(0.3)
    pyautogui.press('delete')      # Delete selected text
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'v')  # Paste from clipboard

    if changed_deadlines:
        sync_result = sync_deadlines_to_google_calendar(changed_deadlines, interactive_auth=True)
        if sync_result["ok"]:
            show_windows_popup(
                "Classflow Calendar Sync",
                "Google Calendar API sync completed.\n\n"
                f"Inserted: {sync_result['inserted']}\n"
                f"Updated: {sync_result['updated']}\n"
                f"Skipped (invalid dates): {sync_result['skipped']}"
            )
        else:
            show_windows_popup(
                "Classflow Calendar Sync",
                "Google Calendar API sync failed for some items.\n\n"
                f"Inserted: {sync_result['inserted']}\n"
                f"Updated: {sync_result['updated']}\n"
                f"Failed: {sync_result['failed']}"
            )
    else:
        show_windows_popup(
            "Classflow Calendar Sync",
            "No changed deadlines were detected, so no Google Calendar API sync was needed."
        )
    
    log_output("Deadlines were copied to clipboard and pasted into Sticky Notes.", show_popup=IS_FROZEN, title="Classflow")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Classflow assignment automation")
    parser.add_argument(
        "--install-schedule",
        action="store_true",
        help="Create Windows Task Scheduler jobs at 12:00 and 18:00 daily.",
    )
    parser.add_argument(
        "--remove-schedule",
        action="store_true",
        help="Remove Windows Task Scheduler jobs created by Classflow.",
    )
    args = parser.parse_args()

    if args.install_schedule and args.remove_schedule:
        log_output("Use only one option at a time: --install-schedule or --remove-schedule", show_popup=False)
        sys.exit(1)

    if args.install_schedule:
        try:
            install_schedule()
            log_output("Classflow schedule installed for 12:00 and 18:00 daily.", show_popup=True, title="Classflow Setup")
        except Exception as e:
            log_output(f"Failed to install schedule: {e}", show_popup=True, title="Classflow Setup Error")
            sys.exit(1)
        sys.exit(0)

    if args.remove_schedule:
        try:
            remove_schedule()
            log_output("Classflow schedule removed.", show_popup=True, title="Classflow Setup")
        except Exception as e:
            log_output(f"Failed to remove schedule: {e}", show_popup=True, title="Classflow Setup Error")
            sys.exit(1)
        sys.exit(0)

    try:
        run()
    except KeyboardInterrupt:
        log_output("Run interrupted by user.", show_popup=False)