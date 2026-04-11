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
import atexit
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

DEADLINE_FILE = os.path.join(PROFILE_DIR, "deadlines.txt")
HISTORY_FILE = os.path.join(PROFILE_DIR, "assignment_history.json")
GOOGLE_SETUP_FILE = os.path.join(PROFILE_DIR, "google_calendar_setup.flag")
TEAMS_SETUP_FILE = os.path.join(PROFILE_DIR, "teams_setup.flag")
CLASSFLOW_INTRO_FILE = os.path.join(PROFILE_DIR, "classflow_intro_shown.flag")
GOOGLE_TOKEN_FILE = os.path.join(PROFILE_DIR, "google_token.json")
SETTINGS_FILE = os.path.join(PROFILE_DIR, "classflow_settings.json")
GOOGLE_SCOPES = ["https://www.googleapis.com/auth/calendar.events"]
GOOGLE_CLIENT_SECRET_CANDIDATES = [
    "google_client_secret.json",
    "credentials.json",
    "client_secret.json"
]
DEFAULT_TASK_NAME = "Classflow"
DEFAULT_SETTINGS = {
    "download_dir": "",
    "sticky_notes_enabled": True,
    "calendar_sync_enabled": True,
    "calendar_id": "primary",
    "scheduler_enabled": False,
    "scheduler_time": "12:00",
    "scheduler_task_name": DEFAULT_TASK_NAME,
}
WINDOWS_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0

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

def load_settings():
    """Load user settings and merge with defaults."""
    merged = dict(DEFAULT_SETTINGS)
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                merged.update(loaded)
        except Exception:
            pass

    merged["download_dir"] = (merged.get("download_dir") or "").strip()
    merged["calendar_id"] = (merged.get("calendar_id") or "primary").strip() or "primary"
    merged["scheduler_task_name"] = (merged.get("scheduler_task_name") or DEFAULT_TASK_NAME).strip() or DEFAULT_TASK_NAME
    merged["scheduler_time"] = (merged.get("scheduler_time") or "12:00").strip()
    merged["sticky_notes_enabled"] = bool(merged.get("sticky_notes_enabled", True))
    merged["calendar_sync_enabled"] = bool(merged.get("calendar_sync_enabled", True))
    merged["scheduler_enabled"] = bool(merged.get("scheduler_enabled", False))
    return merged

def save_settings(settings):
    """Persist runtime settings to user profile."""
    os.makedirs(PROFILE_DIR, exist_ok=True)
    payload = dict(DEFAULT_SETTINGS)
    payload.update(settings or {})
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)

def normalize_directory_path(path_value):
    """Normalize and expand a path value from config or CLI."""
    return os.path.abspath(os.path.expandvars(os.path.expanduser(path_value.strip())))

def is_valid_time_format(hhmm):
    """Return True when time uses 24h HH:MM format."""
    try:
        datetime.strptime(hhmm, "%H:%M")
        return True
    except Exception:
        return False

def select_download_directory():
    """Show native Windows folder picker for choosing assignment download path."""
    try:
        if os.name != "nt":
            return ""

        ps_script = (
            "Add-Type -AssemblyName System.Windows.Forms;"
            "$dialog = New-Object System.Windows.Forms.FolderBrowserDialog;"
            "$dialog.Description = 'Choose folder for Classflow assignment downloads';"
            "$dialog.ShowNewFolderButton = $true;"
            "$result = $dialog.ShowDialog();"
            "if ($result -eq [System.Windows.Forms.DialogResult]::OK) {"
            "  [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;"
            "  Write-Output $dialog.SelectedPath"
            "}"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-Command", ps_script],
            capture_output=True,
            text=True,
            check=False,
            creationflags=WINDOWS_NO_WINDOW,
        )
        selected = (result.stdout or "").strip()
        return normalize_directory_path(selected) if selected else ""
    except Exception:
        return ""

def ensure_download_directory_configured(settings):
    """Ensure required download directory is available before runtime."""
    chosen = (settings.get("download_dir") or "").strip()
    if chosen:
        normalized = normalize_directory_path(chosen)
        os.makedirs(normalized, exist_ok=True)
        settings["download_dir"] = normalized
        return normalized

    show_windows_popup(
        "Classflow Setup",
        "Please choose where assignments should be downloaded."
    )
    selected = select_download_directory()
    if not selected:
        show_windows_popup(
            "Classflow Setup",
            "A download folder is required. Run setup again and choose a folder."
        )
        return ""

    os.makedirs(selected, exist_ok=True)
    settings["download_dir"] = selected
    save_settings(settings)
    return selected

def apply_setup_preferences(
    download_dir=None,
    sticky_notes_enabled=None,
    calendar_sync_enabled=None,
    calendar_id=None,
    scheduler_enabled=None,
    scheduler_time=None,
    scheduler_task_name=None,
):
    """Apply first-run preferences, typically passed by MSI/WPF setup."""
    settings = load_settings()

    if download_dir is not None:
        normalized = normalize_directory_path(download_dir)
        os.makedirs(normalized, exist_ok=True)
        settings["download_dir"] = normalized

    if sticky_notes_enabled is not None:
        settings["sticky_notes_enabled"] = bool(sticky_notes_enabled)

    if calendar_sync_enabled is not None:
        settings["calendar_sync_enabled"] = bool(calendar_sync_enabled)

    if calendar_id is not None:
        settings["calendar_id"] = (calendar_id or "primary").strip() or "primary"

    if scheduler_time is not None:
        if not is_valid_time_format(scheduler_time):
            raise ValueError("Scheduler time must be HH:MM in 24-hour format.")
        settings["scheduler_time"] = scheduler_time

    if scheduler_task_name is not None:
        settings["scheduler_task_name"] = (scheduler_task_name or DEFAULT_TASK_NAME).strip() or DEFAULT_TASK_NAME

    if scheduler_enabled is not None:
        settings["scheduler_enabled"] = bool(scheduler_enabled)

    save_settings(settings)

    if settings.get("scheduler_enabled"):
        create_windows_task(settings["scheduler_task_name"], settings["scheduler_time"])

    return settings

def show_first_time_setup_dialog():
    """Show interactive first-time setup dialog with checkboxes and return selected values."""
    if os.name != "nt":
        return None

    ps_script = r'''
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Classflow First-Time Setup"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(620, 390)
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.TopMost = $true

$lblDownload = New-Object System.Windows.Forms.Label
$lblDownload.Text = "Download folder for assignments"
$lblDownload.Location = New-Object System.Drawing.Point(20, 20)
$lblDownload.Size = New-Object System.Drawing.Size(280, 20)
$form.Controls.Add($lblDownload)

$txtDownload = New-Object System.Windows.Forms.TextBox
$txtDownload.Location = New-Object System.Drawing.Point(20, 45)
$txtDownload.Size = New-Object System.Drawing.Size(470, 24)
$form.Controls.Add($txtDownload)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse"
$btnBrowse.Location = New-Object System.Drawing.Point(500, 43)
$btnBrowse.Size = New-Object System.Drawing.Size(90, 28)
$btnBrowse.Add_Click({
    $folder = New-Object System.Windows.Forms.FolderBrowserDialog
    $folder.Description = "Choose folder for Classflow assignment downloads"
    $folder.ShowNewFolderButton = $true
    if ($folder.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtDownload.Text = $folder.SelectedPath
    }
})
$form.Controls.Add($btnBrowse)

$chkSticky = New-Object System.Windows.Forms.CheckBox
$chkSticky.Text = "Enable Sticky Notes sync"
$chkSticky.Location = New-Object System.Drawing.Point(20, 90)
$chkSticky.Size = New-Object System.Drawing.Size(260, 24)
$chkSticky.Checked = $true
$form.Controls.Add($chkSticky)

$chkCalendar = New-Object System.Windows.Forms.CheckBox
$chkCalendar.Text = "Enable timetable calendar sync"
$chkCalendar.Location = New-Object System.Drawing.Point(20, 125)
$chkCalendar.Size = New-Object System.Drawing.Size(280, 24)
$chkCalendar.Checked = $true
$form.Controls.Add($chkCalendar)

$lblCalendar = New-Object System.Windows.Forms.Label
$lblCalendar.Text = "Calendar ID"
$lblCalendar.Location = New-Object System.Drawing.Point(40, 155)
$lblCalendar.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($lblCalendar)

$txtCalendar = New-Object System.Windows.Forms.TextBox
$txtCalendar.Location = New-Object System.Drawing.Point(160, 152)
$txtCalendar.Size = New-Object System.Drawing.Size(300, 24)
$txtCalendar.Text = "primary"
$form.Controls.Add($txtCalendar)

$chkScheduler = New-Object System.Windows.Forms.CheckBox
$chkScheduler.Text = "Create Windows scheduled task"
$chkScheduler.Location = New-Object System.Drawing.Point(20, 195)
$chkScheduler.Size = New-Object System.Drawing.Size(280, 24)
$chkScheduler.Checked = $false
$form.Controls.Add($chkScheduler)

$lblTime = New-Object System.Windows.Forms.Label
$lblTime.Text = "Task time (24h)"
$lblTime.Location = New-Object System.Drawing.Point(40, 225)
$lblTime.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($lblTime)

$timePicker = New-Object System.Windows.Forms.DateTimePicker
$timePicker.Location = New-Object System.Drawing.Point(160, 222)
$timePicker.Size = New-Object System.Drawing.Size(120, 24)
$timePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$timePicker.CustomFormat = "HH:mm"
$timePicker.ShowUpDown = $true
$timePicker.Value = [datetime]::Today.AddHours(12)
$timePicker.Enabled = $false
$form.Controls.Add($timePicker)

$lblTask = New-Object System.Windows.Forms.Label
$lblTask.Text = "Task name"
$lblTask.Location = New-Object System.Drawing.Point(40, 255)
$lblTask.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($lblTask)

$txtTaskName = New-Object System.Windows.Forms.TextBox
$txtTaskName.Location = New-Object System.Drawing.Point(160, 252)
$txtTaskName.Size = New-Object System.Drawing.Size(300, 24)
$txtTaskName.Text = "Classflow Daily"
$txtTaskName.Enabled = $false
$form.Controls.Add($txtTaskName)

$chkCalendar.Add_CheckedChanged({
    $txtCalendar.Enabled = $chkCalendar.Checked
})

$chkScheduler.Add_CheckedChanged({
    $timePicker.Enabled = $chkScheduler.Checked
    $txtTaskName.Enabled = $chkScheduler.Checked
})

$btnOk = New-Object System.Windows.Forms.Button
$btnOk.Text = "Save Setup"
$btnOk.Location = New-Object System.Drawing.Point(380, 305)
$btnOk.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnOk)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(490, 305)
$btnCancel.Size = New-Object System.Drawing.Size(100, 30)
$btnCancel.Add_Click({
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Close()
})
$form.Controls.Add($btnCancel)

$script:SetupJson = $null
$btnOk.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtDownload.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please choose a download folder.", "Classflow Setup") | Out-Null
        return
    }

    if ($chkCalendar.Checked -and [string]::IsNullOrWhiteSpace($txtCalendar.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter Calendar ID or leave timetable sync unchecked.", "Classflow Setup") | Out-Null
        return
    }

    $calendarId = "primary"
    if ($chkCalendar.Checked) {
        $calendarId = $txtCalendar.Text.Trim()
    }

    $taskName = $txtTaskName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($taskName)) {
        $taskName = "Classflow Daily"
    }

    $payload = @{
        download_dir = $txtDownload.Text.Trim()
        sticky_notes_enabled = $chkSticky.Checked
        calendar_sync_enabled = $chkCalendar.Checked
        calendar_id = $calendarId
        scheduler_enabled = $chkScheduler.Checked
        scheduler_time = $timePicker.Value.ToString("HH:mm")
        scheduler_task_name = $taskName
    }

    $script:SetupJson = $payload | ConvertTo-Json -Compress
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

$form.AcceptButton = $btnOk
$form.CancelButton = $btnCancel

$result = $form.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $script:SetupJson) {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    Write-Output $script:SetupJson
}
'''

    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-Command", ps_script],
            capture_output=True,
            text=True,
            check=False,
            creationflags=WINDOWS_NO_WINDOW,
        )
        setup_json = (result.stdout or "").strip()
        if not setup_json:
            return None
        return json.loads(setup_json)
    except Exception:
        return None

def ensure_first_time_setup_completed():
    """Run interactive setup wizard once and persist settings."""
    if os.path.exists(SETTINGS_FILE):
        return load_settings()

    show_windows_popup(
        "Welcome to Classflow",
        "Welcome to Classflow!\n\n"
        "Press OK to start your first-time setup."
    )

    payload = show_first_time_setup_dialog()
    if not payload:
        log_output("First-time setup cancelled by user.", show_popup=False)
        return None

    try:
        apply_setup_preferences(
            download_dir=payload.get("download_dir"),
            sticky_notes_enabled=payload.get("sticky_notes_enabled"),
            calendar_sync_enabled=payload.get("calendar_sync_enabled"),
            calendar_id=payload.get("calendar_id"),
            scheduler_enabled=payload.get("scheduler_enabled"),
            scheduler_time=payload.get("scheduler_time"),
            scheduler_task_name=payload.get("scheduler_task_name"),
        )
        return load_settings()
    except Exception as exc:
        show_windows_popup("Classflow Setup Error", f"Setup failed: {exc}")
        return None

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

def is_classflow_intro_shown():
    """Check whether the intro dialogue has been shown to the user."""
    return os.path.exists(CLASSFLOW_INTRO_FILE)

def mark_classflow_intro_shown():
    """Persist intro dialogue shown marker."""
    os.makedirs(PROFILE_DIR, exist_ok=True)
    with open(CLASSFLOW_INTRO_FILE, "w", encoding="utf-8") as f:
        f.write(datetime.now().isoformat())

def clean_date_string(raw_text):
    """Just strips the word 'Due ' from the text."""
    return raw_text.replace("Due ", "").strip()

def parse_due_date(raw_due_date):
    """Parse Teams due date text into a datetime when possible."""
    if not raw_due_date or raw_due_date == "No date specified":
        return None

    cleaned = raw_due_date.strip()
    cleaned = re.sub(r"\s+", " ", cleaned)

    # Handle relative Teams strings like "Today at 11:59 PM" and "Tomorrow at 11:59 PM".
    relative_match = re.match(
        r"^(today|tomorrow)(?:\s+at)?(?:\s+(\d{1,2}:\d{2}\s*[ap]m))?$",
        cleaned,
        flags=re.IGNORECASE,
    )
    if relative_match:
        day_word = relative_match.group(1).lower()
        time_part = relative_match.group(2)

        base_date = datetime.now().date()
        if day_word == "tomorrow":
            base_date += timedelta(days=1)

        if time_part:
            try:
                parsed_time = datetime.strptime(time_part.strip().upper(), "%I:%M %p").time()
            except ValueError:
                return None
        else:
            parsed_time = datetime.min.time()

        return datetime.combine(base_date, parsed_time)

    cleaned = cleaned.replace(" at ", " ")

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
EXE_LOG_FILE = os.path.join(PROFILE_DIR, "classflow_log.txt") if IS_FROZEN else None
RUNTIME_LOG_FILE = os.path.join(PROFILE_DIR, "classflow_runtime.log")
LOGGER_SCRIPT_FILE = os.path.join(PROFILE_DIR, "classflow_logger_viewer.ps1")
_RUNTIME_LOG_HANDLE = None

class TeeRuntimeStream:
    """Mirror stdout/stderr writes to a runtime log file."""
    def __init__(self, original_stream, mirror_handle):
        self.original_stream = original_stream
        self.mirror_handle = mirror_handle
        self.encoding = getattr(original_stream, "encoding", "utf-8")

    def write(self, data):
        if not data:
            return 0
        if self.original_stream:
            try:
                self.original_stream.write(data)
            except Exception:
                pass
        try:
            self.mirror_handle.write(data)
            self.mirror_handle.flush()
        except Exception:
            pass
        return len(data)

    def flush(self):
        if self.original_stream:
            try:
                self.original_stream.flush()
            except Exception:
                pass
        try:
            self.mirror_handle.flush()
        except Exception:
            pass

def _write_logger_viewer_script(script_path):
    """Create a native WPF log viewer script used by frozen EXE mode."""
    ps_script = r'''param([Parameter(Mandatory=$true)][string]$LogPath)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Classflow Logger"
        WindowStartupLocation="CenterScreen"
        SizeToContent="Manual"
        Width="760"
        Height="420"
        Topmost="True"
        ResizeMode="CanResizeWithGrip"
        Background="#FFF0F0F0">
    <Grid Margin="10">
        <TextBox Name="LogBox"
                 FontFamily="Consolas"
                 FontSize="12"
                 IsReadOnly="True"
                 AcceptsReturn="True"
                 TextWrapping="NoWrap"
                 VerticalScrollBarVisibility="Auto"
                 HorizontalScrollBarVisibility="Auto"/>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)
$logBox = $window.FindName("LogBox")

function Update-LogBox {
    if (Test-Path $LogPath) {
        try {
            $lines = Get-Content -Path $LogPath -Tail 300 -ErrorAction Stop
            $text = [string]::Join("`r`n", $lines)
            if ($logBox.Text -ne $text) {
                $logBox.Text = $text
                $logBox.ScrollToEnd()
            }
        } catch {
        }
    }
}

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(700)
$timer.Add_Tick({ Update-LogBox })
$timer.Start()

Update-LogBox
[void]$window.ShowDialog()
'''
    with open(script_path, "w", encoding="utf-8") as f:
        f.write(ps_script)

def _start_wpf_runtime_logger(log_path):
    """Disabled: runtime logs are now file-only and no external log window is launched."""
    return

def initialize_runtime_logging():
    """Route print output to file in EXE mode without opening an external log window."""
    global _RUNTIME_LOG_HANDLE
    if _RUNTIME_LOG_HANDLE is not None:
        return
    try:
        os.makedirs(PROFILE_DIR, exist_ok=True)
        _RUNTIME_LOG_HANDLE = open(RUNTIME_LOG_FILE, "w", encoding="utf-8", buffering=1)
        _RUNTIME_LOG_HANDLE.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Classflow run started\n")
        sys.stdout = TeeRuntimeStream(sys.stdout, _RUNTIME_LOG_HANDLE)
        sys.stderr = TeeRuntimeStream(sys.stderr, _RUNTIME_LOG_HANDLE)
    except Exception:
        _RUNTIME_LOG_HANDLE = None

def close_runtime_logging():
    """Close runtime log file handle on process shutdown."""
    global _RUNTIME_LOG_HANDLE
    if _RUNTIME_LOG_HANDLE is not None:
        try:
            _RUNTIME_LOG_HANDLE.flush()
            _RUNTIME_LOG_HANDLE.close()
        except Exception:
            pass
        _RUNTIME_LOG_HANDLE = None

atexit.register(close_runtime_logging)

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

def remove_configured_schedule_task(settings):
    """Remove the configurable single schedule task if present."""
    task_name = (settings.get("scheduler_task_name") or "").strip()
    if task_name:
        delete_windows_task(task_name)

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

def get_google_calendar_service(interactive_auth, show_prompt_before_auth=False):
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
            if show_prompt_before_auth:
                show_windows_popup(
                    "Classflow Google Authentication",
                    "Press OK to open Google sign-in and complete authentication."
                )
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

    has_explicit_time = "am" in due_text.lower() or "pm" in due_text.lower()
    resolved_due_text = (
        parsed_due.strftime("%A, %B %d, %Y %I:%M %p")
        if has_explicit_time
        else parsed_due.strftime("%A, %B %d, %Y")
    )

    event = {
        "summary": unique_name,
        "description": f"Assignment deadline: {resolved_due_text}"
    }

    if parsed_due.hour == 0 and parsed_due.minute == 0 and not has_explicit_time:
        event["start"] = {"date": parsed_due.strftime("%Y-%m-%d")}
        event["end"] = {"date": (parsed_due + timedelta(days=1)).strftime("%Y-%m-%d")}
    else:
        local_tz = datetime.now().astimezone().tzinfo
        if parsed_due.tzinfo is None:
            parsed_due = parsed_due.replace(tzinfo=local_tz)

        due_utc = parsed_due.astimezone(timezone.utc)
        start_utc = (parsed_due - timedelta(minutes=60)).astimezone(timezone.utc)
        event["start"] = {"dateTime": start_utc.strftime("%Y-%m-%dT%H:%M:%SZ")}
        event["end"] = {"dateTime": due_utc.strftime("%Y-%m-%dT%H:%M:%SZ")}

    return event

def make_google_event_id(unique_name):
    """Build a deterministic Google Calendar event id for upsert behavior."""
    digest = hashlib.sha1(unique_name.encode("utf-8")).hexdigest()
    return f"cf{digest}"

def sync_deadlines_to_google_calendar(deadlines, interactive_auth, calendar_id):
    """Insert/update changed deadlines in Google Calendar through API."""
    if not deadlines:
        return {"ok": True, "inserted": 0, "updated": 0, "skipped": 0, "failed": 0}

    service = get_google_calendar_service(interactive_auth=interactive_auth)
    if not service:
        return {"ok": False, "inserted": 0, "updated": 0, "skipped": 0, "failed": len(deadlines)}

    calendar_id = (calendar_id or "primary").strip() or "primary"
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
    log_output("Opening Microsoft Teams...", show_popup=False, title="Classflow")
    page.set_default_navigation_timeout(60000)
    page.set_default_timeout(600000 if setup_mode else 60000)

    try:
        page.goto("https://teams.microsoft.com/v2/", timeout=60000)
    except Exception as nav_err:
        log_output(f"Teams navigation timeout, retrying once: {nav_err}", show_popup=False)
        page.goto("https://teams.microsoft.com/v2/", timeout=60000)

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

def show_classflow_info_dialogue(settings):
    """Show comprehensive info dialogue about what Classflow will do."""
    sticky_notes_enabled = settings.get("sticky_notes_enabled", True)
    calendar_sync_enabled = settings.get("calendar_sync_enabled", True)
    
    info_text = (
        "Classflow is ready to track your assignments!\n\n"
        "Here's what will happen:\n\n"
        "1. Classflow will open Microsoft Teams and guide first-time login \n"
        "2. If Calendar Sync is enabled, Google authentication will open once\n"
        "3. It will detect deadlines and download your assignments from Teams \n\n"
    )
    
    if sticky_notes_enabled:
        info_text += "• Your Sticky Notes will be updated with your assignment deadlines\n"
    
    if calendar_sync_enabled:
        info_text += "• Your Google Calendar will be synced with all assignment deadlines\n"
    
    if not sticky_notes_enabled and not calendar_sync_enabled:
        info_text += "• Assignment data will be extracted and saved only\n"
    
    info_text += (
        "\n4. At the end you will receive a summary of the sync results\n\n"
        "This process typically takes a few minutes."
    )
    
    show_windows_popup("About Classflow", info_text)

def first_time_setup():
    """Complete one-time Teams and Google setup, then ask user to rerun."""
    log_output("First-time setup detected. Starting account setup...", show_popup=False, title="Classflow Setup")

    settings = load_settings()
    if not is_classflow_intro_shown():
        show_classflow_info_dialogue(settings)
        mark_classflow_intro_shown()

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

        google_ok = True
        if settings.get("calendar_sync_enabled"):
            google_ok = is_google_setup_complete()
            if not google_ok:
                google_ok = get_google_calendar_service(
                    interactive_auth=True,
                    show_prompt_before_auth=False,
                ) is not None
                if google_ok:
                    mark_google_setup_complete()

        context.close()

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
    settings = ensure_first_time_setup_completed()
    if not settings:
        return

    download_dir = ensure_download_directory_configured(settings)
    if not download_dir:
        return

    if not is_teams_setup_complete() or (settings.get("calendar_sync_enabled") and not is_google_setup_complete()):
        first_time_setup()
        return
    
    # Show intro dialogue on first run after setup
    if not is_classflow_intro_shown():
        show_classflow_info_dialogue(settings)
        mark_classflow_intro_shown()
    
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
        
        log_output(f"Found {total_assignments} upcoming assignment(s).", show_popup=False, title="Classflow")
        
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
                        final_path = os.path.join(download_dir, new_filename)

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
    
    if settings.get("sticky_notes_enabled"):
        try:
            subprocess.Popen([
                "explorer.exe",
                "shell:appsFolder\\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App"
            ])
            time.sleep(3)
            pyautogui.hotkey('ctrl', 'a')  # Select all
            time.sleep(0.3)
            pyautogui.press('delete')      # Delete selected text
            time.sleep(0.3)
            pyautogui.hotkey('ctrl', 'v')  # Paste from clipboard
            log_output("Deadlines were copied to clipboard and pasted into Sticky Notes.", show_popup=False, title="Classflow")
        except Exception as e:
            log_output(f"Could not update Sticky Notes: {e}", show_popup=False)
    else:
        log_output("Sticky Notes sync disabled in setup settings.", show_popup=False)

    if settings.get("calendar_sync_enabled"):
        if changed_deadlines:
            sync_result = sync_deadlines_to_google_calendar(
                changed_deadlines,
                interactive_auth=True,
                calendar_id=settings.get("calendar_id", "primary"),
            )
            if sync_result["ok"]:
                log_output(
                    "Google Calendar API sync completed. "
                    f"Inserted: {sync_result['inserted']}, "
                    f"Updated: {sync_result['updated']}, "
                    f"Skipped: {sync_result['skipped']}",
                    show_popup=False,
                )
                show_windows_popup(
                    "Classflow Calendar Sync",
                    "Google Calendar API sync completed.\n\n"
                    f"Inserted: {sync_result['inserted']}\n"
                    f"Updated: {sync_result['updated']}\n"
                    f"Skipped: {sync_result['skipped']}"
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
            log_output("No changed deadlines were detected, so no Google Calendar API sync was needed.", show_popup=False)
            show_windows_popup(
                "Classflow Calendar Sync",
                "No changed deadlines were detected, so no Google Calendar API sync was needed."
            )
    else:
        log_output("Google Calendar sync disabled in setup settings.", show_popup=False)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Classflow assignment automation")
    args = parser.parse_args()

    initialize_runtime_logging()

    try:
        run()
    except KeyboardInterrupt:
        log_output("Run interrupted by user.", show_popup=False)