import os
import re
import json
import pyperclip
import subprocess
import time
import pyautogui
import ctypes
from playwright.sync_api import sync_playwright

# --- 1. DYNAMIC CONFIGURATION ---
# Safely hide the browser profile in the user's hidden AppData folder
PROFILE_DIR = os.path.join(os.environ['LOCALAPPDATA'], "Classflow")

# Dynamically locate the current user's Desktop
USER_DESKTOP = os.path.join(os.environ['USERPROFILE'], 'Desktop')

DOWNLOAD_DIR = os.path.join(USER_DESKTOP, "Assignments")
DEADLINE_FILE = os.path.join(USER_DESKTOP, "deadlines.txt")
HISTORY_FILE = os.path.join(USER_DESKTOP, "assignment_history.json")

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

def clean_date_string(raw_text):
    """Just strips the word 'Due ' from the text."""
    return raw_text.replace("Due ", "").strip()

def show_windows_popup(title, message):
    """Show a Windows popup message."""
    ctypes.windll.user32.MessageBoxW(0, message, title, 0)

def show_windows_yes_no_popup(title, message):
    """Show a Windows Yes/No popup. Returns True if Yes clicked, False if No."""
    result = ctypes.windll.user32.MessageBoxW(0, message, title, 4)  # 4 = Yes/No buttons
    return result == 6  # 6 = IDYES


def run():
    # --- CHECK IF FIRST TIME SETUP ---
    is_first_time = not os.path.exists(PROFILE_DIR)
    
    if is_first_time:
        print("\n" + "="*50)
        print("🔐 FIRST TIME SETUP DETECTED")
        print("="*50 + "\n")
        
        show_windows_popup(
            "Teams Bot Setup",
            "You need to log in to Microsoft Teams.\n\n"
            "A Chrome window will open.\n"
            "Please log in and click 'Yes' to stay signed in.\n\n"
            "Once login is saved, the window will close automatically."
        )
    
    master_deadlines = {}
    history = load_history()
    assignments_to_process = []
    
    # no GUI needed for background runs

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            channel="chrome",
            headless=not is_first_time,  # headless=False for first time, True for subsequent runs
            accept_downloads=True
        )
        page = context.pages[0]
        
        print("Navigating to Teams...")
        # raise default timeouts so slow loads don't immediately fail
        page.set_default_navigation_timeout(60000)
        page.set_default_timeout(60000)
        try:
            page.goto("https://teams.microsoft.com/v2/", timeout=60000)
        except Exception as nav_err:
            print(f"Navigation timeout, retrying: {nav_err}")
            page.goto("https://teams.microsoft.com/v2/", timeout=60000)
        
        # --- 2. FIRST-TIME LOGIN DETECTION ---
        print("Checking login status...")
        
        # configure default timeout based on first-time vs normal run
        page.set_default_timeout(600000 if is_first_time else 300000)

        # log instructions if login is needed
        if is_first_time:
            print("\n" + "="*50)
            print("🚨 LOGIN REQUIRED")
            print("Please complete login in the Chrome window.")
            print("Accept the 'Stay signed in' prompt when asked.")
            print("="*50 + "\n")
        elif "login" in page.url or "microsoftonline" in page.url:
            print("\n" + "="*50)
            print("⚠️  SESSION EXPIRED - PLEASE LOG IN AGAIN")
            print("="*50 + "\n")

        print("Waiting for Teams to load and Assignments button...")
        
        # auto‑wait for the button (Playwright will respect default timeout)
        try:
            assignments_btn = page.get_by_role("button", name="Assignments (Ctrl+Shift+4)")
            assignments_btn.click()
        except Exception:
            try:
                assignments_btn = page.locator("[data-tid='app-bar-edu-assignments']").first
                assignments_btn.click()
            except Exception:
                if is_first_time:
                    print("❌ Login timeout. Please try again.")
                    context.close()
                    return
                else:
                    raise
        
        # --- IF FIRST TIME LOGIN SUCCESSFUL, CLOSE AND SHOW POPUP ---
        if is_first_time:
            print("✅ Login successful! Assignments button found. Closing browser...\n")
            context.close()
            
            # --- SHOW SETUP COMPLETE MESSAGE ---
            show_windows_popup(
                "Teams Bot Setup Complete",
                "✅ Login saved successfully!\n\n"
                "The browser has closed.\n\n"
                "Please run this program again to start downloading assignments."
            )
            print("\n" + "="*50)
            print("✅ FIRST TIME SETUP COMPLETE")
            print("Please run the program again.")
            print("="*50 + "\n")
            return  # Exit early
        
        # --- CLICK ASSIGNMENTS BUTTON FOR SUBSEQUENT RUNS ---
        assignments_btn.click()

        print("Waiting for assignment list to fetch...")
        iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame
        target_card_id = "[id*='CardHeader__headerEDUASSIGN']:visible"
        iframe.locator(target_card_id).first.wait_for(state="visible", timeout=30000)

        assignment_cards = iframe.locator(target_card_id)
        total_assignments = assignment_cards.count()
        
        print(f"\nFound {total_assignments} clickable upcoming assignment(s)!\n")
        
        # Update progress bar with total assignments
        # no spinner updates
        
        # --- FIRST PASS: EXTRACT CARD INFO AND CHECK HISTORY ---
        for i in range(total_assignments):
            iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame
            current_card = iframe.locator(target_card_id).nth(i)
            
            # --- EXTRACT BASIC INFO FROM CARD TEXT ---
            full_card_text = current_card.locator("xpath=..").inner_text()
            assignment_title = full_card_text.split('\n')[0].strip()
            
            course_name = "Other" 
            for key, clean_name in COURSE_MAP.items():
                if key.upper() in full_card_text.upper():
                    course_name = clean_name
                    break
            
            unique_display_name = f"[{course_name}] {assignment_title}"
            
            # --- CLICK CARD TO EXTRACT DUE DATE PROPERLY ---
            current_card.click()
            iframe.locator("[data-test=\"back-button\"]").first.wait_for(state="visible", timeout=15000)
            
            due_date = "No date specified"
            try:
                date_element = iframe.get_by_text(re.compile(r"^Due ")).first
                if date_element.is_visible(timeout=3000):
                    raw_due_text = date_element.inner_text()
                    due_date = clean_date_string(raw_due_text)
            except:
                pass
            
            # --- CHECK IF DEADLINE HAS CHANGED ---
            if unique_display_name in history:
                if history[unique_display_name] == due_date:
                    print(f"⏭️  SKIPPED (unchanged): {unique_display_name}")
                    print(f"   Due: {due_date}")
                    master_deadlines[unique_display_name] = due_date
                    # Go back without processing
                    try:
                        iframe.locator("[data-test=\"back-button\"]").first.click()
                    except:
                        page.go_back()
                    iframe.locator(target_card_id).first.wait_for(state="visible", timeout=15000)
                    continue
            
            # --- ADD TO PROCESS LIST ---
            assignments_to_process.append({
                "index": i,
                "title": assignment_title,
                "course": course_name,
                "display_name": unique_display_name,
                "due_date": due_date
            })
            
            # Go back to list
            try:
                iframe.locator("[data-test=\"back-button\"]").first.click()
            except:
                page.go_back()
            iframe.locator(target_card_id).first.wait_for(state="visible", timeout=15000)
        
        # --- SECOND PASS: PROCESS CHANGED/NEW ASSIGNMENTS ---
        # no spinner updates
        
        for idx, assignment_info in enumerate(assignments_to_process):
            print(f"--- Processing: {assignment_info['display_name']} ---")
            print(f"  -> 📅 Deadline: {assignment_info['due_date']}")
            
            # no spinner update

            
            iframe = page.locator("iframe[name=\"embedded-page-container\"]").content_frame
            current_card = iframe.locator(target_card_id).nth(assignment_info['index'])
            
            current_card.click()
            iframe.locator("[data-test=\"back-button\"]").first.wait_for(state="visible", timeout=15000)

            try:
                # --- DOWNLOAD & RENAME FILES ---
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
                        
                        new_filename = f"[{safe_course}] {safe_title} - {original_filename}"
                        final_path = os.path.join(DOWNLOAD_DIR, new_filename)

                        if os.path.exists(final_path):
                            print(f"    => Skipping: '{new_filename}' already exists.")
                            download.cancel() 
                        else:
                            print(f"    => Saving as: '{new_filename}'")
                            download.save_as(final_path)
                        
                        # wait for download to finish (should be handled by expect_download)
                        page.wait_for_timeout(1000)
                else:
                    print("  -> No files to download.")

            except Exception as e:
                print(f"  -> Error: {e}")

            try:
                iframe.locator("[data-test=\"back-button\"]").first.click()
            except:
                page.go_back()
            
            # Update history and master deadlines
            master_deadlines[assignment_info['display_name']] = assignment_info['due_date']
            history[assignment_info['display_name']] = assignment_info['due_date']
            
            iframe.locator(target_card_id).first.wait_for(state="visible", timeout=15000)

        context.close()

    # --- 4. FORMATTED TEXT GENERATION ---
    print("\n=== GENERATING CUSTOM STICKY NOTE ===\n")
    
    sorted_deadlines = dict(sorted(master_deadlines.items()))

    formatted_list = "📅 ASSIGNMENT TRACKER 📅\n"
    formatted_list += "------------------------\n"
    
    for unique_name, date in sorted_deadlines.items():
        formatted_list += f"🔴 {unique_name}\n   Due: {date}\n\n"
        
    with open(DEADLINE_FILE, "w", encoding="utf-8") as f:
        f.write(formatted_list)
    
    # Save updated history
    save_history(history)
        
    pyperclip.copy(formatted_list)
    
    # --- 5. OPEN WINDOWS STICKY NOTES AND PASTE ---
    # no spinner update
    
    try:
        subprocess.Popen([
            "explorer.exe",
            "shell:appsFolder\\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App"
        ])
    except Exception as e:
        print(f"Could not open Sticky Notes: {e}")
    
    # Give the app time to launch
    time.sleep(3)
    
    # Send keyboard commands
    pyautogui.hotkey('ctrl', 'a')  # Select all
    time.sleep(0.3)
    pyautogui.press('delete')      # Delete selected text
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'v')  # Paste from clipboard
    
    print("📋 Formatted deadlines pasted to Sticky Notes!")
    
    # Close progress bar
    # nothing to close after completion

if __name__ == "__main__":
    run()