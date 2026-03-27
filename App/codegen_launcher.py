import os
from playwright.sync_api import sync_playwright

PROFILE_DIR = os.path.join(os.getcwd(), "automation_proxy")

def launch_inspector():
    with sync_playwright() as p:
        # Boot up using your exact saved session
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            channel="chrome",
            headless=False,
            accept_downloads=True
        )
        
        page = context.pages[0]
        page.goto("https://teams.microsoft.com/v2/")
        
        print("Opening Playwright Inspector...")
        
        # --- THE MAGIC LINE ---
        # This freezes the browser and opens the Codegen tool!
        page.pause() 

if __name__ == "__main__":
    launch_inspector()