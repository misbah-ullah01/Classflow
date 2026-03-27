import os
from playwright.sync_api import sync_playwright

def run():
    with sync_playwright() as p:
        # Launch a CLEAN browser, not your main one
        browser = p.chromium.launch(headless=False)
        # Create a fresh context
        context = browser.new_context()
        page = context.new_page()

        print("1. Log in to Google Classroom manually.")
        print("2. Log in to Microsoft Teams manually.")
        print("3. Once you are fully in, come back here and press Enter.")
        
        page.goto("https://classroom.google.com")
        
        input("Press Enter here ONLY after you have logged into BOTH sites...")

        # This saves ONLY the cookies/session to a small file
        context.storage_state(path="state.json")
        print("Success! 'state.json' created. You can close the browser.")
        browser.close()

if __name__ == "__main__":
    run()