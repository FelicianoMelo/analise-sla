from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(
        user_data_dir=r"C:\Users\3868136\AppData\Local\Microsoft\Edge\User Data",
        channel="msedge",
        headless=False
    )

    page = context.new_page()
    page.goto("https://purpleid.okta.com/app/UserHome")