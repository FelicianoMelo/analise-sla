from playwright.sync_api import sync_playwright

with sync_playwright() as p:
        browser = p.chromium.launch(
                executable_path=r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                headless=False
            )
             
        page = browser.new_page()
        page.goto("https://pdsm.service-now.com/now/platform-analytics-workspace/pages/params/library/data-visualizations/filter/created_by_me")
        page.get_by_role("link", name="FedEx PurpleID Login If you").click()
        page.get_by_role("textbox", name="Username").fill("3868136")
        page.get_by_role("button", name="Next").click()
        page.get_by_role("textbox", name="Password").fill("2436Rywe#961")
        page.get_by_role("button", name="Verify").click()
        browser.close()