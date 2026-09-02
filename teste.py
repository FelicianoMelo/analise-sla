import asyncio
from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            executable_path=r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            headless=False
        )
         
        page = await browser.new_page()
        # acessar o portal Global PDSM Enterprise Portal 
        await page.goto("https://pdsm.service-now.com/now/platform-analytics-workspace/pages/params/library/data-visualizations/filter/created_by_me")
        await page.get_by_role("link", name="FedEx PurpleID Login If you").click()
        await page.get_by_role("textbox", name="Username").fill("3868136")
        await page.get_by_role("button", name="Next").click()
        await page.get_by_role("textbox", name="Password").fill("2436Rywe#961")
        await page.get_by_role("button", name="Verify").click()
        await page.pause()
        await browser.close()

asyncio.run(main())