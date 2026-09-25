import re
from playwright.sync_api import Playwright, sync_playwright, expect
import time

def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(channel="msedge", headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://purpleid.okta.com/oauth2/v1/authorize?client_id=okta.2b1959c8-bcc0-56eb-a589-cfcfb7422f26&code_challenge=p_zgnLFDzH8_it6dImfO6Qb-2feEJVwZ0Y86KrIytm4&code_challenge_method=S256&nonce=3Vo3KxFeMTPMgsHc5OzKSiDqwkcNpTTkO4z3QjsOv62bBRogxa5qajuzkCgMSBz8&redirect_uri=https%3A%2F%2Fpurpleid.okta.com%2Fenduser%2Fcallback&response_type=code&state=FIKAynTLEgx1CiX47NCLis3o0OfCVt12N9XkWa8AW2xzEDADIJMVFeWhbRwFPtRH&scope=openid%20profile%20email%20okta.users.read.self%20okta.users.manage.self%20okta.internal.enduser.read%20okta.internal.enduser.manage%20okta.enduser.dashboard.read%20okta.enduser.dashboard.manage%20okta.myAccount.sessions.manage%20okta.internal.navigation.enduser.read")
    page.get_by_role("textbox", name="Username").fill("3868136")
    page.get_by_role("textbox", name="Username").press("Enter")
    #
    page.get_by_role("textbox", name="Password").fill("2436Rywe#961")
    page.get_by_role("textbox", name="Password").press("Enter")
    #
    page.get_by_role("button", name="Verify").click()
    #
    page.get_by_role("link", name="Select to get a push").click()
    page.goto("https://purpleid.okta.com/app/UserHome")
    with page.expect_popup() as page1_info:
        page.get_by_role("link", name="launch app PDSM").first.click()
    page1 = page1_info.value
    page.close()
    page1.close()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
