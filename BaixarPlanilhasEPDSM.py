from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    # Usa o Google Chrome instalado na máquina
    browser = p.chromium.launch(channel="msedge", headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://google.com")
    page.get_by_role("button", name="Aceitar tudo").click()
    page.get_by_role("combobox", name="Pesquisar").click()
    context.close()
    browser.close()
