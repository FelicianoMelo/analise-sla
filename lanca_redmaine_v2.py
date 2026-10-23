import re
from playwright.sync_api import sync_playwright, Playwright, expect

with sync_playwright() as pw:
    navegador = pw.chromium.launch(headless=False)
    pagina = navegador.new_page()

    pagina.goto('http://redmine.stefanini.com.br/redmine/')

    # login no portal
    pagina.get_by_role("textbox", name="Usuário").fill("fajunior1")
    pagina.get_by_role("textbox", name="Senha Perdi minha senha").fill("12345678")
    pagina.get_by_role("button", name="Entrar").press("Enter")

    # seleciona projeto 
    pagina.get_by_text("Ir para o projeto...").click()
    pagina.get_by_role("link", name="TMS").first.click()

    pagina.get_by_role("link", name="Implementação").click()
    pagina.get_by_role("link", name="35053").click()
    pagina.get_by_role("link", name="Tempo de trabalho").first.click()

    navegador.close()
