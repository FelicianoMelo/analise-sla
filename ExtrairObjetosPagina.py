import json
from playwright.sync_api import sync_playwright


def extrair_dados_web(url: str):
    # headless=False abre o navegador visivelmente
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        print(f"Acessando: {url}")
        page.goto(url)
        page.wait_for_load_state("networkidle")  # Aguarda a página carregar completamente

        # 1. Extrair Campos de Entrada
        campos = []
        for el in page.locator("input, textarea, select").all():
            campos.append(
                {
                    "tag": el.evaluate("node => node.tagName.toLowerCase()"),
                    "tipo": el.get_attribute("type") or "",
                    "nome": el.get_attribute("name") or "",
                    "id": el.get_attribute("id") or "",
                    "placeholder": el.get_attribute("placeholder") or "",
                }
            )

        # 2. Extrair Botões
        botoes = []
        for el in page.locator(
            "button, input[type='button'], input[type='submit']"
        ).all():
            botoes.append(
                {
                    "texto": el.inner_text().strip(),
                    "id": el.get_attribute("id") or "",
                    "tipo": el.get_attribute("type") or "button",
                }
            )

        # 3. Extrair Links
        links = []
        for el in page.locator("a[href]").all():
            links.append(
                {
                    "texto": el.inner_text().strip(),
                    "url": el.get_attribute("href") or "",
                }
            )

        # Agrupar e salvar os dados extraídos
        dados_finais = {"campos": campos, "botoes": botoes, "links": links}

        with open("dados_extraidos.json", "w", encoding="utf-8") as f:
            json.dump(dados_finais, f, indent=4, ensure_ascii=False)

        print("\n--- Extração Concluída ---")
        print(f"Campos encontrados: {len(campos)}")
        print(f"Botões encontrados: {len(botoes)}")
        print(f"Links encontrados: {len(links)}")
        print("Dados salvos com sucesso em 'dados_extraidos.json'")

        browser.close()


# Execução do script
if __name__ == "__main__":
    # Substitua pela URL que você deseja raspar
    url_alvo = "http://redmine.stefanini.com.br/redmine/"
    extrair_dados_web(url_alvo)
