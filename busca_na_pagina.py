from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    # Substitua pelo caminho correto se não for o padrão
    caminho_edge = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

    # Abre o Edge local e inicia o modo de gravação do Codegen
    browser = p.chromium.launch(executable_path=caminho_edge, headless=False)
    context = browser.new_context()

    # Isto abre o recetor de gravação (Codegen) dentro do script
    page = context.new_page()
    page.pause() 
    
    
    # get_by_role("link", name="FedEx PurpleID Login If you")
    
    # get_by_role("textbox", name="Username")
    # get_by_role("button", name="Verify")
    # get_by_role("textbox", name="Password")
    # get_by_role("button", name="Verify")
    
    # get_by_role("link", name="Select to get a push")