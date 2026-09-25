from playwright.sync_api import sync_playwright
import keyboard
import time
import pyautogui

with sync_playwright() as p:
        browser = p.chromium.launch(
                executable_path=r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                headless=False
            )
        page = browser.new_context().new_page()
        page.goto("https://pdsm.service-now.com/now/platform-analytics-workspace/pages/params/library/data-visualizations/filter/created_by_me")
        page.get_by_role("link", name="FedEx PurpleID Login If you").click()
        page.get_by_role("textbox", name="Username").fill("3868136")
        page.get_by_role("button", name="Next").click()
        page.get_by_role("textbox", name="Password").fill("2436Rywe#961")
        page.get_by_role("button", name="Verify").click()
        
        page.get_by_role("link", name="Select to get a push").click()
        cliques = []
        espacos = 0
        
        print("Pressione ENTER para registrar um clique.")
        print("Pressione ESPAÇO duas vezes para finalizar.")
       
        while True:
            if keyboard.is_pressed("enter"):
                pos_mouse = pyautogui.position()
                cliques.append((pos_mouse.x, pos_mouse.y))
                print(f"Clique registrado na posição: ({pos_mouse.x}, {pos_mouse.y})")
                time.sleep(0.5)  # Pequena pausa para evitar múltiplos registros
                
            if keyboard.is_pressed("space"):
                espacos += 1
                print(f"Espaço pressionado ({espacos}/2)")
                time.sleep(0.5)  # Pequena pausa para evitar múltiplos registros
                
                if espacos >= 2:
                    print("\nRegistro de cliques finalizado!")
                    break
            
        browser.close()
                