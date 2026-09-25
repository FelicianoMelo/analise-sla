import time
import pyautogui as pg
pg.PAUSE = 2.0  # Define uma pausa de 2 segundos entre as ações do PyAutoGUI

pg.hotkey('win', 'r')  # Abre a janela "Executar"
time.sleep(1)  # Aguarda 1 segundo para a janela abrir
pg.write('msedge')  # Digita "msedge" na janela "Executar"
pg.press('enter')  # Pressiona Enter para abrir o Microsoft Edge
pg.sleep(3)  # Aguarda 3 segundos para o navegador abrir completamente
pg.write('https://pdsm.service-now.com/now/platform-analytics-workspace/pages/params/library/data-visualizations/filter/created_by_me')  # Digita o endereço do site
pg.press('enter')  # Pressiona Enter para acessar o site    
time.sleep(4)  # Aguarda 4 segundos para o site carregar
x, y=pg.locateOnScreen('C:/Users/3868136/projetos/analise-sla/img/btn_rel_01.png', confidence=0)  # Localiza a imagem do botão de login na tela 
pg.moveTo(x, y)  # Move o cursor do mouse para a posição do botão de login
pg.click(x, y)  # Clica no link do relatorio
  


