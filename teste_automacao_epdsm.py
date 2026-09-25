import pyautogui as pa   # Automação do mouse, teclado e tela
import time                     # Controle de tempo e pausas no código
import pandas as pd      # Manipulação e análise de dados em tabelas
import os                        # Interação com arquivos e pastas do sistema
import glob                     # Busca de arquivos usando padrões (*.csv, *.xlsx)

pa.PAUSE = 0.5

def abrir_navegador():
    # Apertar tecla Windows para abrir o menu Iniciar
    pa.press('win')
    
    # Digitar "Edge" na barra de pesquisa
    pa.write('edge')
    
    # Apertar Enter para abrir o Edge
    pa.press('enter')
    
    # Esperar 3 segundos para garantir que o navegador abriu completamente
    time.sleep(3)
    
def abrir_scihub():
    # Aperta a combinação de teclas CTRL + L para selecionar barra de navegação
    pa.hotkey('ctrl', 'l')
    
    # Digitar o endereço do site
    pa.write('https://sci-hub.in')
    
    # Apertar Enter
    pa.press('enter')
    
    # Esperar 4 segundos para o site carregar
    time.sleep(4)
       
abrir_navegador()
abrir_scihub()

# import time
# import pyautogui
# from pynput import mouse, keyboard

# gravando = True
# espacos_pressionados = 0
# eventos = []


# def on_click(x, y, button, pressed):
#     global eventos

#     if pressed and gravando:
#         pos_mouse = pyautogui.position()

#         evento = {
#             "timestamp": time.time(),
#             "botao": str(button),
#             "x": pos_mouse.x,
#             "y": pos_mouse.y
#         }

#         eventos.append(evento)

#         print(
#             f"Clique: {button} | Posição: ({pos_mouse.x}, {pos_mouse.y})"
#         )


# def on_press(key):
#     global gravando, espacos_pressionados

#     try:
#         if key == keyboard.Key.space:
#             espacos_pressionados += 1
#             print(
#                 f"Espaço pressionado ({espacos_pressionados}/2)"
#             )

#             if espacos_pressionados >= 2:
#                 gravando = False
#                 print("\nGravação finalizada!")
#                 return False

#     except Exception as e:
#         print(f"Erro: {e}")


# print("Gravando cliques...")
# print("Pressione ESPAÇO duas vezes para parar.\n")

# mouse_listener = mouse.Listener(on_click=on_click)
# keyboard_listener = keyboard.Listener(on_press=on_press)

# mouse_listener.start()
# keyboard_listener.start()

# keyboard_listener.join()

# mouse_listener.stop()

# print("\nEventos capturados:")
# for i, evento in enumerate(eventos, start=1):
#     print(
#         f"{i}. Botão: {evento['botao']} | "
#         f"Posição: ({evento['x']}, {evento['y']})"
#     )

# # Salvar em arquivo
# with open("cliques_gravados.txt", "w", encoding="utf-8") as arquivo:
#     for evento in eventos:
#         arquivo.write(
#             f"{evento['timestamp']};"
#             f"{evento['botao']};"
#             f"{evento['x']};"
#             f"{evento['y']}\n"
#         )

# print("\nArquivo salvo como: cliques_gravados.txt")