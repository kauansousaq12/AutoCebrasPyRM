import pyautogui
from tkinter import *
from tkinter.ttk import *
import tkinter as tk
from tkinter import messagebox
import os
from datetime import datetime
import time
import subprocess
import tkinter as tk
from tkinter import simpledialog

caminho_imagens = r'\\172.16.0.22\Publico\Antony\ImagensEtapasRm'

def desinstalar_aplicacao():
    try:
        # Construir o comando para desinstalação usando wmic
        comando_desinstalacao = f'wmic product where "name like \'%BibliotecaRM -%\'" call uninstall'

        # Executar o comando de desinstalação
        subprocess.run(comando_desinstalacao, shell=True, check=True)
        print("Aplicações 'BibliotecaRM' desinstaladas com sucesso.")

    except subprocess.CalledProcessError as e:
        print(f"Erro ao desinstalar as aplicações 'BibliotecaRM': {e}")


try:
    subprocess.run(['control'])
except Exception as e:
    print(f"Erro ao abrir o painel de desinstalação: {e}")

time.sleep(5)
# Capturar a tela
script_dir = os.path.dirname(os.path.abspath(__file__))
image_path = os.path.join(script_dir, 'programas_recursos.png')

try:
    # Localizar a posição da imagem na tela
    img = pyautogui.locateCenterOnScreen(image_path, confidence=0.8)

    # Se a imagem for encontrada, clicar com o botão direito do mouse
    if img:
        pyautogui.click(img.x, img.y)
    else:
        print("Imagem não encontrada. Continuando...")

except pyautogui.ImageNotFoundException:
    print("Imagem não encontrada. Continuando...")


#time.sleep(5)
#script_dir2 = os.path.dirname(os.path.abspath(__file__))
#image_path2 = os.path.join(script_dir2, 'imagem_teste.png')

#try:
    # Localizar a posição da imagem na tela
    #img2 = pyautogui.locateCenterOnScreen(image_path2, confidence=0.8)

    # Se a imagem for encontrada, clicar com o botão direito do mouse
    #if img2:
        #pyautogui.rightClick(img2.x, img2.y)
    #else:
        #print("Imagem não encontrada. Continuando...")

#except pyautogui.ImageNotFoundException:
    #print("Imagem não encontrada. Continuando...")

# Abrir o painel de desinstalação
root = tk.Tk()
# Adicione aqui o restante do código para a interface gráfica...
