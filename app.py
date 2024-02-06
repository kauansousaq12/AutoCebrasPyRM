from tkinter import *
from tkinter.ttk import *
import tkinter as tk
from tkinter import messagebox
import os
from datetime import datetime
import shutil
import filecmp
import subprocess
import pyautogui
import time
import win32api
import winreg
from pywinauto import application, timings
import re

def etapas_realiazada():
    caminho_area_trabalho = os.path.join(os.path.expanduser("~"), "Desktop")
    novo_caminho_pasta = os.path.join(caminho_area_trabalho, formato_data)

    caminho_arquivo_txt = os.path.join(novo_caminho_pasta, 'rm-diretorios.txt')

    with open(caminho_arquivo_txt, 'r') as arquivo_txt:
            conteudo_txt = arquivo_txt.read()
def centralizar_janela(janela):
    janela.update_idletasks()
    largura_janela = janela.winfo_width()
    altura_janela = janela.winfo_height()

    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()

    x = (largura_tela - largura_janela) // 2
    y = (altura_tela - altura_janela) // 3

    janela.geometry("+{}+{}".format(x, y))
def verificar_pasta_existente(caminho_pasta):
    return os.path.exists(caminho_pasta)
def criar_pasta_data_atual(label_status, botao_copiar):
    try:
        data_atual = datetime.now()
        formato_data = data_atual.strftime("%Y%m%d" + "_RM") 

        caminho_area_trabalho = os.path.join(os.path.expanduser("~"), "Desktop")

        novo_caminho_pasta = os.path.join(caminho_area_trabalho, formato_data)

        if verificar_pasta_existente(novo_caminho_pasta):
            print(f"A pasta já existe: {novo_caminho_pasta}")
            label_status.config(text="Status: Não Copiado (Pasta já existe)", fg="orange")
            botao_copiar.config(state=tk.DISABLED) 
            return

        os.makedirs(novo_caminho_pasta)
        print(f"Pasta criada com sucesso: {novo_caminho_pasta}")

        path = r'\\172.16.0.22\Home\TI\RMInstaller'
        destino_path = novo_caminho_pasta

        for file_name in os.listdir(path):
            full_file_path_source = os.path.join(path, file_name)
            full_file_path_destino = os.path.join(destino_path, file_name)
            shutil.copy(full_file_path_source, full_file_path_destino)

        print(f"Arquivos copiados para: {destino_path}")
        label_status.config(text="Status: Copiado", fg="green")

    except Exception as e:
        print(f"Erro: {str(e)}")
        label_status.config(text="Status: Não Copiado (Erro)", fg="red")
def verificar_aplcativo_ativo(aplicativo):
    try:
        subprocess.check_output(f'where {aplicativo}', shell=True)
        return True
    except subprocess.CalledProcessError:
        return False
def verificar_instalacao_registro(caminho_chave):
    try:
        chave_registro = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, caminho_chave, 0, winreg.KEY_READ | winreg.KEY_WOW64_32KEY)
        winreg.CloseKey(chave_registro)
        return True
    except FileNotFoundError:
        return False
    except Exception as e:
        print(f"Erro ao acessar o Registro: {str(e)}")
        return False


def desinstalar_programa_direto():
    try:
        # Use "uia" backend
        app = application.Application(backend="uia")

        # Abrir o Painel de Controle diretamente na seção "Programas e Recursos"
        app.start("control.exe /name Microsoft.ProgramsAndFeatures")
        time.sleep(3)  # Aguardar um pouco para o Painel de Controle abrir

        script_dir = os.path.dirname(os.path.abspath(__file__))
        image_path = os.path.join(script_dir, 'programas_recursos_item1.png')

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

        time.sleep(3)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        image_path = os.path.join(script_dir, 'imagem_biblioteca.png')

        try:
            # Localizar a posição da imagem na tela
            img = pyautogui.locateCenterOnScreen(image_path, confidence=0.8)

            # Se a imagem for encontrada, clicar com o botão direito do mouse
            if img:
                pyautogui.rightClick(img.x, img.y)
            else:
                print("Imagem não encontrada. Continuando...")#imagem_desinstalar

        except pyautogui.ImageNotFoundException:
            print("Imagem não encontrada. Continuando...")

        time.sleep(3)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        image_path = os.path.join(script_dir, 'imagem_desinstalar.png')

        try:
            # Localizar a posição da imagem na tela
            img = pyautogui.locateCenterOnScreen(image_path, confidence=0.8)

            # Se a imagem for encontrada, clicar com o botão direito do mouse
            if img:
                pyautogui.click(img.x, img.y)
            else:
                print("Imagem não encontrada. Continuando...")#imagem_desinstalar_0.2

        except pyautogui.ImageNotFoundException:
            print("Imagem não encontrada. Continuando...")
        time.sleep(3)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        image_path = os.path.join(script_dir, 'imagem_desinstalar_caixalt.png')

        try:
            # Localizar a posição da imagem na tela
            img = pyautogui.locateCenterOnScreen(image_path, confidence=0.8)

            # Se a imagem for encontrada, clicar com o botão direito do mouse
            if img:
                pyautogui.click(img.x, img.y)
            else:
                print("Imagem não encontrada. Continuando...")#imagem_desinstalar_final

        except pyautogui.ImageNotFoundException:
            print("Imagem não encontrada. Continuando...")

        time.sleep(3)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        image_path = os.path.join(script_dir, 'imagem_desinstalar_final.png')

        try:
            # Localizar a posição da imagem na tela
            img = pyautogui.locateCenterOnScreen(image_path, confidence=0.8)

            # Se a imagem for encontrada, clicar com o botão direito do mouse
            if img:
                pyautogui.click(img.x, img.y)
            else:
                print("Imagem não encontrada. Continuando...")#imagem_desinstalar_finalse

        except pyautogui.ImageNotFoundException:
            print("Imagem não encontrada. Continuando...")
        time.sleep(20)
        loop_verificar_e_reiniciar()
    except Exception as e:
        print(f"{str(e)}")

def verificar_aplicacao_rm():
    # Verificar se a aplicação RM está ativa
    return verificar_aplcativo_ativo("NomeDoAplicativoRM")

def loop_verificar_e_reiniciar():
    while True:
        if not verificar_aplicacao_rm():
            print("A aplicação RM não está mais instalada. Reiniciando o computador...")
            
            time.sleep(5)
            
            os.system("shutdown /r /t 1")
            
            time.sleep(60)  
        else:
            print("A aplicação RM está instalada.")

        time.sleep(60)  
def desinstalar_aplicacao(nome_chave):

    def _remover_do_startup(nome_chave):
        if verificar_chave_existente(nome_chave):
            chave_registro = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
            winreg.DeleteValue(chave_registro, nome_chave)
            winreg.CloseKey(chave_registro)
            print(f"A aplicação foi removida do Registro de inicialização com sucesso.")
        else:
            print(f"A aplicação não está presente no Registro de inicialização.")

    def _desinstalar():
        try:
            # Abrir o Painel de Controle
            app = application.Application()
            app.start("control.exe")
            timings.sleep(2)  # Aguardar um pouco para o Painel de Controle abrir

            # Focar na janela do Painel de Controle
            painel_controle = app.window(title="Painel de Controle")
            painel_controle.set_focus()

            # Esperar até que a janela do Painel de Controle esteja totalmente carregada
            timings.wait_until(lambda: painel_controle.is_active(), timeout=10)

            # Abrir a seção "Programas e Recursos"
            programas_recursos = painel_controle.child_window(title="Programas e Recursos", control_type="Pane")
            programas_recursos.click()

            # Esperar para garantir que a lista de programas seja carregada
            timings.sleep(2)

            # Localizar o programa na lista
            programa = painel_controle.child_window(title_re=".*BibliotecaRM -.*", control_type="DataItem")

            # Clicar com o botão direito no programa e selecionar "Desinstalar"
            programa.right_click_input()
            context_menu = app.window(title="Context")
            desinstalar_item = context_menu.menu_item("Desinstalar")

            if desinstalar_item.exists():
                desinstalar_item.click_input()

        except Exception as e:
            print(f"Erro ao desinstalar o programa: {str(e)}")


    #messagebox.showwarning(title="Aguarde 10 segundos", message="Para desinstalar o RM, fique atento!")
    #time.sleep(10)

    #_remover_do_startup(nome_chave)

    #thread_desinstalacao = threading.Thread(target=_desinstalar)
    #thread_desinstalacao.start()



root = Tk()
root.title('Instalador RM')
root.geometry("500x280")
root.eval('tk::PlaceWindow . center')

# Botão de Copiar Arquivo
label_copiar = tk.Label(root, text="Ao clicar cria uma pasta na sua área de trabalho")
label_unistall = tk.Label(root, text="Ao clicar ele faz automaticamente a desintalação do RM")
label_status = tk.Label(root, text="")

label_instalar = tk.Label(root, text="Essa parte realiza toda a instalação automatica do RM é só preciso esperar")
#botao_verificar_copia = tk.Button(root, text="Verificar autenticação", command=verificar_autenticacao)
botao_copiar = tk.Button(root, text='Copiar Arquivos', command=lambda: criar_pasta_data_atual(label_status, botao_copiar))
botao_unistall = tk.Button(root, text='Realizar a Desintalação',  command=lambda: desinstalar_programa_direto())

botao_remover_registro =  tk.Button(root, text='Remover do Regedit', command=lambda: remover_do_startup(nome_chave))

#instalação
botao_instalar = tk.Button(root, text='Instalar RM')


botao_copiar.pack(pady=25)
botao_unistall.pack(pady=25)
#botao_verificar_copia.pack(pady=40)
label_copiar.pack(pady=25)
label_status.pack(pady=25)
label_instalar.pack(pady=25)
botao_unistall.pack(pady=25)
botao_remover_registro.pack(pady=25)

botao_instalar.pack(pady=25)

botao_copiar.place(x=40, y=60)
botao_unistall.place(x=40, y=130)
botao_remover_registro.place(x=180, y=130)
botao_instalar.place(x=40, y=190)
#botao_verificar_copia.place(x=140, y=60)
label_copiar.place(x=40, y=40)
label_instalar.place(x=40, y=170)
label_unistall.place(x=40, y=110)
label_status.place(x=40, y=85)

# Verificar se a pasta já existe ao iniciar o programa
caminho_area_trabalho = os.path.join(os.path.expanduser("~"), "Desktop")
formato_data_inicial = datetime.now().strftime("%Y%m%d" + "_RM")
pasta_inicial = os.path.join(caminho_area_trabalho, formato_data_inicial)

arquivo_servidor_local = r'\\172.16.0.22\Home\TI\RMInstaller\rm-diretorios.txt'

# Caminho para o arquivo na área de trabalho do usuário
caminho_area_trabalho = os.path.join(os.path.expanduser("~"), "Desktop")
formato_data = datetime.now().strftime("%Y%m%d" + "_RM")
caminho_arquivo_area_trabalho = os.path.join(caminho_area_trabalho, formato_data, 'rm-diretorios.txt')

# Verificar se os arquivos são iguais
def verificar_autenticacao_arquivos():
    try:
        if filecmp.cmp(arquivo_servidor_local, caminho_arquivo_area_trabalho):
            messagebox.showinfo("Verificação de Autenticação", "Os arquivos já existem e são veridicos em seu computador!")
        else:
            messagebox.showerror("Verificação de Autenticação", "Os arquivos foram encontrados mas constam como diferentes seus conteudos!")

    except FileNotFoundError:
        messagebox.showerror("Verificação de Autenticação", "Arquivo de copia não foi encontrado!")
    except Exception as e:
        messagebox.showerror("Verificação de Autenticação", f"Erro: {str(e)}")


if verificar_pasta_existente(pasta_inicial):
    botao_copiar.config(state=tk.DISABLED)
    label_status.config(text="Status: Não Copiado (Pasta já existe) Está tudo no conforme!", fg="green")


caminho_chave = r"SOFTWARE\WOW6432Node\TOTVS"

instalado = verificar_instalacao_registro(caminho_chave)

if instalado:
    print(f"O RM está instalado no Registro.")
    botao_unistall.config(state=tk.NORMAL)
    botao_instalar.config(state=tk.DISABLED)
else:
    print(f"O RM não está instalado no Registro.")
    botao_unistall.config(state=tk.DISABLED)
    botao_instalar.config(state=tk.NORMAL)

def verificar_chave_existente(nome_chave):
    try:
        chave_registro = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
        valor, tipo = winreg.QueryValueEx(chave_registro, nome_chave)
        winreg.CloseKey(chave_registro)
        return True
    except FileNotFoundError:
        return False

def adicionar_ao_startup(nome_chave, valor):
    if not verificar_chave_existente(nome_chave):
        chave_registro = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(chave_registro, nome_chave, 0, winreg.REG_SZ, valor)
        winreg.CloseKey(chave_registro)
        print(f"A aplicação foi adicionada ao Registro de inicialização com sucesso.")
    else:
        print(f"A aplicação já está presente no Registro de inicialização.")

def remover_do_startup(nome_chave):
    if verificar_chave_existente(nome_chave):
        chave_registro = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
        winreg.DeleteValue(chave_registro, nome_chave)
        winreg.CloseKey(chave_registro)
        print(f"A aplicação foi removida do Registro de inicialização com sucesso.")
    else:
        print(f"A aplicação não está presente no Registro de inicialização.")


nome_chave = "AutomatizarRM"
caminho_executavel = r"C:\Users\antonykaua\Desktop\AutoCerbrasAtualizador\dist\app.exe"

centralizar_janela(root)
verificar_autenticacao_arquivos();
remover_do_startup(nome_chave)

mainloop() 