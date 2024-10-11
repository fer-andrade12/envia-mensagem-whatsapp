import subprocess
import sys
import importlib

# Função para verificar e instalar bibliotecas automaticamente
def verificar_instalar_biblioteca(bibliotecas):
    for biblioteca in bibliotecas:
        try:
            importlib.import_module(biblioteca)
        except ImportError:
            subprocess.check_call([sys.executable, "-m", "pip", "install", biblioteca])

# Lista de bibliotecas necessárias
bibliotecas_necessarias = ['pandas', 'pyautogui', 'time', 'pyperclip', 'emoji', 'openpyxl']

# Verificar e instalar bibliotecas
verificar_instalar_biblioteca(bibliotecas_necessarias)

import pandas as pd
import pyautogui
import time
import pyperclip
import emoji

# Carregar a planilha
file_path = "C:\\Teste\\Teste.xlsx"
df = pd.read_excel(file_path)

# Função para abrir o WhatsApp Desktop (Microsoft Store)
def abrir_whatsapp_desktop():
    subprocess.Popen(["powershell", "Start-Process shell:AppsFolder\\5319275A.WhatsAppDesktop_cv1g1gvanyjgm!WhatsApp"], shell=True)
    time.sleep(5)

# Função para enviar uma mensagem
def enviar_mensagem(nome, contato, mensagem):
    pyautogui.hotkey('ctrl', 'n')
    time.sleep(2)

    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')
    time.sleep(1)

    pyautogui.write(contato)
    time.sleep(3)

    # Pressionar 'Tab' duas vezes para selecionar o contato (sem verificação)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)

    # Enviar mensagem diretamente sem verificação do contato
    pyautogui.press('enter')
    time.sleep(2)

    pyautogui.hotkey('ctrl', 'shift', 'm')
    time.sleep(2)

    pyperclip.copy(mensagem)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)

    pyautogui.press('enter')
    time.sleep(4)
    
    return True

# Função para fechar o WhatsApp após envio das mensagens
def fechar_whatsapp_desktop():
    subprocess.Popen(["powershell", "Stop-Process -Name 'WhatsApp'"], shell=True)

# Abrir o WhatsApp Desktop
abrir_whatsapp_desktop()

contatos_enviados = []

# Iterar sobre as linhas da planilha
for index, row in df.iterrows():
    nome = row['Nome']
    contato = str(row['Contato'])
    print(f"Tentando enviar mensagem para: {nome} - {contato}")

    mensagem = f"""Olá, {nome}

VOCÊ CONCLUIU O CÓDIGO COM SUCESSO!

Estamos ansiosos para vê-lo(a) na próxima turma!"""

    # Tentar enviar a mensagem
    enviar_mensagem(nome, contato, mensagem)
    contatos_enviados.append(f"{nome} - {contato}")

# Exibir mensagem final com os contatos encontrados
if contatos_enviados:
    pyautogui.alert(f"Mensagens enviadas para os seguintes contatos:\n" + "\n".join(contatos_enviados))

# Fechar o WhatsApp após enviar todas as mensagens
fechar_whatsapp_desktop()

