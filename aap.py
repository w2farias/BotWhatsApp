import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

webbrowser.open('https://web.whatsapp.com/')
sleep(30)


workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2, max_row=6):

    nome = linha[0].value
    telefone = linha[1].value

    mensagem = f'Olá {
        nome} aqui é Wellington estou testando um novo progama que acabei de criar'

    try:
        link_menagem_whatsapp = f'https://web.whatsapp.com/send?phone={
            telefone}&text={quote(mensagem)}'
        webbrowser.open(link_menagem_whatsapp)
        sleep(10)

        pyautogui.hotkey('enter')
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Não foi possível enviar mensagem para{nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8')as arquivo:

            arquivo.write(f'{nome}, {telefone}{os.linesep}')
