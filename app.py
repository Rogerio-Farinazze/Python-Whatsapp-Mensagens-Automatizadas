import openpyxl
from urllib.parse import quote
from datetime import date
import webbrowser
from time import sleep
import pyautogui

workbook = openpyxl.load_workbook('clientes.xlsx')

pagina_clientes = workbook['Página1']
data_atual = date.today()

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    mes = data_atual.month
    hoje = data_atual.day
    if (vencimento > hoje):
        print("Teste1")
        mensagem = f'Olá {nome}, seu boleto vence dia {vencimento}/{mes}, acesse https://site.com.br/boleto'
   
        try:
            link_whats = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
            webbrowser.open(link_whats)
            sleep(7)
            seta = pyautogui.locateCenterOnScreen('seta.png')
            sleep(2)
            pyautogui.click(seta[0],seta[1])
            sleep(2)
            pyautogui.hotkey('ctrl', 'w')
            sleep(2)
        except:
            print(f'Envio falhou para {nome}')
            with open('erros.csv','a+',newline='',encoding='utf-8') as arquivo:
                arquivo.write(f'\r\n{hoje}/{mes},{nome},{telefone}')
            pyautogui.hotkey('ctrl', 'w')
    print("Teste2")

