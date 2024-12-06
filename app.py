import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

workbook = openpyxl.load_workbook('./clientes.xlsx')

pagina_clientes = workbook['Página1']

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    data_limite = linha[2].value
    valor_devido = linha[3].value
        
    mensagem = f'Olá {nome}, as rematrículas do Colégio Conectado já começaram, garanta sua vaga atá dia {vencimento}, acesse https://site.com.br/boleto'

    try:
        link_whats = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_whats)
        sleep(6)
        pyautogui.press('esc')
        sleep(2)
        pyautogui.press('enter')
        sleep(2)
        pyautogui.hotkey('ctrl', 'w')
        sleep(2)
    except:
        print(f'Envio falhou para {nome}')
        with open('erros.csv','a+',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'\r\n{nome},{telefone}')
        pyautogui.hotkey('ctrl', 'w')


    
    