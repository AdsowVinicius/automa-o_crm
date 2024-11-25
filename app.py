import openpyxl
from urllib.parse import quote 
import webbrowser
from time import sleep
import pyautogui


import re

def formatar_telefone(telefone, codigo_pais='+55'):
    # Remove todos os caracteres que não são números
    telefone_formatado = re.sub(r'\D', '', telefone)
    
    # Verifica se o número tem 11 dígitos e adiciona o código do país
    if len(telefone_formatado) == 11:
        telefone_formatado = f'{codigo_pais}{telefone_formatado}'
    else:
        raise ValueError("Número de telefone inválido")
    
    return telefone_formatado

# Testando a função
try:
    telefone = " (12) 98765-4321 "
    print(formatar_telefone(telefone))  # Saída: +5512987654321
except ValueError as e:
    print(e)


webbrowser.open('https://web.whatsapp.com/')
sleep(10)
teste = openpyxl.load_workbook(r'C:\desenv\python\RPAs(automaçoes)\bot_para_envio_de_mensagens\teste2.xlsx')

pagina_teste = teste['Página1']

for linha in pagina_teste.iter_rows(min_row=1
                                    ):
    nome = linha[0].value
    telefone = linha[1].value
    # Caso o número venha como float, converta para int e depois para string
    telefone = str(int(telefone))
    
    formatar_telefone(telefone, codigo_pais="55" )

    vencimento = linha[2].value
    mensagem = f'olha so {nome} deu certo'
    
    try:
        #https://web.whatsapp.com/send?phone=&text
        link_msg_whatsapp = f'https://web.whatsapp.com/send?phone={(telefone)}&text={quote(mensagem)}'
        print(link_msg_whatsapp)
        webbrowser.open(link_msg_whatsapp)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('seta.PNG')
        sleep(5)
        pyautogui.click(seta[0], seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl','w')
        sleep(5)
    except:
        (print(f'nao foi possivel enviar essa mensagens{nome}'))
        with open('erros.csv','a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}')