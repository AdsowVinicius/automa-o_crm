from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from urllib.parse import quote
from time import sleep
from datetime import datetime
import openpyxl
import re


def padronizar_data(data):
    """Padroniza a data para o formato datetime.date."""
    if isinstance(data, datetime):
        return data.date()
    elif isinstance(data, str):
        data = data.strip()
        try:
            if len(data.split('/')) == 2 or len(data.split('-')) == 2:
                data_com_ano = f"{data}/{datetime.now().year}"
                return datetime.strptime(data_com_ano, "%d/%m/%Y").date()
            else:
                return datetime.strptime(data, "%d/%m/%Y").date()
        except ValueError:
            try:
                return datetime.strptime(data, "%d-%m-%Y").date()
            except ValueError:
                return None
    elif isinstance(data, (int, float)):
        try:
            return datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(data) - 2).date()
        except Exception:
            return None
    return None


def formatar_nome(nome):
    """Formata o nome para ter a primeira letra maiúscula de cada palavra."""
    return " ".join(p.capitalize() for p in nome.strip().split()) if nome else "Cliente"


def formatar_telefone(telefone, codigo_pais='+55'):
    """Formata o número de telefone, removendo caracteres especiais e adicionando o código do país."""
    if telefone:
        telefone_formatado = re.sub(r'\D', '', telefone)
        if len(telefone_formatado) == 11 or len(telefone_formatado) == 12:
            return f'{codigo_pais}{telefone_formatado}'
        else:
            raise ValueError(f"Número de telefone inválido: {telefone}")
    return None


# Configuração do navegador (substitua pelo caminho do seu WebDriver)
driver_path = r'C:\caminho\para\chromedriver.exe'  # Substitua pelo caminho correto
driver = webdriver.Chrome(driver_path)

# Abertura do WhatsApp Web
driver.get('https://web.whatsapp.com/')
print("Por favor, escaneie o código QR no WhatsApp Web.")
sleep(30)  # Aguarda tempo suficiente para escanear o QR Code

# Abrindo a planilha
arquivo_excel = r'C:\Users\Comercial\Desktop\automa-o_crm-master\CRM (2).xlsx'
planilha = openpyxl.load_workbook(arquivo_excel)
pagina = planilha['Copy of lista de contato diario']

# Loop pelas linhas da planilha
for linha in pagina.iter_rows(min_row=2):  # Ignora o cabeçalho
    nome = linha[2].value
    telefone = linha[3].value
    data_contato = linha[0].value
    mensagem_especifica = linha[6].value  # Motivo para contato (MEC)
    
    # Recuperando os "próximos contatos"
    proximos_contatos = [
        linha[5].value,  # 1º Próximo Contato
        linha[7].value,  # 2º Próximo Contato
        linha[9].value,  # 3º Próximo Contato
        linha[11].value  # 4º Próximo Contato
    ]

    if not nome and not telefone and not data_contato and not mensagem_especifica:
        print("Fim dos dados na planilha.")
        break

    try:
        # Validar e formatar campos
        nome = formatar_nome(nome) if nome else "Cliente"
        telefone = formatar_telefone(str(telefone)) if telefone else None
        data_contato = padronizar_data(data_contato)
        hoje = datetime.now().date()

        # Verifica e processa cada próximo contato
        for i, proximo_contato in enumerate(proximos_contatos, start=1):
            if proximo_contato:
                proximo_contato_data = padronizar_data(proximo_contato)

                if proximo_contato_data == hoje:
                    mensagem = f"Olá, {nome}! {mensagem_especifica} (Próximo Contato {i})"
                    mensagem_encoded = quote(mensagem)
                    link = f"https://web.whatsapp.com/send?phone={telefone}&text={mensagem_encoded}"
                    
                    # Abrir o link no WhatsApp
                    driver.get(link)
                    sleep(10)  # Aguardar carregamento da página

                    # Localizar o campo de mensagem e pressionar ENTER
                    try:
                        campo_mensagem = driver.find_element(By.XPATH, '//div[@title="Digite uma mensagem aqui"]')
                        campo_mensagem.send_keys(Keys.RETURN)  # Envia a mensagem
                        print(f"Mensagem enviada para {telefone}: {mensagem}")
                        sleep(5)
                    except Exception as e:
                        print(f"Erro ao enviar mensagem para {nome}: {e}")

    except Exception as e:
        print(f"Erro ao processar o contato {nome}: {e}")
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f"{nome},{telefone},{data_contato},{e}\n")

# Fechar o navegador
driver.quit()
