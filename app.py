from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from urllib.parse import quote
from time import sleep
import openpyxl
from datetime import datetime
from utils.padronizarTelefone import formatar_telefone
from utils.formatarNome import formatar_nome
from utils.padronizarData import padronizar_data

# Configuração do navegador (substitua pelo caminho do seu WebDriver)
try:
    driver = webdriver.Chrome()
    print("Navegador iniciado com sucesso.")
except Exception as e:
    print(f"Erro ao iniciar o navegador: {e}")
    exit()

# Abertura do WhatsApp Web
try:
    driver.get('https://web.whatsapp.com/')
    print("WhatsApp Web aberto. Por favor, escaneie o código QR.")
    sleep(30)  # Aguarda tempo suficiente para escanear o QR Code
except Exception as e:
    print(f"Erro ao abrir o WhatsApp Web: {e}")
    driver.quit()
    exit()

# Abrindo a planilha
arquivo_excel = r'C:\desenv\python\RPAs(automaçoes)\bot_para_envio_de_mensagens\CRM (2).xlsx'
try:
    planilha = openpyxl.load_workbook(arquivo_excel)
    pagina = planilha['Copy of lista de contato diario']
    print("Planilha carregada com sucesso.")
except Exception as e:
    print(f"Erro ao carregar a planilha: {e}")
    driver.quit()
    exit()

# Loop pelas linhas da planilha
for linha in pagina.iter_rows(min_row=2):  # Ignora o cabeçalho
    try:
        # Leitura dos dados da linha
        nome = linha[2].value
        telefone = linha[3].value
        data_contato = linha[0].value
        mensagem_especifica = linha[6].value  # Motivo para contato (MEC)
        
        proximos_contatos = [
            linha[5].value,  # 1º Próximo Contato
            linha[7].value,  # 2º Próximo Contato
            linha[9].value,  # 3º Próximo Contato
            linha[11].value  # 4º Próximo Contato
        ]

        print(f"Processando contato: Nome={nome}, Telefone={telefone}, DataContato={data_contato}")

        # Validação e formatação dos campos
        if not telefone:
            print("Telefone ausente, ignorando linha.")
            continue

        nome = formatar_nome(nome) if nome else "Cliente"
        telefone = formatar_telefone(str(telefone))
        data_contato = padronizar_data(data_contato) if data_contato else None
        hoje = datetime.now().date()

        # Verifica e processa cada próximo contato
        for i, proximo_contato in enumerate(proximos_contatos, start=1):
            if proximo_contato:
                proximo_contato_data = padronizar_data(proximo_contato)

                if proximo_contato_data == hoje:
                    print(f"Próximo contato válido encontrado para {nome}.")
                    mensagem = f"Olá, {nome}! {mensagem_especifica} (Próximo Contato {i})"
                    mensagem_encoded = quote(mensagem)
                    link = f"https://web.whatsapp.com/send?phone={telefone}&text={mensagem_encoded}"
                    
                    # Abrir o link no WhatsApp
                    try:
                        driver.get(link)
                        print(f"Enviando mensagem para {telefone}: {mensagem}")
                        sleep(10)  # Aguardar carregamento da página

                        # Localizar o campo de mensagem e pressionar ENTER
                        campo_mensagem = driver.find_element(By.XPATH, '//div[@title="Digite uma mensagem aqui"]')
                        campo_mensagem.send_keys(Keys.RETURN)  # Envia a mensagem
                        print(f"Mensagem enviada para {telefone}: {mensagem}")
                        sleep(5)
                    except Exception as e:
                        print(f"Erro ao enviar mensagem para {nome} ({telefone}): {e}")

    except Exception as e:
        print(f"Erro ao processar linha: {e}")
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f"{nome},{telefone},{data_contato},{e}\n")

# Fechar o navegador
driver.quit()
print("Processo concluído. Navegador fechado.")
