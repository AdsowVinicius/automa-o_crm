import openpyxl

try:
    teste = openpyxl.load_workbook(r'C:\desenv\python\RPAs(automaçoes)\bot_para_envio_de_mensagens\teste2.xlsx')
    print("Planilha aberta com sucesso!")
    pagina_teste = teste['Planilha1']
    for linha in pagina_teste.iter_rows(min_row=1):
        print([cell.value for cell in linha])  # Imprime os valores de cada linha
except Exception as e:
    print(f"Ocorreu um erro: {e}")