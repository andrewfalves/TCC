import csv                          #importando biblioteca csv(arquivos csv)
from pprint import pprint           #importando biblioteca pprint(impressão em linhas diferentes)
import gspread                      # importando biblioteca gspread
from oauth2client.service_account import ServiceAccountCredentials  #importando biblioteca de credenciais de conta

# Configuração para acessar documentos no google drive
scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

parametros = ServiceAccountCredentials.from_json_keyfile_name("parametros.json", scope)

client = gspread.authorize(parametros)

sheet = client.open("IR 2020").sheet1  # abre a planilha no google sheet

data = sheet.get_all_records()  # todos os valores esritos na tabela

#pprint(len(data))  # Numero de linhas preenchidas na tabela

#print(LastLine)  # imprime qual a ultima linha

with open('D:\Downloads\Easyinvest\ResumoNegociacaoJunho.csv','r') as file: #Caminho do arquivo

    reader = csv.reader(file)       #leitura do csv

    dados = list(file)              #criando uma lista a partir dos dados lidos

    print (dados)                   #impressão da lista
    numero_elementos = (len(dados)) #contagem do numero de linhas dos dados
    print(numero_elementos)         #impressão do numero de elementos

    index = 1                       #definindo uma variável de indexação

    while (index < (numero_elementos)): #loop. Enquanto index for menor que o numero de linhas

        LastLine = len(data) + 2  # Escreve na proxima linha em branco
        #print(dados[index])

        dados[index] = dados[index].split(';')
    # os dados foram criados como elementos únicos, este comando separa cada elemento usando o ";" como separador

    ##Manipulação dos dados
        del (dados[index][1])
        #deleta o numero da conta
        qtde1 = int(dados[index][3])
        qtde2=  int(dados[index][4])
        (dados[index][3])=  qtde1+qtde2
        #as colunas 3 e 4 são referentes ao volume negociado, soma-se os 2
        del (dados[index][4])
        #exclui uma das colunas após a soma de quantidades
        a=(dados[index][2])
        #define uma váriável para receber o valor da coluna 2
        b=(dados[index][3])
        # define uma váriável para receber o valor da coluna 3
        dados[index][3]=  a
        # Troca a ordem das colunas 2 e 3
        dados[index][2] = b
        # Troca a ordem das colunas 2 e 3
        index = index+1
        #soma 1 ao index

    pprint(dados)
    #impressão da lista final
    index_atualizar = 1
    while (index_atualizar  < numero_elementos):
        L1 = index_atualizar
        C1=  dados[L1][0]
        C2 = dados[L1][1]
        C3 = dados[L1][2]
        C4 = dados[L1][3]
        print (index_atualizar)
        sheet.update_cell((L1+1), 1, C1)  # atualiza a célula especifica
        sheet.update_cell((L1+1), 2, C2)  # atualiza a célula especifica
        sheet.update_cell((L1+1), 3, C3)  # atualiza a célula especifica

        if ((dados[L1][4]) != "0"):
            sheet.update_cell((L1 + 1), 4, C4)  # atualiza a célula especifica

        else:
            sheet.update_cell((L1 + 1), 7, C4)  # atualiza a célula especifica

        index_atualizar = (index_atualizar+1)