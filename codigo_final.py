import csv                          #importando biblioteca csv(arquivos csv)
from pprint import pprint           #importando biblioteca pprint(impressão em linhas diferentes)
import gspread                      # importando biblioteca gspread
from oauth2client.service_account import ServiceAccountCredentials  #importando biblioteca de credenciais de conta
from gspread.models import Cell

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
    Lista2=[]
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
        if ((dados[index][4]) != "0"):
            dados[index][4]=0
        else:
            (dados[index][4])=(dados[index][3])
            (dados[index][3])=0
        del dados[index][-1]
        #Lista2[index]=dados[index][-1]
        #del dados[index][-1]
        index = index+1
        #soma 1 ao index

    pprint(dados)
    #impressão da lista final
    index_atualizar = 1
    cells = []
    for i in range(numero_elementos):
        for j in range(4):
            cells.append(Cell(row=i + 1, col=j + 1, value=dados[i][j]))
        cells.append(Cell(row=i + 1, col=7, value=dados[i][4]))

    pprint(cells)
    sheet.update_cells(cells)
