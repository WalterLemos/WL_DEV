import pyodbc
import openpyxl
from openpyxl import Workbook
from tqdm import tqdm

# Cria um novo arquivo Excel
planilha_resultado = Workbook()

# Seleciona a primeira aba do arquivo
aba_resultado = planilha_resultado.active

# Conecta ao banco de dados Access
conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\walter.oliveira\Documents\Database2.accdb;'
conn = pyodbc.connect(conn_str)

# Obtém um cursor para executar as consultas SQL
cursor = conn.cursor()

# Realiza a consulta na primeira tabela

cursor.execute('SELECT [rd-bichara-advogados].[Identificação], [rd-bichara-advogados].[Empresa] FROM [rd-bichara-advogados] ORDER BY [rd-bichara-advogados].[Identificação];')
# Obtém os resultados da consulta
resultados_tabela1 = cursor.fetchall()

# Realiza a consulta na segunda tabela

cursor.execute('SELECT SISJURI_GRUPO_CLIENTES.[Identificação], SISJURI_GRUPO_CLIENTES.[Codigo Grupo de Empresa], SISJURI_GRUPO_CLIENTES.[Descrição do Grupo de Empresa], SISJURI_GRUPO_CLIENTES.[Razão Social] FROM SISJURI_GRUPO_CLIENTES ORDER BY SISJURI_GRUPO_CLIENTES.[Identificação];')

# Obtém os resultados da consulta
resultados_tabela2 = cursor.fetchall()

# Inicializa a barra de progresso
total_linhas1 = len(resultados_tabela1)
total_linhas2 = len(resultados_tabela2)
barra_progresso = tqdm(total=total_linhas1, desc='Comparando registros', unit='linha')

# Lista para armazenar as alterações
alteracoes = []

# Percorre os registros da primeira tabela
for registro1 in resultados_tabela1:
    campo1 = registro1[0]  # Acessa a primeira coluna pelo índice 0
    campo2 = registro1[1]

    # Percorre os registros da segunda tabela
    for registro2 in resultados_tabela2:
        campo3 = registro2[0]  # Acessa a primeira coluna pelo índice 0
        campo4 = registro2[2]
        campo5 = registro2[3]   

        if campo2 != campo4 and campo2 != campo3:
            # Armazena as informações da linha atualizada na lista de alterações
            alteracoes.append((campo1, campo2, campo3, campo4, campo5))

            # Grava os dados na planilha de resultado
            aba_resultado.append([ campo1, campo2, campo3, campo4, campo5 ])
           
            break

    # Atualiza a barra de progresso
    barra_progresso.update(1)

# Salva o arquivo Excel com os resultados
planilha_resultado.save('resultado4.xlsx')

# Fecha o cursor e a conexão com o banco de dados
cursor.close()
conn.close()

# Finaliza a barra de progresso
barra_progresso.close()

# Imprime as alterações realizadas
print("Alterações:")
for alteracao in alteracoes:
    print(f"Campo1: {alteracao[0]},  Campo3: {alteracao[0]} ")

































