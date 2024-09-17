import pandas as pd
from datetime import datetime

# Dados fornecidos
data = {
    "Data": ["22/05/2024", "16/10/2023", "20/10/2023", "16/08/2024", "27/01/2023", "20/10/2023", "10/11/2023", "19/05/2023"],
    "Nome": ["Kamila"] * 8,
    "Sobrenome": ["Duque Honorato da Silva"] * 8
}

# Criar DataFrame
df = pd.DataFrame(data)

# Converter coluna de datas para datetime
df["Data"] = pd.to_datetime(df["Data"], format="%d/%m/%Y")

# Definir a data pesquisada e o intervalo de datas
data_pesquisada = datetime.today()
data_inicial = data_pesquisada - pd.DateOffset(months=3)
data_final = data_pesquisada

# Encontrar a última data de submissão
ultima_data_submissao = df["Data"].max()

# Verificar se a última data de submissão está dentro do intervalo
if data_inicial <= ultima_data_submissao <= data_final:
    bloqueados = df[df["Data"] == ultima_data_submissao]
else:
    bloqueados = pd.DataFrame()

# Registros restantes (não movidos para Clientes Liberados)
liberados = pd.DataFrame()

# Contador de registros (contar apenas Nome e Sobrenome)
contador = len(df)

# Resultados esperados
print("Clientes Bloqueados:")
print(bloqueados)
print("\nClientes Liberados:")
print(liberados)
print("\nContador:", contador)
