import pandas as pd
from collections import Counter
import random

# Função para ler a planilha e contar os números nas colunas especificadas
def contar_frequencia_numeros(arquivo_excel):
    # Lê a planilha do Excel
    df = pd.read_excel(arquivo_excel)

    # Seleciona as colunas de C até Q (colunas com bolas)
    colunas_bolas = df.iloc[:, 2:18]  # Colunas C até Q correspondem ao índice 2 até 17

    # Conta a frequência de cada número em todas as colunas especificadas
    numeros = colunas_bolas.values.flatten()  # Converte as colunas em um array unidimensional
    contador = Counter(numeros)  # Conta a frequência dos números

    # Filtra para incluir apenas números entre 1 e 25 (baseado nos dados da planilha)
    contador_filtrado = {numero: freq for numero, freq in contador.items() if 1 <= numero <= 25}

    # Ordena os números pela frequência em ordem decrescente
    numeros_ordenados = [numero for numero, _ in sorted(contador_filtrado.items(), key=lambda x: x[1], reverse=True)]

    return numeros_ordenados, contador_filtrado

# Função para gerar 3 combinações aleatórias de 15 números baseadas nos mais frequentes
def gerar_combinacoes_aleatorias(numeros_ordenados, n_combinacoes=3, tamanho_combinacao=15):
    combinacoes = set()
    base_numeros = numeros_ordenados[:25]  # Usar os números mais frequentes

    while len(combinacoes) < n_combinacoes:
        # Gera uma combinação aleatória de 15 números da base de números mais frequentes
        combinacao = tuple(sorted(random.sample(base_numeros, tamanho_combinacao)))
        combinacoes.add(combinacao)  # Adiciona ao conjunto para garantir unicidade

    return list(combinacoes)

# Função para mostrar as combinações geradas e a frequência dos números
def mostrar_frequencia(combinacoes, contador):
    for i, combinacao in enumerate(combinacoes, 1):
        print(f'Combinação {i}: {combinacao}')
        for numero in combinacao:
            print(f'Número {numero} aparece {contador.get(numero, 0)} vezes')
        print()  # Linha em branco para separar combinações

# Caminho para o arquivo Excel
arquivo_excel = r'C:\Users\walter.oliveira\Documents\ProjetosPython\dev\Bichara_Dev\repository\Lotofácil.xlsx'

# Executa a contagem de frequência e gera as combinações aleatórias
numeros_ordenados, contador = contar_frequencia_numeros(arquivo_excel)
combinacoes_aleatorias = gerar_combinacoes_aleatorias(numeros_ordenados)

# Exibe as combinações geradas e a frequência dos números
mostrar_frequencia(combinacoes_aleatorias, contador)
