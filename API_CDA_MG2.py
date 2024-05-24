
import requests
from bs4 import BeautifulSoup
import pandas as pd

def extrai_informacoes(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')

    tabela = soup.find('table')
    linhas = tabela.find_all('tr')

    informacoes = []
    for linha in linhas[1:]:
        colunas = linha.find_all('td')
        informacoes.append([coluna.text.strip() for coluna in colunas])

    return informacoes

def main():
    # Carrega o arquivo excel com os cda
    df = pd.read_excel('Análise das CDAs-BV_HISTORICO.xlsx')

    # Extrai as informações de cada cda
    informacoes = []
    for index, row in df.iterrows():
        url = f"https://www2.fazenda.mg.gov.br/sol/ctrl/SOL/DIVATIV/SERVICO_001?ACAO=IMPRIMIR&CDA={row['cda']}"
        informacoes_cda = extrai_informacoes(url)
        informacoes.extend(informacoes_cda)

    # Salva as informações extraídas em um arquivo excel
    df_informacoes = pd.DataFrame(informacoes, columns=['cda', 'nome', 'cpf', 'data_nascimento', 'endereco'])
    df_informacoes.to_excel('informacoes_extraidas.xlsx', index=False)

if __name__ == "__main__":
    main()