import os
import pdfplumber  # Usando pdfplumber para ler PDFs
import pandas as pd
import re
from tqdm import tqdm

def ler_pdf(arquivo):
    with pdfplumber.open(arquivo) as pdf_file:
        page = pdf_file.pages[0]  # Lê a primeira página (ajuste conforme necessário)
        texto = page.extract_text()
        return texto
    
def extrair_valor_total(texto):
    # Padrao para encontrar o valor total da nota
    padrao_valor_total = r"VALOR TOTAL DA NOTA\s*-\s*RS\s*([\d,.]+)"
    resultado = re.search(padrao_valor_total, texto)
    if resultado:
        valor_total = resultado.group(1).replace(",", ".")  # Substitui vírgulas por pontos
        return valor_total
    return None

def localizar_descricao(texto):
    # Implementa a lógica para localizar a palavra "Local:" na descrição
    # e capturar o texto subsequente
    indice_local = texto.find("Local:")
    if indice_local != -1:
        descricao = texto[indice_local + len("Local:"):].strip()
        indice_cnl = descricao.find("CNL:")
        if indice_cnl != -1:
            descricao = descricao[:indice_cnl].strip()
        return descricao[:30]  # Retorna os primeiros 30 caracteres
    return None  # Caso não encontre a palavra "Local:"

def extrair_numero_nota(texto):
    indice_inicio = texto.find("Numero da Nota")
    if indice_inicio != -1:
        linhas = texto.split('\n')
        for i, linha in enumerate(linhas):
            if "Numero da Nota" in linha:
                if i + 2 < len(linhas):  # Verifica se existe uma segunda linha após o "Numero da Nota"
                    if "CURITIBA" in linhas[i + 1]:
                        indice_curitiba = linhas[i + 1].find("CURITIBA")
                        numero_nota = linhas[i + 1][indice_curitiba + len("CURITIBA"):indice_curitiba + len("CURITIBA") + 8]
                        return numero_nota.strip()
    return None

def extrair_data_emissao(texto):
    padrao_data = r'\b\d{1,2}/\d{1,2}/\d{4}\b'  # Padrão de expressão regular para data (dd/mm/yyyy)
    resultado = re.search(padrao_data, texto)
    if resultado:
        return resultado.group()
    return None

def main():
    pasta_notas = r'C:\Users\walter.oliveira\Documents\NF_OCR'
    dados = []  # Lista para armazenar os dados extraídos

    # Conta quantos arquivos PDF existem na pasta
    total_arquivos = len([nome for nome in os.listdir(pasta_notas) if nome.lower().endswith('.pdf')])
     # Cria uma barra de progresso
    with tqdm(total=total_arquivos, desc="Progresso") as pbar:
        for arquivo in os.listdir(pasta_notas):
            if arquivo.lower().endswith('.pdf'):
                texto = ler_pdf(os.path.join(pasta_notas, arquivo))
                valor_total = extrair_valor_total(texto)
                descricao_local = localizar_descricao(texto)
                numero_nota = extrair_numero_nota(texto)
                data_emissao = extrair_data_emissao(texto)
                if descricao_local:
                    dados.append({'Arquivo': arquivo, 'Numero da Nota': numero_nota, 'Data de Emissao': data_emissao, 'Valor Total': valor_total, 'Descrição Local': descricao_local})
                pbar.update(1)  # Atualiza a barra de progresso

    # Cria um DataFrame com os dados
    df = pd.DataFrame(dados)

    # Salva o DataFrame em um arquivo Excel
    nome_arquivo_excel = 'notas_fiscais.xlsx'
    df.to_excel(nome_arquivo_excel, index=False)
    print(f"Dados salvos no arquivo {nome_arquivo_excel}")

if __name__ == "__main__":
    main()


