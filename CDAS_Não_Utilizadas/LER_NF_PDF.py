import os
import pandas as pd
import fitz  # PyMuPDF
import re

def ler_pdf(arquivo):
    texto = ""
    with fitz.open(arquivo) as pdf_file:
        for page in pdf_file:
            texto += page.get_text()
    return texto

def extrair_valor_total(texto):
    # Usando expressão regular para encontrar o valor total
    padrao_valor_total = r'VALOR TOTAL DA NOTA-RS\s+(\d+,\d{2})'
    resultado = re.search(padrao_valor_total, texto)
    if resultado:
        return resultado.group(1)
    return None

def localizar_descricao(texto):
    # Usando expressão regular para encontrar a descrição local
    padrao_descricao_local = r'Local:\s+([^CNL]+)'
    resultado = re.search(padrao_descricao_local, texto)
    if resultado:
        return resultado.group(1).strip()[:30]
    return None

def extrair_numero_nota(texto):
    # Usando expressão regular para encontrar o número da nota
    padrao_numero_nota = r'Numero da Nota\s+(\d{8})'
    resultado = re.search(padrao_numero_nota, texto)
    if resultado:
        return resultado.group(1)
    return None

def extrair_data_emissao(texto):
    # Usando expressão regular para encontrar a data de emissão
    padrao_data_emissao = r'\b\d{1,2}/\d{1,2}/\d{4}\b'
    resultado = re.search(padrao_data_emissao, texto)
    if resultado:
        return resultado.group()
    return None

def main():
    pasta_notas = r'C:\ProjetorPython\Leitor_NF_PDF\NF_OCR'
    dados = []  # Lista para armazenar os dados extraídos
    for arquivo in os.listdir(pasta_notas):
        if arquivo.lower().endswith('.pdf'):
            texto = ler_pdf(os.path.join(pasta_notas, arquivo))
            valor_total = extrair_valor_total(texto)
            descricao_local = localizar_descricao(texto)
            numero_nota = extrair_numero_nota(texto)
            data_emissao = extrair_data_emissao(texto)
            if descricao_local:
                dados.append({'Arquivo': arquivo, 'Numero da Nota': numero_nota, 'Data de Emissao': data_emissao, 'Valor Total': valor_total, 'Descrição Local': descricao_local})

    # Cria um DataFrame com os dados
    df = pd.DataFrame(dados)

    # Salva o DataFrame em um arquivo Excel
    nome_arquivo_excel = 'notas_fiscais.xlsx'
    df.to_excel(nome_arquivo_excel, index=False)
    print(f"Dados salvos no arquivo {nome_arquivo_excel}")

if __name__ == "__main__":
    main()

