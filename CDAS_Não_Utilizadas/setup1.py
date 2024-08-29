import os
import pdfplumber
import pandas as pd
import re

def ler_pdf(arquivo):
    with pdfplumber.open(arquivo) as pdf_file:
        texto = ''.join([page.extract_text() for page in pdf_file.pages if page.extract_text()])
        return texto

def extrair_numero_nota(texto):
    padrao_numero_nota = r"Número da Nota\s*\n\s*(\d+)"
    resultado = re.search(padrao_numero_nota, texto)
    if resultado:
        return resultado.group(1).strip()
    return None

def extrair_data_emissao(texto):
    padrao_data = r'\d{2}/\d{2}/\d{4}'
    resultado = re.search(padrao_data, texto)
    if resultado:
        return resultado.group()
    return None

def extrair_valor_total(texto):
    padrao_valor_total = r"VALOR TOTAL DA NOTA-R\$([\d,.]+)"
    resultado = re.search(padrao_valor_total, texto)
    if resultado:
        valor_total = resultado.group(1).replace(",", ".")
        return valor_total
    return None

def extrair_descricao_local(texto):
    padrao_descricao_local = r"Local:(.*?)CNL"
    resultado = re.search(padrao_descricao_local, texto, re.DOTALL)
    if resultado:
        descricao_local = resultado.group(1).strip()
        return descricao_local
    return None

pasta_notas = r'C:\ProjetorPython\files\OneDrive_1_29-04-2024\NF_OCR2'
dados = []

for arquivo in os.listdir(pasta_notas):
    if arquivo.lower().endswith('.pdf'):
        caminho_arquivo = os.path.join(pasta_notas, arquivo)
        texto = ler_pdf(caminho_arquivo)
        
        numero_nota = extrair_numero_nota(texto)
        data_emissao = extrair_data_emissao(texto)
        valor_total = extrair_valor_total(texto)
        descricao_local = extrair_descricao_local(texto)
        
        dados.append({
            'Arquivo': arquivo,
            'Numero da Nota': numero_nota,
            'Data de Emissao': data_emissao,
            'Valor Total': valor_total,
            'Descricao Local': descricao_local
        })

df = pd.DataFrame(dados)
os.makedirs('output', exist_ok=True)
nome_arquivo_excel = os.path.join('output', 'notas_fiscais.xlsx')
df.to_excel(nome_arquivo_excel, index=False)
print(f"Dados salvos no arquivo {nome_arquivo_excel}")
