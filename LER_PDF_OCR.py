import os
import pdfplumber
import comtypes.client
import pandas as pd
import re

def abrir_pdf_no_word(arquivo_pdf, arquivo_word):
    wdFormatPDF = 17
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(arquivo_pdf)
    doc.SaveAs(arquivo_word, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

def extrair_texto_pdf(arquivo_pdf):
    with pdfplumber.open(arquivo_pdf) as pdf:
        texto = ''
        for page in pdf.pages:
            texto += page.extract_text()
    return texto

def extrair_valor_total(texto):
    # Localiza o texto "VALOR TOTAL DA NOTA-RS"
    indice_inicio = texto.find("VALOR TOTAL DA NOTA-RS")
    if indice_inicio != -1:
        # Busca o próximo caractere numérico após a string encontrada
        indice_valor_inicio = indice_inicio + len("VALOR TOTAL DA NOTA-RS")
        while not texto[indice_valor_inicio].isdigit():
            indice_valor_inicio += 1
        # Encontra o próximo caractere não numérico após o valor e extrai o valor total
        indice_valor_final = indice_valor_inicio
        while texto[indice_valor_final].isdigit() or texto[indice_valor_final] == ',':
            indice_valor_final += 1
        valor_total = texto[indice_valor_inicio:indice_valor_final]
        return valor_total
    return None  # Caso não encontre o texto

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
    pasta_notas = r'C:\ProjetorPython\Leitor_NF_PDF\NF_OCR'
    pasta_temporaria = r'C:\ProjetorPython\files\temp'  # Pasta temporária para armazenar os arquivos do Word
    dados = []  # Lista para armazenar os dados extraídos

    # Abrir cada arquivo PDF no Word e extrair texto
    for arquivo in os.listdir(pasta_notas):
        if arquivo.lower().endswith('.pdf'):
            arquivo_pdf = os.path.join(pasta_notas, arquivo)
            arquivo_word = os.path.join(pasta_temporaria, f"{os.path.splitext(arquivo)[0]}.docx")
            abrir_pdf_no_word(arquivo_pdf, arquivo_word)
            texto = extrair_texto_pdf(arquivo_pdf)
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








