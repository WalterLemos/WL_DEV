import os
import fitz  # PyMuPDF
import pandas as pd
import re
from tqdm import tqdm
from unidecode import unidecode

def ler_pdf(arquivo):
    try:
        documento = fitz.open(arquivo)
        texto = ""
        for pagina in documento:
            texto += pagina.get_text()
        # Garantir que o texto está codificado em UTF-8, remover quebras de linha e caracteres especiais
        texto_sem_quebra = unidecode(texto).replace('\n', ' ').encode('utf-8').decode('utf-8')
        # Remover caracteres não ASCII
        texto_limpo = re.sub(r'[^\x00-\x7F]+', '', texto_sem_quebra)
        return texto_limpo
    except Exception as e:
        print(f"Erro ao ler o arquivo {arquivo}: {e}")
        return ""

def extrair_valor_total(texto):
    padrao_valor_total = r"VALOR\s+TOTAL\s+DA\s+NOTA-R\$([\d,.]+)\s+Codigo"
    resultado = re.search(padrao_valor_total, texto, re.IGNORECASE)
    if resultado:
        valor_total = resultado.group(1)
        # Corrigir o formato do valor
        valor_total = valor_total.replace('.', '').replace(',', '.')
        try:
            valor_total_formatado = "{:,.2f}".format(float(valor_total)).replace(",", "v").replace(".", ",").replace("v", ".")
            return valor_total_formatado
        except ValueError:
            return None
    return None

def localizar_descricao(texto):
    padrao_descricao = (
        r"LOCAL[:.\s]*(.*?)(?=\s*CNL:|$|OJL:|CN L:|LOC.AL:|LOC.ll.L|LOC. AL|GIJL|LOC.AL S.ANT.*? GIJL)"
    )
    resultado = re.search(padrao_descricao, texto, re.IGNORECASE)
    if resultado:
        descricao = resultado.group(1).strip()
        # Limpar espaços extras e caracteres especiais
        descricao = re.sub(r'\s+', ' ', descricao)
        descricao = re.sub(r'[^\w\s]', '', descricao)
        return descricao[:20]
    return None

def extrair_numero_nota(texto):
    padrao_numero_nota_prefeitura = r"Numero da Nota\s+(\d+)\s+PREFEITURA"
    padrao_numero_nota_data = r"Numero da Nota\s+(\d+)\s+Data"
    
    resultado_prefeitura = re.search(padrao_numero_nota_prefeitura, texto, re.IGNORECASE)
    resultado_data = re.search(padrao_numero_nota_data, texto, re.IGNORECASE)
    
    if resultado_prefeitura:
        return resultado_prefeitura.group(1)
    elif resultado_data:
        return resultado_data.group(1)
    
    return None

def extrair_data_emissao(texto):
    padrao_data = r'\b\d{2}/\d{2}/\d{4}\b'
    resultado = re.search(padrao_data, texto)
    if resultado:
        return resultado.group()
    return None

def main():
    pasta_notas = r'C:\Users\walter.oliveira\Documents\NF_OCR'
    
    try:
        if not os.path.exists(pasta_notas):
            print(f"Erro: A pasta {pasta_notas} não existe.")
            return

        arquivos_pdf = [nome for nome in os.listdir(pasta_notas) if nome.lower().endswith('.pdf')]
        total_arquivos = len(arquivos_pdf)

        if total_arquivos == 0:
            print(f"Nenhum arquivo PDF encontrado na pasta {pasta_notas}.")
            return

        dados = []

        with tqdm(total=total_arquivos, desc="Progresso") as pbar:
            for arquivo in arquivos_pdf:
                caminho_arquivo = os.path.join(pasta_notas, arquivo)
                texto = ler_pdf(caminho_arquivo)
                #print(f"Texto extraído do arquivo {arquivo}:\n{texto}\n")
                valor_total = extrair_valor_total(texto)
                descricao_local = localizar_descricao(texto)
                numero_nota = extrair_numero_nota(texto)
                data_emissao = extrair_data_emissao(texto)
                dados.append({
                    'Arquivo': arquivo,
                    'Numero da Nota': numero_nota,
                    'Data de Emissao': data_emissao,
                    'Valor Total': valor_total,
                    'Descrição Local': descricao_local
                })
                pbar.update(1)

        df = pd.DataFrame(dados)

        nome_arquivo_excel = 'notas_fiscais.xlsx'
        df.to_excel(nome_arquivo_excel, index=False)
        print(f"Dados salvos no arquivo {nome_arquivo_excel}")

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

if __name__ == "__main__":
    main()
