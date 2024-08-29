import PyPDF2
import re
import openpyxl

def extrair_informacoes_pdf(arquivo_pdf):
    with open(arquivo_pdf, 'rb') as f:
        pdf_reader = PyPDF2.PdfFileReader(f)
        texto_pdf = ""
        for pagina in range(pdf_reader.numPages):
            pagina_pdf = pdf_reader.getPage(pagina)
            texto_pdf += pagina_pdf.extractText()
        return texto_pdf

def localizar_informacoes(texto_pdf):
    local = re.search(r'Local:\s*(.{30})', texto_pdf)
    if local:
        local = local.group(1)
    else:
        local = ''
    valor_total = re.search(r'Valor Total da Nota:\s*(.{15})', texto_pdf)
    if valor_total:
        valor_total = valor_total.group(1)
    else:
        valor_total = ''
    return local, valor_total

def gravar_informacoes_excel(local, valor_total, arquivo_excel):
    wb = openpyxl.load_workbook(arquivo_excel)
    ws = wb.active
    ws.append([local, valor_total])
    wb.save(arquivo_excel)

arquivo_pdf = r'C:\ProjetorPython\files\OneDrive_1_29-04-2024\PR_69699742001710_1000_BU9AL80U_Data-_1_9_2020_Hora-_23_4_16_OCR.pdf'
arquivo_excel = r'C:\ProjetorPython\files\OneDrive_1_29-04-2024\informacoes.xlsx'
texto_pdf = extrair_informacoes_pdf(arquivo_pdf)
local, valor_total = localizar_informacoes(texto_pdf)
gravar_informacoes_excel(local, valor_total, arquivo_excel)