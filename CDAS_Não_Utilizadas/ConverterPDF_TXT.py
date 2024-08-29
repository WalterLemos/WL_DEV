import os
import datetime
import fitz  # PyMuPDF
from docx import Document

# Definindo a pasta de entrada e saída
input_directory = r'C:\ProjetorPython\files\OneDrive_1_29-04-2024\NF_OCR'
output_directory = r'C:\ProjetorPython\files\OCR_TXT'

# Registra o Horário de Início da Execução
Hora_Inicio = datetime.datetime.now()

# Lista todos os arquivos PDF na pasta de entrada
pdf_files = [file for file in os.listdir(input_directory) if file.lower().endswith('.pdf')]

# Iterar para cada arquivo na pasta de entrada
for pdf_file in pdf_files:
    input_pdf_path = os.path.join(input_directory, pdf_file)
    output_docx_path = os.path.join(output_directory, pdf_file.replace('.pdf', '.docx'))

    # Abrir o arquivo PDF
    pdf_document = fitz.open(input_pdf_path)

    # Criar um novo documento Word
    docx_document = Document()

    # Iterar sobre as páginas do PDF
    for pagina_numero in range(len(pdf_document)):
        # Extrair texto da página
        pagina_texto = pdf_document[pagina_numero].get_text()

        # Adicionar o texto da página ao documento Word, mantendo a formatação
        docx_document.add_paragraph(pagina_texto)

    # Salvar o documento Word
    docx_document.save(output_docx_path)

    print(f"PDF converted to DOCX: {output_docx_path}")

    # Fechar o documento PDF
    pdf_document.close()

# Registra o horário do final da Execução
Hora_FIM = datetime.datetime.now()

print(f"Processo iniciado: {Hora_Inicio}. Processo concluído {Hora_FIM}")




