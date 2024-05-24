import os
import datetime
import fitz  # PyMuPDF
from PIL import Image
from google.cloud import vision
from docx import Document

def extract_text_from_image(image_path):
    client = vision.ImageAnnotatorClient()
    with open(image_path, 'rb') as image_file:
        content = image_file.read()
    image = vision.Image(content=content)
    response = client.text_detection(image=image)
    texts = response.text_annotations
    if response.error.message:
        raise Exception(response.error.message)
    if texts:
        return texts[0].description
    return ""

# Definindo a pasta de entrada e saída
input_directory = r'C:\ProjetorPython\files\NF'
output_directory = r'C:\ProjetorPython\files\OCR_TXT'

# Certifique-se de que o diretório de saída exista
os.makedirs(output_directory, exist_ok=True)

# Registra o Horário de Início da Execução
hora_inicio = datetime.datetime.now()

# Lista todos os arquivos PDF na pasta de entrada
pdf_files = [file for file in os.listdir(input_directory) if file.lower().endswith('.pdf')]

# Iterar para cada arquivo na pasta de entrada
for pdf_file in pdf_files:
    input_pdf_path = os.path.join(input_directory, pdf_file)
    output_docx_path = os.path.join(output_directory, pdf_file.replace('.pdf', '.docx'))

    try:
        # Abrir o arquivo PDF
        pdf_document = fitz.open(input_pdf_path)

        # Criar um novo documento Word
        docx_document = Document()

        # Iterar sobre as páginas do PDF
        for page_number in range(len(pdf_document)):
            # Extrair a imagem da página
            page = pdf_document.load_page(page_number)
            pix = page.get_pixmap()
            image_path = os.path.join(output_directory, f"page_{page_number}.png")
            pix.save(image_path)

            # Realizar OCR na imagem usando Google Cloud Vision
            ocr_text = extract_text_from_image(image_path)
            
            # Adicionar o texto OCR ao documento Word
            docx_document.add_paragraph(ocr_text)
            
            # Remover a imagem temporária
            os.remove(image_path)

        # Salvar o documento Word
        docx_document.save(output_docx_path)

        print(f"PDF convertido para DOCX com OCR: {output_docx_path}")

        # Fechar o documento PDF
        pdf_document.close()
    except Exception as e:
        print(f"Falha ao processar {pdf_file}: {e}")

# Registra o horário do final da Execução
hora_fim = datetime.datetime.now()

print(f"Processo iniciado: {hora_inicio}. Processo concluído: {hora_fim}")


