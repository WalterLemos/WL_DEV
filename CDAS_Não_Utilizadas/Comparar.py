import win32com.client as win32
from pathlib import Path
from tqdm import tqdm
import codecs
import openpyxl
import tkinter as tk
from tkinter import filedialog

# Função para substituir caracteres inválidos por underscores
def clean_filename(filename):
    return "".join(c if c.isalnum() or c in ['-', '_'] else '_' for c in filename)

# Função para processar as mensagens com base no destinatário fornecido e no arquivo Excel selecionado
def processar_mensagens():
    # Obter o destinatário de pesquisa inserido pelo usuário
    destinatario_pesquisa = destinatario_entry.get().lower()

    # Obter o caminho do arquivo Excel selecionado
    arquivo_excel = arquivo_excel_path.get()

    # Cria a pasta de Destino
    destino = Path.cwd() / "output"
    destino.mkdir(parents=True, exist_ok=True)

    # Carregar a planilha Excel
    workbook = openpyxl.load_workbook(arquivo_excel)
    planilha = workbook.active

    # Iniciar o Outlook
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Acessando a Caixa de Entrada
    inbox = outlook.GetDefaultFolder(6)
    mensagens = inbox.items

    # Lista para armazenar mensagens filtradas
    mensagens_filtradas = []

    # Percorre as Mensagens e filtra pelo domínio do remetente
    for m in mensagens:
        try:
            sender_email = m.SenderEmailAddress.lower()
            if destinatario_pesquisa in sender_email:
                mensagens_filtradas.append(m)
        except AttributeError:
            pass

    # Se não houver e-mails com o domínio, mostrar uma mensagem
    if not mensagens_filtradas:
        resultado_label.config(text="Nenhum e-mail correspondente ao destinatário encontrado.")
    else:
        resultado_label.config(text=f"Processando {len(mensagens_filtradas)} e-mails...")
        # Percorre as Mensagens filtradas com barra de progresso
        for m in tqdm(mensagens_filtradas, desc="Processando mensagens"):
            assunto = m.Subject
            corpo = m.body
            anexo = m.Attachments

            # Modificando o nome da pasta_destino usando a função clean_filename
            pasta_destino = destino / clean_filename(str(assunto))
            pasta_destino.mkdir(parents=True, exist_ok=True)

            # Escrevendo o conteúdo do corpo do e-mail em UTF-8
            with codecs.open(pasta_destino / 'Corpo_Email.txt', 'w', 'utf-8') as corpo_file:
                corpo_file.write(corpo)

            # Percorre os Anexos
            for att in tqdm(anexo, desc="Salvando anexos", leave=False):
                att.SaveAsFile(pasta_destino / clean_filename(str(att)))

        resultado_label.config(text="E-mails processados com sucesso!")
        workbook.close()

# Cria a interface
root = tk.Tk()
root.title("Filtro de E-mails")

# Label e Entry para digitar o destinatário de pesquisa
destinatario_label = tk.Label(root, text="Digite o destinatário de pesquisa:")
destinatario_label.pack()
destinatario_entry = tk.Entry(root)
destinatario_entry.pack()

# Botão para selecionar o arquivo Excel
arquivo_excel_path = tk.StringVar()

def selecionar_arquivo_excel():
    arquivo_excel_path.set(filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx")]))
    arquivo_excel_label.config(text=f"Arquivo Excel selecionado: {arquivo_excel_path.get()}")

selecionar_arquivo_excel_button = tk.Button(root, text="Selecionar Arquivo Excel", command=selecionar_arquivo_excel)
selecionar_arquivo_excel_button.pack()
arquivo_excel_label = tk.Label(root, text="Arquivo Excel não selecionado")
arquivo_excel_label.pack()

# Botão para processar as mensagens
processar_button = tk.Button(root, text="Processar E-mails", command=processar_mensagens)
processar_button.pack()

# Label para mostrar o resultado do processamento
resultado_label = tk.Label(root, text="")
resultado_label.pack()

root.mainloop()