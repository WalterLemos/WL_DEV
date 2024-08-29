import win32com.client as win32
from pathlib import Path
from tqdm import tqdm
import codecs
import openpyxl
import tkinter as tk
from tkinter import ttk, filedialog

# Função para substituir caracteres inválidos por underscores
def clean_filename(filename):
    return "".join(c if c.isalnum() or c in ['-', '_'] else '_' for c in filename)

# Função para processar as mensagens com base no destinatário fornecido e no arquivo Excel selecionado
def processar_mensagens(pasta, destinatario_pesquisa, progress_bar):
    # Cria a pasta de Destino
    destino = Path.cwd() / "output"
    destino.mkdir(parents=True, exist_ok=True)

    # Iniciar o Outlook
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Lista para armazenar mensagens filtradas
    mensagens_filtradas = []

    # Função recursiva para percorrer pastas e subpastas
    def percorrer_pastas(pasta):
        for m in pasta.Items:
            try:
                sender_email = m.SenderEmailAddress.lower()
                if destinatario_pesquisa in sender_email:
                    mensagens_filtradas.append(m)
            except AttributeError:
                pass

        for subpasta in pasta.Folders:
            percorrer_pastas(subpasta)

    # Chama a função recursiva para a pasta raiz
    percorrer_pastas(pasta)

    # Se não houver e-mails com o domínio, mostrar uma mensagem
    if not mensagens_filtradas:
        resultado_label.config(text="Nenhum e-mail correspondente ao destinatário encontrado.")
    else:
        resultado_label.config(text=f"Processando {len(mensagens_filtradas)} e-mails...")

        # Criar a barra de progresso
        progress_bar.config(max=len(mensagens_filtradas), value=0)
        progress_bar.pack()
        # Percorre as Mensagens filtradas com barra de progresso
        for idx, m in enumerate(tqdm(mensagens_filtradas, desc="Processando mensagens")):
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

# Cria a interface
root = tk.Tk()
root.title("Filtro de E-mails")

# Label e Entry para digitar o destinatário de pesquisa
destinatario_label = tk.Label(root, text="Digite o destinatário de pesquisa:")
destinatario_label.pack()
destinatario_entry = tk.Entry(root)
destinatario_entry.pack()

# Botão para processar as mensagens
def processar_button_click():
    destinatario_pesquisa = destinatario_entry.get().lower()

    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    root_folder = outlook.GetDefaultFolder(6)

    progress_bar = ttk.Progressbar(root, orient="horizontal", mode="determinate")
    processar_mensagens(root_folder, destinatario_pesquisa, progress_bar)


processar_button = tk.Button(root, text="Processar E-mails", command=processar_button_click)
processar_button.pack()

# Label para mostrar o resultado do processamento
resultado_label = tk.Label(root, text="")
resultado_label.pack()

root.mainloop()