import win32com.client as win32
from pathlib import Path
import codecs
from tqdm import tqdm
import tkinter as tk
from tkinter import filedialog

# Função para substituir caracteres inválidos por underscores
def clean_filename(filename):
    return "".join(c if c.isalnum() or c in ['-', '_'] else '_' for c in filename)

# Função para percorrer as mensagens de uma pasta
def processar_pasta(pasta, destinatario_pesquisa, progress_label):
    mensagens_filtradas = []
    mensagens = pasta.Items
    for m in tqdm(mensagens, desc=f"Processando pasta: {pasta.Name}", leave=False):
        try:
            if destinatario_pesquisa == "todos" or destinatario_pesquisa in m.SenderEmailAddress.lower():
                mensagens_filtradas.append(m)
        except AttributeError:
            pass
        progress_label.config(text=f"Processando pasta: {pasta.Name} - Lidas: {len(mensagens_filtradas)} mensagens")
    return mensagens_filtradas

# Função para iniciar o processamento das mensagens
def processar_mensagens():
    destinatario_pesquisa = dominio_entry.get().lower()

    # Iniciar o Outlook
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Cria a pasta de Destino
    destino = Path.cwd() / "output"
    destino.mkdir(parents=True, exist_ok=True)

    # Acessar as pastas do Outlook
    pastas_raiz = outlook.Folders

    # Atualiza a label de progresso
    progress_label.config(text="Iniciando processo...")

    # Percorrer as pastas e subpastas
    for pasta_raiz in pastas_raiz:
        for pasta in pasta_raiz.Folders:
            mensagens_totais = processar_pasta(pasta, destinatario_pesquisa, progress_label)
            if mensagens_totais:
                for m in tqdm(mensagens_totais, desc=f"Lendo pasta: {pasta.Name}", leave=False):
                    assunto = m.Subject
                    corpo = m.Body
                    anexo = m.Attachments

                    pasta_destino = destino / clean_filename(str(assunto).replace(":", "").replace("/", ""))
                    pasta_destino.mkdir(parents=True, exist_ok=True)

                    progress_label.config(text=f"Lendo pasta: {pasta.Name} - Arquivo: {assunto}")
                    
                    # Escrevendo o conteúdo do corpo do e-mail em UTF-8
                    with codecs.open(pasta_destino / 'Corpo_Email.txt', 'w', 'utf-8') as corpo_file:
                        corpo_file.write(corpo)

                    # Percorre os Anexos
                    for att in anexo:
                        att.SaveAsFile(pasta_destino / clean_filename(str(att)))

    progress_label.config(text="Processo concluído.")

# Criar a interface gráfica
root = tk.Tk()
root.title("Processamento de Emails")

# Configura a janela para abrir maximizada
root.state("zoomed")

# Label e Entry para inserir o domínio ou 'todos'
dominio_label = tk.Label(root, text="Digite o domínio de pesquisa (ou 'todos' para retornar todos os emails):")
dominio_label.pack()
dominio_entry = tk.Entry(root)
dominio_entry.pack()

# Botão para iniciar o processamento
processar_button = tk.Button(root, text="Processar Mensagens", command=processar_mensagens)
processar_button.pack()

# Label para mostrar o progresso
progress_label = tk.Label(root, text="")
progress_label.pack()

root.mainloop()