import win32com.client as win32
from pathlib import Path
from tqdm import tqdm
import codecs  # Importando a biblioteca codecs

# Função para substituir caracteres inválidos por underscores
def clean_filename(filename):
    return "".join(c if c.isalnum() or c in ['-', '_'] else '_' for c in filename)

# Cria a pasta de Destino
destino = Path.cwd() / "output"
destino.mkdir(parents=True, exist_ok=True)

# Iniciar o Outlook
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Acessando as Pastas
pastas = outlook.Folders.Item(1)

for pasta in pastas.Folders:
    print(pasta.Name)

# Caixa de Entrada / Mensagens
inbox = outlook.GetDefaultFolder(6)
mensagens = inbox.items

# Percorre as Mensagens com barra de progresso
for m in tqdm(mensagens, desc="Processando mensagens"):
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





