import imaplib
import email
import os
import zipfile
from pyunpack import Archive
from datetime import datetime

# Configurações do e-mail
EMAIL = "novvacontabilidade@gmail.com"
SENHA = "hlos pfrd jzwp dntd"
IMAP_SERVER = "imap.gmail.com"

# Obtém o caminho da Área de Trabalho do usuário
desktop_path = os.path.join(os.path.expanduser("~"), "Área de Trabalho")  # Windows em PT-BR
if not os.path.exists(desktop_path):  # Se não existir, tentar "Desktop" (caso o idioma do SO seja diferente)
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

# Criar a pasta "fiscal_2025" na Área de Trabalho
fiscal_path = os.path.join(desktop_path, "fiscal_2025")
os.makedirs(fiscal_path, exist_ok=True)  # Garante que a pasta "fiscal_2025" exista

# Obter o ano e o mês atual
current_year = datetime.now().year
current_month = datetime.now().strftime("%m")

# Conectar ao servidor IMAP
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(EMAIL, SENHA)
mail.select("inbox")

# Buscar todos os e-mails
status, messages = mail.search(None, 'ALL')
emails = messages[0].split()

for num in emails:
    status, msg_data = mail.fetch(num, "(RFC822)")
    tem_anexo = False  # Variável para verificar se há anexo no e-mail

    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            remetente = msg["From"].split()[-1].strip("<>")  # Extrai o e-mail do remetente

            # Verificar anexos
            for part in msg.walk():
                if part.get_content_maintype() == "multipart":
                    continue
                if part.get("Content-Disposition") is None:
                    continue
                
                filename = part.get_filename()
                if filename:
                    tem_anexo = True  # Marca que há anexo

                    # Criar a estrutura de pastas somente se houver anexo
                    remetente_pasta = os.path.join(fiscal_path, remetente)
                    year_path = os.path.join(remetente_pasta, str(current_year))
                    month_path = os.path.join(year_path, current_month)
                    os.makedirs(month_path, exist_ok=True)  # Garante que as pastas do ano e mês existam

                    filepath = os.path.join(month_path, filename)
                    with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    print(f"Anexo {filename} salvo em {month_path}")

                    # Verificar se é um arquivo .rar e converter para .zip
                    if filename.endswith(".rar"):
                        # Caminho do arquivo .rar
                        rar_filepath = os.path.join(month_path, filename)

                        # Caminho do arquivo .zip
                        zip_filepath = rar_filepath.replace(".rar", ".zip")

                        # Descompactar e criar arquivo .zip
                        Archive(rar_filepath).extractall(month_path)
                        with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_DEFLATED) as zipf:
                            for foldername, subfolders, filenames in os.walk(os.path.splitext(rar_filepath)[0]):
                                for file in filenames:
                                    zipf.write(os.path.join(foldername, file), os.path.relpath(os.path.join(foldername, file), month_path))

                        # Excluir o arquivo .rar após conversão
                        os.remove(rar_filepath)
                        print(f"Arquivo .rar convertido para .zip: {zip_filepath}")

    # Se o e-mail não tinha anexo, ele é ignorado
    if not tem_anexo:
        print(f"E-mail de {msg['From']} ignorado (sem anexo).")

mail.logout()
