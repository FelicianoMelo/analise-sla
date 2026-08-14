import pandas as pd
import smtplib
import os

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Configurações SMTP
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

USUARIO = "feliciano.junior.osv@fedex.com"
SENHA = "9616Rywe#243"

# Arquivo Excel
ARQUIVO_XLSX = "emails.xlsx"

# Carrega planilha
df = pd.read_excel(ARQUIVO_XLSX)

for index, row in df.iterrows():

    destinatario = row["email"]
    assunto = row["assunto"]
    mensagem = row["mensagem"]
    anexo_path = row["anexo"]

    try:
        # Monta email
        email = MIMEMultipart()
        email["From"] = USUARIO
        email["To"] = destinatario
        email["Subject"] = assunto

        email.attach(MIMEText(mensagem, "plain", "utf-8"))

        # Anexo
        if pd.notna(anexo_path) and os.path.exists(anexo_path):

            with open(anexo_path, "rb") as arquivo:
                parte = MIMEBase("application", "octet-stream")
                parte.set_payload(arquivo.read())

            encoders.encode_base64(parte)

            parte.add_header(
                "Content-Disposition",
                f'attachment; filename="{os.path.basename(anexo_path)}"'
            )

            email.attach(parte)

        # Conexão SMTP
        servidor = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        servidor.starttls()
        servidor.login(USUARIO, SENHA)

        servidor.sendmail(
            USUARIO,
            destinatario,
            email.as_string()
        )

        servidor.quit()

        print(f"Email enviado para: {destinatario}")

    except Exception as e:
        print(f"Erro ao enviar para {destinatario}: {e}")