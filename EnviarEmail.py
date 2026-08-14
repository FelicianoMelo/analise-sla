import json
import smtplib
from email.message import EmailMessage
from pathlib import Path
from string import Template


# ==========================
# CONFIGURAÇÕES SMTP
# ==========================
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

USUARIO = "feliciano.junior.osv@fedex.com"
SENHA = "9616Rywe#243"

def carregar_json(caminho_json):
    with open(caminho_json, "r", encoding="utf-8") as arquivo:
        return json.load(arquivo)


def substituir_variaveis(texto, variaveis):
    return Template(texto).safe_substitute(variaveis)


def montar_email(dados):
    msg = EmailMessage()

    assunto = substituir_variaveis(
        dados["assunto"],
        dados.get("variaveis", {})
    )

    corpo = substituir_variaveis(
        dados["mensagem"],
        dados.get("variaveis", {})
    )

    msg["Subject"] = assunto
    msg["From"] = USUARIO
    msg["To"] = ", ".join(dados.get("email_para", []))

    if dados.get("email_cc"):
        msg["Cc"] = ", ".join(dados["email_cc"])

    msg.set_content(corpo)

    # Anexos
    for arquivo in dados.get("anexo", []):
        caminho = Path(arquivo)

        if caminho.exists():
            with open(caminho, "rb") as f:
                msg.add_attachment(
                    f.read(),
                    maintype="application",
                    subtype="octet-stream",
                    filename=caminho.name
                )
        else:
            print(f"Anexo não encontrado: {arquivo}")

    return msg


def enviar_email(msg):
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login(USUARIO, SENHA)
        smtp.send_message(msg)

    print("E-mail enviado com sucesso.")


def main():
    dados = carregar_json("dados.json")
    email = montar_email(dados)
    enviar_email(email)


if __name__ == "__main__":
    main()