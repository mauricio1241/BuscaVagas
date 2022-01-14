import requests
import xlsxwriter
import smtplib
import ssl

from requests.exceptions import HTTPError

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def buscaVagas():
    try:
        r = requests.get(
            'https://apiteams.goobee.com.br/api/publicavaga/vagas/SITE_CADMUS')

        r.raise_for_status()

        jsonResponse = r.json()

        return jsonResponse

    except HTTPError as http_err:
        print(f'HTTP error occurred: {http_err}')
        return False
    except Exception as err:
        print(f'Other error occurred: {err}')
        return False


def montarPlanilha(jsonResponse):
    try:
        workbook = xlsxwriter.Workbook('Vagas.xlsx')
        worksheet = workbook.add_worksheet()
        bold = workbook.add_format({'bold': True})

        worksheet.write('A1', "Nome da vaga", bold)
        worksheet.write('B1', "Local", bold)
        worksheet.write('C1', "Descrição", bold)

        worksheet.set_column(0, 0, 50)
        worksheet.set_column(1, 1, 25)

        row = 0
        col = 0

        row += 1
        for vaga in jsonResponse:
            worksheet.write(row, col, vaga["name"])
            worksheet.write(row, col + 1, vaga["cidade_Regi_o__c"])
            worksheet.write(row, col + 2, vaga["descricao_da_vaga__c"])
            row += 1

        workbook.close()

        return True
    except Exception as err:
        print(f'Other error occurred: {err}')
        return False


def enviarEmail():
    try:
        subject = "Listagem de vagas"
        body = "Olá, esse é um e-mail enviado automaticamente com a listagem das vagas existentes no site."
        sender_email = "sendedr@gmail.com"
        receiver_email = "receiver@gmail.com"
        password = input("Informe a senha do e-mail:")

        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = subject
        message["Bcc"] = receiver_email

        message.attach(MIMEText(body, "plain"))

        filename = "Vagas.xlsx"

        with open(filename, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        encoders.encode_base64(part)

        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )

        message.attach(part)
        text = message.as_string()

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, text)

        return True
    except Exception as err:
        print(f'Other error occurred: {err}')
        return False
