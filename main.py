import os
from datetime import datetime
import time
from PyPDF2 import PdfFileReader, PdfFileWriter
import pandas as pd
import re
import smtplib
import sys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

import config

def extract_information(pdf_path):
    current_dir = os.path.abspath(os.path.join(os.path.dirname(__file__)))
    path_processados = os.path.join(current_dir, config.DIRETORIO_RECIBOS_PROCESSADOS)

    data_atual = datetime.now()
    ano_mes = data_atual.strftime('%Y-%m')

    print("**************************************************")
    print("Processamento do arquivo:")
    print(f"{pdf_path}")
    print("**************************************************")

    with open(pdf_path, 'rb') as f:
        pdf = PdfFileReader(f)
        information = pdf.getDocumentInfo()
        number_of_pages = pdf.getNumPages()

        pages = pdf.pages
        for page in pages:
            texto = page.extract_text()
            linhas = texto.splitlines()
            mes_ano = ""

            for linha in linhas:
                if linha.strip().endswith('Código'):
                    codigo = re.findall(r'\d+', linha.strip())
                    codigo = "".join(codigo)

                if linha.strip().endswith('GERAL'):
                    mes_ano = linha.replace("GERAL", "")
                    mes_ano = mes_ano.strip()

            writer = PdfFileWriter()

            page.cropBox.setLowerLeft((0, 450))
            page.cropBox.setUpperRight((595, 842))
            writer.add_page(page)

            path_recibo = os.path.join(path_processados, "{}-recibo-codigo-{}.pdf".format(ano_mes, codigo))
            with open(path_recibo, "wb") as fp:
                writer.write(fp)

            email, nome = email_por_codigo_funci(codigo)

            print(f"Código do funcionário: {codigo}")
            print(f"E-mail: {email}")
            print(f"Nome: {nome}")
            print(f"Caminho do holerite processado: {path_recibo}")

            # Espera 3 segundos para cada holerite
            time.sleep(3)

            if len(email.strip()) and len(nome.strip()):
                subject = "{}".format(config.EMAIL_TITULO).replace('[mes_ano]', mes_ano)
                body = "{}".format(config.EMAIL_BODY).replace('[mes_ano]', mes_ano)
                sender = config.EMAIL_SENDER

                if config.EXECUCAO_TESTE:
                    recipients = config.EXECUCAO_TESTE_EMAILS
                else:
                    recipients = [email.strip()]

                user = config.EMAIL_USER
                password = config.EMAIL_PASSWORD

                send_email(subject, body, sender, recipients, user, password, path_recibo)


def email_por_codigo_funci(codigo_funci):
    email = ''
    nome = ''
    try:
        df = pd.read_excel(config.ARQUIVO_EXCEL_COLABORADORES)
        for i in range(len(df)):
            if (df['Contratação'][i] == 'PJ'):
                continue
            if pd.isna(df['Código da Contabilidade'][i]):
                continue
            if int(df['Código da Contabilidade'][i]) > 0 and int(df['Código da Contabilidade'][i]) == int(codigo_funci):
                email = df['E-mail'][i]
                nome = df['Nome '][i]
    except FileNotFoundError:
        print(f"Arquivo 'colaboradores.xlsx' não encontrado. Verifique o caminho e o nome do arquivo no arquivo 'config.py'")
    except Exception as e:
        print(f"Erro ao ler o arquivo 'colaboradores.xlsx': {str(e)}")

    return email, nome


def send_email(subject, body, sender, recipients, user, password, path_pdf=''):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ', '.join(recipients)
    msg.attach(MIMEText(body, 'plain'))

    if path_pdf != '':
        pdf_name = os.path.basename(path_pdf)
        binary_pdf = open(path_pdf, 'rb')

        payload = MIMEBase('application', 'octate-stream', Name=pdf_name)
        payload.set_payload((binary_pdf).read())

        encoders.encode_base64(payload)
        payload.add_header('Content-Decomposition', 'attachment', filename=pdf_name)
        msg.attach(payload)

    try:
        smtp_server = smtplib.SMTP(config.EMAIL_HOST, 587)
        smtp_server.set_debuglevel(debuglevel=0)
        smtp_server.login(user, password)
        try:
            smtp_server.sendmail(sender, recipients, msg.as_string())
            print("E-mail com o título \"{}\" enviado para \"{}\"".format(subject, ", ".join(recipients)))
            print("**************************************************")
            print()
        finally:
            smtp_server.quit()
    except Exception as e:
        print("Erro: {}".format(str(e)))
        sys.exit("mail failed; %s" % "CUSTOM_ERROR")


if __name__ == '__main__':
    current_dir = os.path.abspath(os.path.join(os.path.dirname(__file__)))
    dir_recibos = os.path.join(current_dir, config.DIRETORIO_RECIBOS)

    if not os.path.exists(dir_recibos):
        os.makedirs(dir_recibos)

    listdir = os.listdir(dir_recibos)
    sorted_list = sorted(listdir)

    for filename in sorted_list:
        if filename.lower().find(".pdf") >= 0:
            path = os.path.join(dir_recibos, filename)
            extract_information(path)

            path_origem = os.path.join(dir_recibos, filename)
            path_destino = os.path.join(current_dir, config.DIRETORIO_RECIBOS_PROCESSADOS, filename)
            os.rename(path_origem, path_destino)

        print("**************************************************")
    print("FIM DO PROCESSAMENTO, ESSA JANELA VAI FECHAR EM 20 SEGUNDOS")
print("**************************************************")
time.sleep(config.EXECUCAO_TEMPO_ESPERA_SEGUNDOS)
