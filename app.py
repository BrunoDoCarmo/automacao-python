import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd

def enviar_email(para, arquivo_anexo, assunto, corpo):
    # Configurações do servidor SMTP
    servidor_smtp = 'mail.infobds.com.br'
    porta_smtp = 587  # Porta padrão para SMTP
    email_enviador = 'bruno@infobds.com.br'
    senha_enviador = '*Ewqiop321'

    # Criando o objeto do e-mail
    msg = MIMEMultipart()
    msg['From'] = email_enviador
    msg['To'] = para
    msg['Subject'] = 'Assunto do e-mail'

    # Anexando o arquivo PDF
    attachment = open(arquivo_anexo, 'rb')

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % arquivo_anexo)

    msg.attach(part)

    # Conectando ao servidor SMTP e enviando e-mail
    servidor = smtplib.SMTP(servidor_smtp, porta_smtp)
    servidor.starttls()
    servidor.login(email_enviador, senha_enviador)
    texto_email = msg.as_string()
    servidor.sendmail(email_enviador, para, texto_email)
    servidor.quit()

    print(f"E-mail enviado para {para}")

# Ler a lista de e-mails da planilha
planilha = pd.read_excel('contatos.xlsx')
emails = planilha['Email']

# Enviar e-mails com o PDF anexado para cada endereço da lista
for email in emails:
    enviar_email(email, 'Anexos\2024\06\Boleto\Bruno.pdf')

print("Todos os e-mails foram enviados com sucesso!")
