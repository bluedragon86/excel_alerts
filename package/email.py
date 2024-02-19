import smtplib
import html
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import make_msgid

# email.sendemail(names, certificates, duedate)

# Função para envio de email
#def send_message(names, certificates, duedates):
def send_message(html_table, filename):
    #print(duedates)
     
    # Set your email credentials 
    sender_email = "no_replay@arsopi.pt"
    subject = "Alertas - Excel: "+filename

    # Recipient email address
    receiver_email = "joao.silva@arsopi.pt"

    # Faz quebra de linha no corpo do email
    #email_body = '<br>'.join(message_to_send)

    #print(email_body)
    # print(send_to_1300)
    # print(send_to_1200)
    
    # Create the MIME object
    message = MIMEMultipart()
    message['From'] = sender_email
    message['Subject'] = subject
    message['To'] = receiver_email


    # # Convert list to HTML table
    # table_html = "<table border='1'>"
    # table_html += "<tr><th>#</th><th>NOMES</th><th>CERTIFICADO</th><th>VALIDADE CERTIFICADO</th></tr>"
    # for index, (name, certificate, duedate) in enumerate(zip(names, certificates, duedates), start=1):
    #     table_html += f"<tr><td>{index}</td><td>{html.escape(name)}</td><td>{html.escape(certificate)}</td><td>{html.escape(str(duedate))}</td></tr>"
    # table_html += "</table>"
    
    # print(table_html)

    # Attach the body to the email
    message.attach(MIMEText(html_table, _subtype ='html', _charset = 'utf-8'))
    

    # # Set up the SMTP server (for Gmail, use port 587)
    smtp_server = "arsopi-pt.mail.protection.outlook.com"
    smtp_port = 25

     # Start the SMTP server session
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()  # Use TLS

        # Send the email
        server.sendmail(sender_email, receiver_email, message.as_string())
