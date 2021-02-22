
import ssl, smtplib, csv
from tabulate import tabulate
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


stmp_server= 'stmp.gmail.com'
port = 465


sender = 'rybki.for.training@gmail.com'
#sender = input ("Enter your email")
password = input ("Enter your password")

context = ssl.create_default_context()
receiver_email = 'arkadiusz.rybski@gmail.com'
message = """\
Subject: Hi there

This message is sent from Python."""

text = """
Hi, 
According to the agreement, You will be part of the 24/7 Support Team during net week.
Please find the information about your team.

{table}

Best regards,

Admin"""

html = """
<html><body><p>Hi, </p>
<p>According to the agreement, You will be part of the 24/7 Support Team during net week.</p>
<p>Please find the information about your team.</p>
{table}
<p>Best regards,</p>
<p>Admin</p>
</body></html>
"""


with open('./data/example.csv') as input_file:
    reader = csv.reader(input_file)
    data = list(reader)

text = text.format(table=tabulate(data, headers="firstrow", tablefmt="grid"))
html = html.format(table=tabulate(data, headers="firstrow", tablefmt="html"))

message = MIMEMultipart(
    "alternative", None, [MIMEText(text), MIMEText(html,'html')])


with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
    server.login(sender,  password)
    server.sendmail(sender, receiver_email, message.as_string())



