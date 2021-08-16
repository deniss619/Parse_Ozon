import os
import smtplib
from email import encoders
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from platform import python_version


class Email:
    def __init__(self, server, user, password, recipients, subject, text):
        self.server = server
        self.user = user
        self.password = password
        self.recipients = recipients
        self.subject = subject
        self.text = text
        self.msg = MIMEMultipart('alternative')

    def generate_mail(self):
        self.msg['Subject'] = self.subject
        self.msg['From'] = 'Python script <' + self.user + '>'
        self.msg['To'] = ', '.join(self.recipients)
        self.msg['Reply-To'] = self.user
        self.msg['Return-Path'] = self.user
        self.msg['X-Mailer'] = 'Python/' + (python_version())
        part_text = MIMEText(self.text, 'plain')
        self.msg.attach(part_text)

    def add_file(self, filepath):
        part_file = MIMEApplication(open(filepath, "rb").read(), Name=os.path.basename(filepath), _subtype="xlsx")
        encoders.encode_base64(part_file)
        self.msg.attach(part_file)

    def add_html(self):
        html = '<html><head></head><body><p>' + self.text + '</p></body></html>'
        part_html = MIMEText(html, 'html')
        self.msg.attach(part_html)

    def send_mail(self):
        mail = smtplib.SMTP_SSL(self.server)
        mail.login(self.user, self.password)
        mail.sendmail(self.user, self.recipients, self.msg.as_string())
        mail.quit()
