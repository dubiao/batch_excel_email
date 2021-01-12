import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header
import json
import os


def create_msg(to_email, subject, content_html, attach_path, attach_name=None):
    message = MIMEMultipart()
    message['To'] = Header(to_email, 'utf-8')  # 接收者
    message['Subject'] = Header(subject, 'utf-8')
    message.attach(MIMEText(content_html, 'html', 'utf-8'))

    part = MIMEApplication(open(attach_path, 'rb').read())
    if not attach_name:
        attach_name = os.path.split(attach_path)[1]
    part.add_header('Content-Disposition', 'attachment', filename=attach_name)
    message.attach(part)
    return message


class Emailer:
    def __init__(self):
        self.from_email = "Administrator<noreply@what.com>"
        self.reply_to_email = 'Administrator<noreply@what.com>'
        self.smtp = {
            'server': 'smtp.what.what.server.com',
            'login_user': 'noreply@what.com',
            'password': 'PasswordToLoing'
        }
        self.smtp_obj = None

    def login(self):
        try:
            self.smtp_obj = smtplib.SMTP_SSL(self.smtp['server'])
            self.smtp_obj.login(self.smtp['login_user'], self.smtp['password'])
        except smtplib.SMTPException as err:
            print(">>>>>>>   登录失败")
            print(err)
            exit(1)

    def send(self, to_email, subject, content_html, attach_path, attach_name=None):
        # print(to_email)
        # print(subject)
        # print(content_html)
        # return True
        try:
            if not self.smtp_obj:
                self.login()
            message = create_msg(to_email, subject, content_html, attach_path, attach_name)
            message['From'] = Header(self.from_email, 'utf-8')  # 发送者
            if self.reply_to_email:
                message['Reply-to'] = Header(self.reply_to_email, 'utf-8')  #
            send_res = self.smtp_obj.sendmail('hr@qingos.co', to_email, message.as_string())
            return True
        except smtplib.SMTPException as err:
            print(">>>>>>>    %s 邮件发送失败" % to_email)
            print(err)
            exit(1)


if __name__ == '__main__':
    pass