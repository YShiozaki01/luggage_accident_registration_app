import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import os


class SendMail:
    def __init__(self, **kwargs):
        # SMTPサーバ情報
        self.smtp_server = kwargs["smtp_server"]
        self.port_number = kwargs["port_number"]
        # ログイン情報
        self.account = kwargs["my_account"]
        self.password = kwargs["password"]
        # 件名、メールアドレス、添付ファイル
        self.my_address = kwargs["my_account"]
        self.to_address = kwargs["to_address"]
        self.subject = kwargs["subject"]
        self.attach_zip = kwargs["attach_zip"]

    def send_mail(self):
        # メールを送信する
        # 1)メッセージの準備
        msg = MIMEMultipart()
        # 2)件名、メールアドレスの設定
        msg["Subject"] = self.subject
        msg["From"] = self.my_address
        msg["To"] = self.to_address
        # 3)メール本文の追加
        text = open("mail_body.txt", encoding="utf-8")
        body_text = text.read()
        text.close()
        body = MIMEText(body_text)
        msg.attach(body)
        # 4)添付ファイル（zipファイル）の追加
        filename = os.path.basename(self.attach_zip)
        with open(self.attach_zip, "br") as f:
            zip_file = f.read()
        attach_file = MIMEApplication(zip_file)
        attach_file.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(attach_file)
        # 5)SMTPサーバに接続
        server = smtplib.SMTP(self.smtp_server, self.port_number)   # SMTPサーバの指定
        res_server = server.noop()  # SMTPサーバーの応答確認
        print(res_server)   # 応答表示
        res_starttls = server.starttls()    # 暗号化通信の開始
        print(res_starttls) # 応答表示
        res_login = server.login(self.account, self.password) # ログイン
        print(res_login)    # 応答表示
        # 6)メール送信
        server.send_message(msg)
        # 7)SMTPサーバとの接続解除
        server.quit()
