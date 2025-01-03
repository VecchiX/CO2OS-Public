import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import ntpath

# SSL証明書が必要となった場合に必要かも
# import os 

def send_email(Conf, Site_Message, fullpaths):

    print("\n\nメール送信処理 Ver.2024.11.13.磯部\n")

    # メール基本情報の設定
    port = Conf['ポート']
    server = Conf['SMTPサーバー']
    sender_email = Conf['送信元アドレス']
    sender_password = Conf['接続パスワード']
    recipient_email = Conf['TO']
    cc_email = Conf['CC']
    subject = Conf['件名']
    body = Site_Message

    if "異常終了" in Site_Message:
        recipient_email = Conf['TO(エラー)']
        cc_email = ""
        subject += "(失敗)"

    # 送信先メールアドレスをリストに変換
    recipient_list = [email.strip() for email in recipient_email.split(",")]
    cc_list = [email.strip() for email in cc_email.split(",")]

    # メールデータの設定
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = ", ".join(recipient_list)
    message["Cc"] = ", ".join(cc_list)
    message["Subject"] = subject

    # メール本文を追加
    message.attach(MIMEText(body, "plain"))

    # ファイルを添付する
    for path in fullpaths:
        with open(path, 'rb') as attachment:
            # MIMEBaseオブジェクトの作成
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())

            # ファイルコンテンツをエンコードする
            encoders.encode_base64(part)
            # ヘッダーを追加
            part.add_header(
                "Content-Disposition",
                "attachment",
                filename=ntpath.basename(path)
            )
            # メールメッセージに添付ファイルを追加
            message.attach(part)            

    # CCのリストを作成
    recipients = recipient_list + cc_list

    # ポート465の場合はSMTPSを使用する(セキュリティ低)
    if port == 465:
        smtp_obj = smtplib.SMTP_SSL(server, port) 

    else:

        # SSLコンテキストを作成
        context = ssl.create_default_context()
        
        # 必要に応じて最も安全なTLSバージョンを明示的に指定
        context.options |= ssl.OP_NO_TLSv1
        context.options |= ssl.OP_NO_TLSv1_1

        # Diffie-Hellman鍵のサイズが小さいサーバーを許容するため、セキュリティレベルを下げる
        context.set_ciphers("DEFAULT:!DH")

        # # SSL証明書を取得（現状は不要）
        # conda_env_path = os.environ.get("CONDA_PREFIX")
        # ca_cert_path = os.path.join(conda_env_path, "Library", "ssl", "cacert.pem")
        # context.load_verify_locations(cafile=ca_cert_path)

        smtp_obj = smtplib.SMTP(server, port)
        smtp_obj.ehlo()
        smtp_obj.starttls(context=context)
        smtp_obj.ehlo()        

    smtp_obj.login(sender_email, sender_password)
    smtp_obj.sendmail(sender_email, recipients, message.as_string())
    smtp_obj.close()
    
    print("メール送信処理 終了")
