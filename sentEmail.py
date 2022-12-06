import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


def mail_sandler():
    subject = "Вам письмо"
    body = """Привет моя дорогая, Маша!
    \nЯ тебя очень жду в Актау!\nСкучаю и схожу с ума по тебе и без тебя!
    """
    sender_email = "babakhinSV@gmail.com"
    receiver_email = "Kristall_68@inbox.ru"
    password = input("Type your password and press enter: ")

    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message["Bcc"] = receiver_email  # Recommended for mass emails

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    filename = "love.jpg"  # In same directory as script

    try:
        # Open PDF file in binary mode
        with open(filename, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        # Encode file in ASCII characters to send by email
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )
        # Add attachment to message and convert message to string
        message.attach(part)
    except Exception as e:
        print(e)
    full_msg = message.as_string()
    # Log in to server using secure context and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        #
        # server.start(context=context)
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, full_msg)
        server.quit()
    print("Message was successfully sent")


if __name__ == "__main__":
    mail_sandler()
