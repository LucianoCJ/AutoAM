import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Establece los detalles del correo electrónico y el servidor SMTP
from_email = "qboqbo@gmail.com"
from_password = "zrnwijlqdljlcxof"
to_email = "qboqbo@gmail.com"
smtp_server = "smtp.gmail.com"
smtp_port = 587

# Construye el mensaje
msg = MIMEMultipart()
msg['From'] = from_email
msg['To'] = to_email
msg['Subject'] = "Archivo adjunto"

# Agrega el cuerpo del mensaje
body = "Este es un mensaje de prueba"
msg.attach(MIMEText(body, 'plain'))

# Agrega el archivo adjunto
filename = "OAG.pdf"
with open(filename, "rb") as f:
    attach = MIMEApplication(f.read(), _subtype="pdf")
    attach.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(attach)

# Envía el correo electrónico
try:
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(from_email, from_password)
    server.sendmail(from_email, to_email, msg.as_string())
    server.quit()
    print("Correo enviado exitosamente!")
except Exception as e:
    print("Error al enviar correo electrónico: ", e)
