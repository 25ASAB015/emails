
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd

# Cargar datos
contacts = pd.read_excel("Base directorio 13052025.xlsx", sheet_name="Hoja1", engine="openpyxl")
contacts = contacts[contacts['Correo Electronico'].notna()]

# Configuración del servidor SMTP de Gmail
smtp_server = "smtp.gmail.com"
smtp_port = 587
sender_email = "TU_CORREO@gmail.com"
sender_password = "TU_CONTRASEÑA"

# Conexión al servidor
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(sender_email, sender_password)

# Envío de correos
for index, row in contacts.iterrows():
    destinatario = row['Correo Electronico']
    nombre = row['Contacto Directorio']
    empresa = row['Nombre Comercial']

    subject = "Presentación Comercial"
    body = f"""
    Estimado/a {nombre},

    Me complace presentarle nuestra empresa y explorar oportunidades de colaboración con {empresa}. 
    Estamos interesados en establecer relaciones comerciales con organizaciones líderes como la suya.

    Quedamos atentos a cualquier consulta o propuesta que desee compartir.

    Atentamente,
    [Tu Nombre]
    [Tu Empresa]
    [Tu Teléfono]
    [Tu Correo Electrónico]
    """

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = destinatario
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        server.sendmail(sender_email, destinatario, msg.as_string())
        print(f"Correo enviado a {destinatario}")
    except Exception as e:
        print(f"Error al enviar a {destinatario}: {e}")

# Cerrar conexión
server.quit()
