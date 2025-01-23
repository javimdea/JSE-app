from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import os
import pandas as pd
import smtplib
import zipfile
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from pypdf import PdfReader, PdfWriter

app = Flask(__name__)
app.secret_key = "supersecretkey"

# Configuración de carpetas
UPLOAD_FOLDER = "static/uploads"
QR_FOLDER = "QR"  # Aseguramos que los QR se guarden directamente en la carpeta QR
EXCEL_FOLDER = "static/excel"
PDF_FOLDER = "pdf_invitaciones"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(QR_FOLDER, exist_ok=True)
os.makedirs(EXCEL_FOLDER, exist_ok=True)
os.makedirs(PDF_FOLDER, exist_ok=True)

# Variables globales
ultimo_excel_path = None
pdf_template_path = None

# Configuración del correo
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_SENDER = "jmdea2013@gmail.com"  # Reemplaza con tu correo
EMAIL_PASSWORD = "eddb lxvc hlbs clbw"  # Usa la contraseña de aplicación generada

@app.route("/", methods=["GET", "POST"])
def index():
    return render_template("index.html", excel_qr_generado=bool(ultimo_excel_path))

@app.route("/subir_excel_sin_qr", methods=["POST"])
def subir_excel_sin_qr():
    global ultimo_excel_path
    if "excel_sin_qr" not in request.files:
        flash("Debes subir un archivo Excel.", "danger")
        return redirect(url_for("index"))

    excel_file = request.files["excel_sin_qr"]
    if excel_file.filename == "":
        flash("Selecciona un archivo antes de subir.", "danger")
        return redirect(url_for("index"))

    excel_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
    excel_file.save(excel_path)
    ultimo_excel_path = excel_path

    flash("Excel subido correctamente. Ahora sube la carpeta con los QR.", "success")
    return redirect(url_for("index"))

@app.route("/subir_qr_folder", methods=["POST"])
def subir_qr_folder():
    global ultimo_excel_path
    if not ultimo_excel_path:
        flash("Primero sube el Excel sin QR.", "danger")
        return redirect(url_for("index"))

    if "qr_zip" not in request.files:
        flash("Debes subir un archivo ZIP con los QR.", "danger")
        return redirect(url_for("index"))

    qr_zip_file = request.files["qr_zip"]
    qr_zip_path = os.path.join(UPLOAD_FOLDER, qr_zip_file.filename)
    qr_zip_file.save(qr_zip_path)

    with zipfile.ZipFile(qr_zip_path, "r") as zip_ref:
        zip_ref.extractall(QR_FOLDER)

    df = pd.read_excel(ultimo_excel_path, engine="openpyxl")

    # **Se guarda la ruta relativa en la columna QR**
    df["QR"] = df["Nombre"].apply(lambda nombre: os.path.join("QR", "QR", f"{nombre.replace(' ', '_')}.png"))

    excel_con_qr_path = os.path.join(EXCEL_FOLDER, "Lista_Invitados_Con_QR.xlsx")
    df.to_excel(excel_con_qr_path, index=False)

    flash("QR asignados correctamente. Puedes descargar el nuevo Excel.", "success")
    return redirect(url_for("index"))

@app.route("/descargar_excel")
def descargar_excel():
    excel_con_qr_path = os.path.join(EXCEL_FOLDER, "Lista_Invitados_Con_QR.xlsx")
    if os.path.exists(excel_con_qr_path):
        return send_file(excel_con_qr_path, as_attachment=True)
    flash("El archivo no está disponible.", "danger")
    return redirect(url_for("index"))

@app.route("/subir_excel_y_pdf", methods=["POST"])
def subir_excel_y_pdf():
    global pdf_template_path
    if "excel" not in request.files or "pdf" not in request.files:
        flash("Debes subir un Excel y un PDF.", "danger")
        return redirect(url_for("index"))

    excel_file = request.files["excel"]
    pdf_file = request.files["pdf"]

    excel_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
    pdf_path = os.path.join(UPLOAD_FOLDER, pdf_file.filename)

    excel_file.save(excel_path)
    pdf_file.save(pdf_path)
    pdf_template_path = pdf_path

    flash("Archivos subidos correctamente. Enviando correos...", "success")
    enviar_correos(excel_path, pdf_template_path)

    return redirect(url_for("index"))

def generar_pdf(nombre, qr_path, output_pdf_path):
    reader = PdfReader(pdf_template_path)
    writer = PdfWriter()
    first_page = reader.pages[0]

    canvas_temp = f"{PDF_FOLDER}/{nombre.replace(' ', '_')}_temp.pdf"
    c = canvas.Canvas(canvas_temp, pagesize=letter)
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(300, 150, nombre)
    c.save()

    name_layer = PdfReader(canvas_temp)
    first_page.merge_page(name_layer.pages[0])
    writer.add_page(first_page)

    for page in reader.pages[1:-1]:
        writer.add_page(page)

    last_page = reader.pages[-1]
    canvas_qr_temp = f"{PDF_FOLDER}/{nombre.replace(' ', '_')}_qr.pdf"
    c = canvas.Canvas(canvas_qr_temp, pagesize=letter)
    c.drawImage(qr_path, 130, 275, width=150, height=150)
    c.save()

    qr_layer = PdfReader(canvas_qr_temp)
    last_page.merge_page(qr_layer.pages[0])
    writer.add_page(last_page)

    with open(output_pdf_path, "wb") as output_file:
        writer.write(output_file)

    os.remove(canvas_temp)
    os.remove(canvas_qr_temp)

def enviar_correos(excel_path, pdf_template_path):
    df = pd.read_excel(excel_path, engine="openpyxl")

    for index, row in df.iterrows():
        nombre = str(row["Nombre"]).strip()
        correo = str(row["Correo"]).strip()
        qr_path = str(row["QR"]).strip()
        pdf_output_path = f"{PDF_FOLDER}/INVITACION_{nombre.replace(' ', '_')}.pdf"

        generar_pdf(nombre, qr_path, pdf_output_path)

        msg = MIMEMultipart()
        msg["From"] = EMAIL_SENDER
        msg["To"] = correo
        msg["Subject"] = "Invitación Recepción Juan Sebastián de Elcano"

        body = f"""
        <p><strong>El Excmo. Sr. Contralmirante,</strong><br>
        Comandante del Mando Naval de Canarias<br>
        D. Santiago de Colsa Trueba</p>
        <p><strong>El Ilmo. Sr. Capitán de Navío,</strong><br>
        Comandante del Buque Escuela “Juan Sebastián de Elcano”<br>
        D. Luis Carreras-Presas do Campo</p>
        <p><strong>Tienen el honor de invitar a {nombre} a la recepción que,</strong><br>
        tendrá lugar a bordo el martes día 21 de enero de 2025.</p>

        <p><strong>Su código QR:</strong></p>
        <img src="cid:qr_image" width="150"/>

        <p><strong>Emblema del evento:</strong></p>
        <img src="cid:emblema_image" width="120"/>
        """
        msg.attach(MIMEText(body, "html"))
                # Ajustar la ruta del QR para cada persona (porque están en QR/QR)
        qr_path = os.path.join("QR", "QR", f"{nombre.replace(' ', '_')}.png")
        
        if os.path.exists(qr_path):
            with open(qr_path, "rb") as qr_file:
                qr_img = MIMEImage(qr_file.read())
                qr_img.add_header("Content-ID", "<qr_image>")  # Para embebido en el HTML
                msg.attach(qr_img)
        else:
            print(f"[ERROR] No se encontró el QR para {nombre} en {qr_path}")

        # Ruta de la imagen del emblema
        emblema_path = os.path.join(UPLOAD_FOLDER, "IMAGEN.png")

        if os.path.exists(emblema_path):
            with open(emblema_path, "rb") as img_file:
                img = MIMEImage(img_file.read())
                img.add_header("Content-ID", "<emblema_image>")  # Para embebido en el HTML
                msg.attach(img)
        else:
            print("[ERROR] No se encontró IMAGEN.png en static/uploads.")


        with open(pdf_output_path, "rb") as pdf_file:
            part = MIMEBase("application", "pdf")
            part.set_payload(pdf_file.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f'attachment; filename="INVITACION_{nombre.replace(" ", "_")}.pdf"')
            msg.attach(part)

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        server.sendmail(EMAIL_SENDER, correo, msg.as_string())
        server.quit()

if __name__ == "__main__":
    app.run(debug=True)






