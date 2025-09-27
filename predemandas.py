import streamlit as st
import pandas as pd
import os
import io
import zipfile
import unicodedata
from PyPDF2 import PdfMerger, PdfReader
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

# ------------------------
# CONFIG STREAMLIT
# ------------------------
st.set_page_config(page_title="📄 Preparación de Demanda Judicial COS", layout="wide")
st.title("📄 Preparación de Demanda Judicial - COS ⚖️")
st.markdown("""
Esta herramienta permite **unificar documentos de demanda** en un único PDF por cliente, 
organizar los archivos en carpetas y luego **enviar automáticamente las demandas por correo** 
según la base de juzgados.
""")

# ------------------------
# ORDEN MAESTRO
# ------------------------
DOCUMENT_ORDER = {
    "DEMANDA": 1,
    "REMISION DEL PODER": 2,
    "PODER": 3,
    "PAGARE": 4,
    "UBICA": 5,
    "CAMARA Y COMERCIO": 6,
    "SUPERFINANCIERA": 7,
    "SIRNA": 8,
    "MEDIDAS": 9
}

# ------------------------
# SMTP CONFIG (desde secrets)
# ------------------------
SMTP_SERVER = st.secrets["SMTP_SERVER"]
SMTP_PORT = int(st.secrets["SMTP_PORT"])
USER = st.secrets["USER"]
PASSWORD = st.secrets["PASSWORD"]

# Copias fijas
CC_LIST = ["yamile.fonseca@contactosolutions.com"]

# ------------------------
# FUNCIONES
# ------------------------
def limpiar_texto(texto):
    return ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    ).upper()

def detectar_tipo(nombre_archivo: str):
    nombre = limpiar_texto(nombre_archivo)
    if "DEMANDA" in nombre:
        return "DEMANDA"
    elif "REMISION" in nombre and "PODER" in nombre:
        return "REMISION DEL PODER"
    elif "PODER" in nombre and "REMISION" not in nombre:
        return "PODER"
    elif "PAGARE" in nombre:
        return "PAGARE"
    elif "UBICA" in nombre:
        return "UBICA"
    elif "CAMARA_COMERCIO" in nombre or "CAMARA" in nombre or "COMERCIO" in nombre:
        return "CAMARA Y COMERCIO"
    elif "SUPERFINANCIERA" in nombre:
        return "SUPERFINANCIERA"
    elif "SIRNA" in nombre:
        return "SIRNA"
    elif "MEDIDAS" in nombre:
        return "MEDIDAS"
    else:
        return None

def es_cedula(valor):
    return valor.strip().split("_")[0].isdigit()

def desencriptar_pdf(file_bytes):
    reader = PdfReader(file_bytes)
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except:
            pass
    return reader

# ------------------------
# FASE 1: PDFs -> ZIP + Excel
# ------------------------
uploaded_files = st.file_uploader("📂 Sube todos los documentos (PDFs)", type="pdf", accept_multiple_files=True)

clientes = {}
documentos_fijos = {}
df = None

if uploaded_files:
    st.success(f"✅ Se cargaron {len(uploaded_files)} archivos")

    for file in uploaded_files:
        filename = file.name
        if es_cedula(filename):
            parts = filename.split("_")
            cedula = parts[0].strip()
            nombre_cliente = parts[1].strip() if len(parts) > 1 else "SIN_NOMBRE"
            tipo_doc = detectar_tipo(filename)

            if cedula not in clientes:
                clientes[cedula] = {"nombre": nombre_cliente, "docs": {}}
            if tipo_doc:
                clientes[cedula]["docs"][tipo_doc] = file
        else:
            tipo_doc = detectar_tipo(filename)
            if tipo_doc:
                documentos_fijos[tipo_doc] = file

    # Excel de trazabilidad
    data_excel = []
    for cedula, info in clientes.items():
        fila = {"CÉDULA": cedula, "NOMBRE CLIENTE": info["nombre"]}
        for tipo, orden in DOCUMENT_ORDER.items():
            if tipo in info["docs"]:
                fila[tipo] = info["docs"][tipo].name
            elif tipo in documentos_fijos:
                fila[tipo] = documentos_fijos[tipo].name
            else:
                fila[tipo] = "NO SE APORTÓ"
        data_excel.append(fila)

    df = pd.DataFrame(data_excel)
    st.subheader("📊 Vista previa del Excel Global")
    st.dataframe(df)

    # ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for cedula, info in clientes.items():
            carpeta_cliente = f"{cedula}_{info['nombre']}"
            os_path = f"{carpeta_cliente}/"

            for tipo, archivo in info["docs"].items():
                zipf.writestr(os_path + archivo.name, archivo.getvalue())

            for tipo, archivo in documentos_fijos.items():
                zipf.writestr(os_path + archivo.name, archivo.getvalue())

            merger = PdfMerger()
            for tipo, orden in sorted(DOCUMENT_ORDER.items(), key=lambda x: x[1]):
                if tipo in info["docs"]:
                    merger.append(desencriptar_pdf(io.BytesIO(info["docs"][tipo].getvalue())))
                elif tipo in documentos_fijos:
                    merger.append(desencriptar_pdf(io.BytesIO(documentos_fijos[tipo].getvalue())))
            unificado_bytes = io.BytesIO()
            merger.write(unificado_bytes)
            merger.close()
            unificado_bytes.seek(0)
            nombre_unificado = f"{cedula}_{info['nombre']}_DEMANDAUNIFICADA.pdf"
            zipf.writestr(os_path + nombre_unificado, unificado_bytes.read())

    # Descargas
    st.download_button("📥 Descargar Carpeta ZIP con demandas", zip_buffer.getvalue(), "Demandas_Unificadas.zip", "application/zip")
    excel_buffer = io.BytesIO()
    df.to_excel(excel_buffer, index=False)
    st.download_button("📊 Descargar Excel Global", excel_buffer.getvalue(), "Trazabilidad_Demandas.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ------------------------
# FASE 2: ENVÍO REAL
# ------------------------
st.subheader("📧 Enviar demandas (prueba - solo a tu correo corporativo + copia a Yamile)")

base_excel = st.file_uploader("📂 Sube la base de juzgados (Excel)", type=["xlsx"])

if base_excel and df is not None:
    base = pd.read_excel(base_excel)
    base.columns = [c.strip().upper().replace(" ", "_") for c in base.columns]

    if st.button("🚀 Enviar correos"):
        log_envios = []
        for _, row in base.iterrows():
            cedula = str(row["CC_DDO"]).strip()
            nombre = row["NOMBRE_DDO"].strip()
            juzgado = row["JUZGADO"].strip()
            cuantia = row["CUANTÍA"].strip() if "CUANTÍA" in row else row["CUANTIA"].strip()

            asunto = f"RADICACIÓN DEMANDA EJECUTIVA DTE: BANCO GNB SUDAMERIS S.A CONTRA {nombre} CC {cedula}"
            cuerpo = f"""Señor
{juzgado}
{USER}
E. S. D.

REF: EJECUTIVO
DEMANDANTE: BANCO GNB SUDAMERIS S.A.
DEMANDADO: {nombre} CC {cedula}

Cordial saludo.

De manera respetuosa, me dirijo a ustedes con el fin de radicar demanda ejecutiva de {cuantia} cuantía en contra de {nombre}, identificado(a) con cédula de ciudadanía No. {cedula}, con el fin de realizar el correspondiente reparto.

Por lo anterior, solicito acusar recibido y remitir acta de reparto.

Agradeciendo su colaboración,

ADRIANA PAOLA HERNANDEZ ACEVEDO
C.C. No. 1022371176
T. P. No. 248.374 del C. S. de la J"""

            msg = MIMEMultipart()
            msg["From"] = USER
            msg["To"] = USER   # En esta fase de prueba, solo a tu correo
            msg["Cc"] = ", ".join(CC_LIST)
            msg["Subject"] = asunto
            msg.attach(MIMEText(cuerpo, "plain"))

            # Adjuntar PDF unificado
            pdf_name = f"{cedula}_{nombre}_DEMANDAUNIFICADA.pdf"
            if cedula in clientes:
                merger = PdfMerger()
                for tipo, orden in sorted(DOCUMENT_ORDER.items(), key=lambda x: x[1]):
                    if tipo in clientes[cedula]["docs"]:
                        merger.append(desencriptar_pdf(io.BytesIO(clientes[cedula]["docs"][tipo].getvalue())))
                    elif tipo in documentos_fijos:
                        merger.append(desencriptar_pdf(io.BytesIO(documentos_fijos[tipo].getvalue())))
                pdf_bytes = io.BytesIO()
                merger.write(pdf_bytes)
                merger.close()
                pdf_bytes.seek(0)
                part = MIMEApplication(pdf_bytes.read(), Name=pdf_name)
                part["Content-Disposition"] = f'attachment; filename="{pdf_name}"'
                msg.attach(part)

            try:
                with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                    server.starttls()
                    server.login(USER, PASSWORD)
                    recipients = [USER] + CC_LIST
                    server.sendmail(USER, recipients, msg.as_string())
                    estado = "ENVIADO ✅"
            except Exception as e:
                estado = f"ERROR ❌ {str(e)}"

            log_envios.append({
                "CEDULA": cedula,
                "NOMBRE_CLIENTE": nombre,
                "ASUNTO": asunto,
                "ESTADO": estado
            })

        log_df = pd.DataFrame(log_envios)
        st.subheader("📊 Log de envíos")
        st.dataframe(log_df)

        # Descargar log
        log_buffer = io.BytesIO()
        log_df.to_excel(log_buffer, index=False)
        st.download_button("📊 Descargar Log de Envíos", log_buffer.getvalue(), "Log_Envios.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
