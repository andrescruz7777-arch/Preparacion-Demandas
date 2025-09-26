import streamlit as st
import pandas as pd
import os
import io
import zipfile
import unicodedata
from PyPDF2 import PdfMerger

# ------------------------
# CONFIGURACIÓN STREAMLIT
# ------------------------
st.set_page_config(page_title="📄 Preparación de Demanda Judicial COS", layout="wide")
st.title("📄 Preparación de Demanda Judicial - COS ⚖️")
st.markdown("""
Esta herramienta permite **unificar documentos de demanda** en un único PDF por cliente, 
organizar los archivos en carpetas y generar un **Excel global horizontal** con trazabilidad.
""")

# ------------------------
# ORDEN MAESTRO DE DOCUMENTOS
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
# FUNCIONES AUXILIARES
# ------------------------
def limpiar_texto(texto):
    """Quita tildes y pasa a mayúsculas"""
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
    elif "PODER" in nombre:
        return "PODER"
    elif "PAGARE" in nombre:
        return "PAGARE"
    elif "UBICA" in nombre:
        return "UBICA"
    elif ("CAMARA" in nombre or "COMERCIO" in nombre) or ("CERTIFICADO" in nombre and "CAMARA" in nombre):
        return "CAMARA Y COMERCIO"
    elif "SUPERFINANCIERA" in nombre or ("CERTIFICADO" in nombre and "EXISTENCIA" in nombre):
        return "SUPERFINANCIERA"
    elif "SIRNA" in nombre:
        return "SIRNA"
    elif "MEDIDAS" in nombre:
        return "MEDIDAS"
    else:
        return None

# ------------------------
# CARGA DE ARCHIVOS
# ------------------------
uploaded_files = st.file_uploader("📂 Sube todos los documentos (PDFs)", type="pdf", accept_multiple_files=True)

if uploaded_files:
    st.success(f"✅ Se cargaron {len(uploaded_files)} archivos")

    # Agrupar documentos por cédula
    clientes = {}
    for file in uploaded_files:
        filename = file.name
        parts = filename.split("_")
        if len(parts) < 2:
            continue  # ignorar archivos mal nombrados
        cedula = parts[0].strip()
        nombre_cliente = parts[1].strip() if len(parts) > 1 else "SIN_NOMBRE"
        tipo_doc = detectar_tipo(filename)

        if cedula not in clientes:
            clientes[cedula] = {
                "nombre": nombre_cliente,
                "docs": {}
            }
        if tipo_doc:
            clientes[cedula]["docs"][tipo_doc] = file

    # Crear Excel de trazabilidad global
    data_excel = []
    for cedula, info in clientes.items():
        fila = {
            "CÉDULA": cedula,
            "NOMBRE CLIENTE": info["nombre"]
        }
        for tipo, orden in DOCUMENT_ORDER.items():
            if tipo in info["docs"]:
                fila[tipo] = info["docs"][tipo].name
            else:
                fila[tipo] = "NO SE APORTÓ"
        data_excel.append(fila)

    df = pd.DataFrame(data_excel)
    st.subheader("📊 Vista previa del Excel Global")
    st.dataframe(df)

    # Generar ZIP en memoria
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for cedula, info in clientes.items():
            carpeta_cliente = f"{cedula}_{info['nombre']}"
            os_path = f"{carpeta_cliente}/"

            # Guardar documentos individuales
            for tipo, archivo in info["docs"].items():
                zipf.writestr(os_path + archivo.name, archivo.getvalue())

            # Crear unificado
            merger = PdfMerger()
            for tipo, orden in sorted(DOCUMENT_ORDER.items(), key=lambda x: x[1]):
                if tipo in info["docs"]:
                    merger.append(io.BytesIO(info["docs"][tipo].getvalue()))
            unificado_bytes = io.BytesIO()
            merger.write(unificado_bytes)
            merger.close()
            unificado_bytes.seek(0)
            nombre_unificado = f"{cedula}_{info['nombre']}_DEMANDAUNIFICADA.pdf"
            zipf.writestr(os_path + nombre_unificado, unificado_bytes.read())

    # Botón para descargar ZIP
    st.download_button(
        label="📥 Descargar Carpeta ZIP con demandas",
        data=zip_buffer.getvalue(),
        file_name="Demandas_Unificadas.zip",
        mime="application/zip"
    )

    # Botón para descargar Excel
    excel_buffer = io.BytesIO()
    df.to_excel(excel_buffer, index=False)
    st.download_button(
        label="📊 Descargar Excel Global",
        data=excel_buffer.getvalue(),
        file_name="Trazabilidad_Demandas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
