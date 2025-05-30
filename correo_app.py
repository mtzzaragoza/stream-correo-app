import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
import mimetypes
import os
import tempfile

st.set_page_config(page_title="Correo personalizado", layout="centered")
st.title("📧 Envío de correos personalizados")

# Ingreso del remitente
remitente = st.text_input("✉️ Correo remitente (Gmail)")

clave_app = None
if remitente:
    claves = st.secrets.get("CUENTAS", {})
    clave_app = claves.get(remitente)

    if clave_app:
        st.success("✅ Clave de aplicación cargada automáticamente.")
    else:
        st.warning("⚠️ Este correo no está registrado en secrets. No se podrá enviar.")

# Subida del archivo Excel
archivo = st.file_uploader("📁 Carga el archivo Excel con los datos", type=["xlsx"])

if archivo and remitente and clave_app:
    df = pd.read_excel(archivo)
    st.success("✔️ Archivo cargado correctamente.")
    st.dataframe(df.head())

    if st.button("📨 Enviar correos"):
        enviados = 0
        errores = []

        for _, fila in df.iterrows():
            try:
                destinatario = fila.get("Email")
                asunto = fila.get("Asunto", "Sin asunto")
                mensaje_intro = str(fila.get("Mensaje_intro", ""))
                mensaje_detalle = str(fila.get("Mensaje_detalle", ""))
                cuerpo = f"{mensaje_intro}\n\n{mensaje_detalle}".strip()

                if pd.isna(destinatario) or not destinatario:
                    continue

                msg = EmailMessage()
                msg["Subject"] = str(asunto)
                msg["From"] = remitente
                msg["To"] = destinatario
                msg.set_content(cuerpo)

                # Adjuntar archivos de columnas que empiecen con "Archivo"
                for col in fila.index:
                    if col.lower().startswith("archivo") and pd.notna(fila[col]):
                        ruta_archivo = fila[col]
                        if os.path.isfile(ruta_archivo):
                            tipo, _ = mimetypes.guess_type(ruta_archivo)
                            maintype, subtype = tipo.split("/", 1) if tipo else ("application", "octet-stream")

                            with open(ruta_archivo, "rb") as f:
                                msg.add_attachment(
                                    f.read(),
                                    maintype=maintype,
                                    subtype=subtype,
                                    filename=os.path.basename(ruta_archivo)
                                )

                with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                    smtp.login(remitente, clave_app)
                    smtp.send_message(msg)
                    enviados += 1

            except Exception as e:
                errores.append((fila.get("Email"), str(e)))

        st.success(f"✅ Correos enviados: {enviados}")
        if errores:
            st.error("❌ Algunos correos no se enviaron:")
            for email, err in errores:
                st.text(f"{email}: {err}")
else:
    if not archivo:
        st.info("🔄 Carga un archivo Excel para comenzar.")
    if remitente and not clave_app:
        st.error("❌ No se encontró una clave para el remitente ingresado.")

