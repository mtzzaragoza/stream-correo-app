import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
import mimetypes

# Título y configuración
st.set_page_config(page_title="Correo personalizado", layout="centered")
st.title("📧 Envío de correos personalizados")

# Entrada del remitente
remitente = st.text_input("Ingresa tu correo Gmail", placeholder="tucorreo@gmail.com")

# Subir archivo Excel
archivo_excel = st.file_uploader("Carga el archivo Excel con los datos", type=["xlsx"])

# Subir archivos adjuntos
archivos_subidos = st.file_uploader("Sube los archivos adjuntos", accept_multiple_files=True)

# Diccionario para acceder a archivos por nombre
archivos_dict = {}
if archivos_subidos:
    for archivo in archivos_subidos:
        archivos_dict[archivo.name] = archivo

# Obtener la clave desde secrets
clave_app = None
if remitente:
    clave_app = st.secrets.get("credenciales", {}).get(remitente)

if remitente and not clave_app:
    st.warning("⚠️ Este correo no está registrado en secrets. No se podrá enviar.")

if archivo_excel and clave_app:
    df = pd.read_excel(archivo_excel)
    st.success("✅ Archivo Excel cargado correctamente.")
    st.dataframe(df.head())

    if st.button("📨 Enviar correos"):
        enviados = 0
        errores = []

        for _, fila in df.iterrows():
            try:
                destinatario = fila.get("Email")
                asunto = str(fila.get("Asunto", "Sin asunto"))
                mensaje_intro = str(fila.get("Mensaje_intro", ""))
                mensaje_detalle = str(fila.get("Mensaje_detalle", ""))
                cuerpo = f"{mensaje_intro}\n\n{mensaje_detalle}".strip()

                if pd.isna(destinatario) or not destinatario:
                    continue

                msg = EmailMessage()
                msg["Subject"] = asunto
                msg["From"] = remitente
                msg["To"] = destinatario
                msg.set_content(cuerpo)

                # Adjuntar archivos según las columnas que empiecen con "Archivo"
                for col in fila.index:
                    if col.lower().startswith("archivo"):
                        nombre_archivo = str(fila[col]).strip()
                        if pd.notna(nombre_archivo) and nombre_archivo in archivos_dict:
                            archivo_obj = archivos_dict[nombre_archivo]
                            tipo, _ = mimetypes.guess_type(archivo_obj.name)
                            maintype, subtype = tipo.split("/", 1) if tipo else ("application", "octet-stream")

                            msg.add_attachment(
                                archivo_obj.read(),
                                maintype=maintype,
                                subtype=subtype,
                                filename=archivo_obj.name
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


