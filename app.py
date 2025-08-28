import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

# === CONFIGURAZIONE ===
EXCEL_FILE = "ELENCO CLIENTI FATTURATI AL 31-12-2021.xls"
WORD_TEMPLATE = "intervento tecnico stufa.docx"

# Carica Excel
df = pd.read_excel(EXCEL_FILE)

# === FUNZIONE PER GENERARE IL DOCUMENTO ===
def genera_documento(nome, cognome, data, tecnico, note):
    cliente = df[(df['Nome'] == nome) & (df['Cognome'] == cognome)]
    if cliente.empty:
        return None

    dati = cliente.iloc[0].to_dict()
    doc = Document(WORD_TEMPLATE)

    for p in doc.paragraphs:
        for key, value in dati.items():
            placeholder = f"{{{{{key.upper()}}}}}"
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, str(value))
        if "{{DATA}}" in p.text:
            p.text = p.text.replace("{{DATA}}", data)
        if "{{TECNICO}}" in p.text:
            p.text = p.text.replace("{{TECNICO}}", tecnico)
        if "{{NOTE}}" in p.text:
            p.text = p.text.replace("{{NOTE}}", note)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# === INTERFACCIA GRAFICA STREAMLIT ===
st.set_page_config(page_title="Gestione Interventi Tecnici", page_icon="🔥", layout="centered")

st.title("📋 Generatore Interventi Tecnici")
st.write("Inserisci i dati e scarica il documento pronto.")

col1, col2 = st.columns(2)
nome = col1.text_input("Nome")
cognome = col2.text_input("Cognome")

data = st.date_input("Data").strftime("%d/%m/%Y")
tecnico = st.text_input("Tecnico")
note = st.text_area("Note")

if st.button("Genera Documento", type="primary"):
    buffer = genera_documento(nome, cognome, data, tecnico, note)
    if buffer:
        st.success("Documento creato con successo!")
        st.download_button(
            label="📥 Scarica Documento",
            data=buffer,
            file_name=f"{nome}_{cognome}_intervento.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.error("Cliente non trovato nell'Excel.")
