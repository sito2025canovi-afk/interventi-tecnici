# Generatore Interventi Tecnici

Questa è una web app Streamlit per generare documenti Word di intervento tecnico.

## Come funziona
1. L'app legge il file Excel (`ELENCO CLIENTI FATTURATI AL 31-12-2021.xls`) con i dati dei clienti.
2. Usa un modello Word (`intervento tecnico stufa.docx`) con i segnaposto (es. `{{NOME}}`, `{{COGNOME}}`, `{{DATA}}`).
3. Inserisci Nome, Cognome, Data, Tecnico, Note → scarichi il documento compilato.

## Deploy su Streamlit Cloud
1. Carica questi file in un repository GitHub (anche privato):  
   - `app.py`  
   - `requirements.txt`  
   - `ELENCO CLIENTI FATTURATI AL 31-12-2021.xls`  
   - `intervento tecnico stufa.docx`  
2. Vai su https://share.streamlit.io → **New app** → collega il repo → scegli `app.py`.
3. Ottieni il link per accedere alla tua app.
