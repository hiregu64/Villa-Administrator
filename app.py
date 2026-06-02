import streamlit as st
from google import genai
from google.genai import types
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import pandas as pd
import io
import datetime
import json

# Google File ID der Excel-Tabelle
FILE_ID = '1FzhWZuO6aRZkdRuQBzaojhkq7bQDyprl'

# ==========================================
# 1. LIVE-DATEN AUS GOOGLE DRIVE LESEN
# ==========================================
@st.cache_data(ttl=30)  
def load_data_from_drive():
    try:
        creds_dict = st.secrets["GOOGLE_CREDENTIALS"]
        creds = service_account.Credentials.from_service_account_info(creds_dict)
        service = build('drive', 'v3', credentials=creds)
        
        request = service.files().get_media(fileId=FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            
        fh.seek(0)
        return pd.read_excel(fh), service
    except Exception as e:
        st.error(f"Fehler bei der Verbindung zur Google Drive Wissensbasis: {e}")
        return None, None

# Visueller Lade-Spinner beim Abruf der Daten
with st.spinner("Verbindung zur Google Drive Wissensbasis wird hergestellt..."):
    df_wissen, drive_service = load_data_from_drive()

# ==========================================
# 2. LIVE-UPDATE IN GOOGLE DRIVE SCHREIBEN
# ==========================================
def append_info_to_drive(df, neuer_text, nutzername, kategorie="Nicht definiert"):
    try:
        neue_zeile = {
            "Zeitstempel": datetime.datetime.now().strftime("%d.%m.%Y %H:%M"),
            "Nutzer": nutzername,
            "Kategorie": kategorie,
            "Eintrag / Update": neuer_text
        }
        
        df_aktualisiert = pd.concat([df, pd.DataFrame([neue_zeile])], ignore_index=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_aktualisiert.to_excel(writer, index=False)
        output.seek(0)
        
        media = MediaIoBaseUpload(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        drive_service.files().update(fileId=FILE_ID, media_body=media).execute()
        st.cache_data.clear()  
        return True
    except Exception as e:
        st.error(f"Fehler beim Schreiben in Google Drive: {e}")
        return False

# ==========================================
# 3. KI-GEHIRN (OFFIZIELLES GOOGLE-GENAI SDK)
# ==========================================
VILLA_PROMPT = """
Du bist „Villa Avatar“, der digitale Helfer für die Bewohner und Helfer der Villa. Deine Aufgabe ist es, den Betrieb und Erhalt des Hauses so einfach wie möglich zu halten.
Beziehe dich bei allgemeinen Abläufen auf 'Villa Wissen_72.jfif' und bei der Wasserversorgung auf 'PXL_20260516_202437801_72.jpg'.

WICHTIGER KONTEXT & VERHALTEN:
- Antworte immer kurz, präzise und smartphone-optimiert.
- Nutze die vom HMI übergebene Rolle und die gewählte Kategorie/Bezeichnung zwingend als Arbeitsgrundlage.
- Wenn das HMI dir eine konkrete Bezeichnung (z. B. "Beregnungssystem") übergibt, beziehe dich exakt darauf.
- Wenn das HMI KEINE konkrete Bezeichnung übergibt (weil der Nutzer das Drop-down ignoriert hat), musst du anhand der Frage des Nutzers logisch nachdenken und den Kontext selbstständig der passenden Kategorie (Systeme, Geräte innen, Geräte außen) zuordnen.
"""

@st.cache_resource
def get_ki_client():
    if "GEMINI_API_KEY" in st.secrets:
        return genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
    return None

client = get_ki_client()

def generate_ki_response(prompt_text):
    if client is None:
        return "KI-Dienst nicht konfiguriert (API Key fehlt in den Secrets)."
    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt_text,
            config=types.GenerateContentConfig(
                system_instruction=VILLA_PROMPT
            )
        )
        return response.text
    except Exception as e:
        try:
            response = client.models.generate_content(
                model="gemini-1.5-flash",
                contents=prompt_text,
                config=types.GenerateContentConfig(
                    system_instruction=VILLA_PROMPT
                )
            )
            return response.text
        except Exception:
            return f"Fehler bei der KI-Verarbeitung: {e}"

# ==========================================
# 4. BENUTZEROBERFLÄCHE (HMI)
# ==========================================
# Issue 17: Titel im Browsertab
st.set_page_config(page_title="Villa Avatar", page_icon="
