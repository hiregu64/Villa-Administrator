import streamlit as st
import google.generativeai as genai
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import pandas as pd
import io

# ==========================================
# 1. LIVE-VERBINDUNG ZU GOOGLE DRIVE
# ==========================================
def load_data_from_drive():
    try:
        # Holt die JSON-Zugangsdaten aus den Streamlit Secrets
        creds_dict = st.secrets["GOOGLE_CREDENTIALS"]
        creds = service_account.Credentials.from_service_account_info(creds_dict)
        service = build('drive', 'v3', credentials=creds)
        
        # Die ID deiner Excel-Datei aus deinem Google Drive Link
        file_id = '1FzhWZuO6aRZkdRuQBzaojhkq7bQDyprl'
        
        # Datei aus Drive in den Arbeitsspeicher laden
        request = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            
        fh.seek(0)
        # Liest die Excel-Wissensbasis live ein
        df = pd.read_excel(fh)
        return df
    except Exception as e:
        st.error(f"Fehler bei der Verbindung zur Google Drive Wissensbasis: {e}")
        return None

# Daten live laden
df_wissen = load_data_from_drive()

# ==========================================
# 2. DAS NEUE KI-GEHIRN (SYSTEM INSTRUCTION)
# ==========================================
VILLA_PROMPT = """
Du bist „Villa“, der fürsorgliche, hilfreiche Freund und digitale Verwalter für die Bewohner und Helfer der Villa. Deine Aufgabe ist es, den Betrieb und Erhalt des Hauses für alle Benutzer (Besucher, Eigentümer, Administratoren und Handwerker) so einfach wie möglich zu halten.

Technische Architektur & Datenbasis:
Du agierst als die Logik-Ebene einer Streamlit-Web-App. Dir steht im Hintergrund eine Excel-Wissensbasis auf Google Drive zur Verfügung (Villa - Systeme und Ausstattung.xlsx). Du liest Daten aus dieser Tabelle und schreibst Updates über eine Google Service Account API direkt dorthin zurück.

Kommunikationsregeln & Gesprächsführung:
- Erstkontakt / Begrüßung: Antworte extrem kurz, kompakt und smartphone-optimiert. Nenne kurz deine Aufgabe und verweise direkt auf die drei wesentlichen Interaktions-Funktionen („Hilfe“, „Frage“, „Information“). Keine langen Erklärungen vorab.
- Verhalten bei expliziter Bitte um „Hilfe“ (oder Keyword Hilfe): Schalte in den ausführlichen, maximal hilfsbereiten Unterstützungsmodus. Erkläre den Nutzern:
  1. Welche Funktionen sie nutzen können.
  2. Wie du zwischen einer reinen Abfrage (Frage:) und dem Einpflegen neuer Daten (Information:) unterscheidest.
  3. Welche konkreten, sinnvollen Fragen sie stellen können (z. B. „Welche Wartung ist überfällig?“, „Welche Wartungen sind demnächst fällig?“, „Welche Störfälle gab es derletzt (Zeitraum)?“).
- Visuelle Referenzen: Bei Erklärungen zu den Abläufen referenzierst du das Use-Case-Diagramm unter dem Namen Villa Wissen_74.jfif. Bei Fragen zur Wasserversorgung oder zum Systemaufbau referenzierst du die handgezeichnete Skizze des Wasserdrucksystems unter dem Namen PXL_20260516_202437801_74.jpg.

Tonalität: Authentisch, empathisch, geerdet und mit einem Hauch von herzlichem Witz – wie ein verlässlicher Partner vor Ort.
"""

# Gemini API konfigurieren
if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
else:
    st.error("Gemini API Key fehlt in den Secrets!")

# Modell mit dem neuen Gehirn starten
model = genai.GenerativeModel(
    model_name="gemini-1.5-flash",
    system_instruction=VILLA_PROMPT
)

# ==========================================
# 3. BENUTZEROBERFLÄCHE (STREAMLIT APP)
# ==========================================
st.set_page_config(page_title="Villa Verwalter", page_icon="☀️", layout="centered")

st.title("☀️ Villa Wissensbasis")
st.subheader("Dein digitaler Verwalter")

# Benutzer-Auswahl (für das Protokoll)
nutzer = st.selectbox("Wer bist du?", ["Bitte auswählen...", "Anja", "Georgos", "Panajotis", "Fotini", "Handwerker/Helfer"])

# Chat-Verlauf initialisieren
if "messages" not in st.session_state:
    st.session_state.messages = []
    # Kurzer Erstkontakt als Standard-Begrüßung durch die KI
    st.session_state.messages.append({
        "role": "assistant", 
        "content": "Hallo! Ich bin „Villa“ – dein digitaler Verwalter. ☀️\nIch helfe dir, den Überblick zu behalten. Nutze einfach dein Tastatur-Mikrofon.\n\nVerwende am Anfang:\n* **`Hilfe`** (Anleitung öffnen)\n* **`Frage: [Deine Frage]`** (Suchen)\n* **`Information: [Dein Update]`** (Eintragen)"
    })

# Chat-Verlauf anzeigen
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# Chat-Eingabe (Smartphone-optimiert)
if prompt := st.chat_input("Wie kann ich helfen?"):
    if nutzer == "Bitte auswählen...":
        st.warning("Bitte wähle oben zuerst deinen Namen aus, bevor du eine Nachricht sendest!")
    else:
        # Nutzer-Nachricht anzeigen
        with st.chat_message("user"):
            st.markdown(prompt)
        st.session_state.messages.append({"role": "user", "content": prompt})
        
        # Kontext aus der Excel-Tabelle für die KI aufbereiten
        kontext = ""
        if df_wissen is not None:
            kontext = f"\n\nAktuelle Daten aus der Wissensbasis:\n{df_wissen.to_string(index=False)}"
        
        # Antwort von Gemini generieren lassen
        with st.chat_message("assistant"):
            try:
                response = model.generate_content(prompt + kontext)
                st.markdown(response.text)
                st.session_state.messages.append({"role": "assistant", "content": response.text})
            except Exception as e:
                st.error(f"Fehler bei der KI-Verarbeitung: {e}")
