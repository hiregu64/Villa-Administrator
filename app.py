import streamlit as st
import google.generativeai as genai
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import pandas as pd
import io
import datetime

# Google File ID der Excel-Tabelle
FILE_ID = '1FzhWZuO6aRZkdRuQBzaojhkq7bQDyprl'

# ==========================================
# 1. LIVE-DATEN AUS GOOGLE DRIVE LESEN
# ==========================================
@st.cache_data(ttl=60)  # Cache für 1 Minute, um API-Limits zu schonen
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
        st.cache_data.clear()  # Cache leeren, um frische Daten zu erzwingen
        return True
    except Exception as e:
        st.error(f"Fehler beim Schreiben in Google Drive: {e}")
        return False

# ==========================================
# 3. KI-GEHIRN INITIALISIERUNG
# ==========================================
VILLA_PROMPT = """
Du bist „Villa“, der fürsorgliche, hilfreiche Freund und digitale Verwalter für die Bewohner und Helfer der Villa. Deine Aufgabe ist es, den Betrieb und Erhalt des Hauses für alle Benutzer (Besucher, Eigentümer, Administratoren und Handwerker) so einfach wie möglich zu halten.

Falls Hilfe angefordert wird: Erkläre die Funktionen und verweise bezüglich der Abläufe auf das Use-Case-Diagramm „Villa Wissen_72.jfif“.
Falls Fragen zur Wasserversorgung aufkommen, verweise direkt auf die Skizze des Wasserdrucksystems „PXL_20260516_202437801_72.jpg“.
"""

if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
model = genai.GenerativeModel(model_name="gemini-1.5-flash", system_instruction=VILLA_PROMPT)

# ==========================================
# 4. BENUTZEROBERFLÄCHE (STREAMLIT UI)
# ==========================================
st.set_page_config(page_title="Villa Verwalter", page_icon="☀️", layout="centered")
st.title("☀️ Villa Wissensbasis")

# Rollen-Auswahl
nutzer_rolle = st.selectbox("Wer bist du?", ["Bitte auswählen...", "Besucher", "Eigentümer", "Administrator", "Handwerker/Helfer"])

if "messages" not in st.session_state:
    st.session_state.messages = [{
        "role": "assistant", 
        "content": "Hallo! Ich bin „Villa“ – dein digitaler Verwalter. ☀️ Ich helfe dir, den Betrieb des Hauses so einfach wie möglich zu halten. Nutze gerne dein Tastatur-Mikrofon!\n\nVerwende die Buttons oder tippe direkt los mit den Begriffen **`Hilfe`**, **`Frage:`** oder **`Information:`**."
    }]

# Chat-Verlauf anzeigen
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# Dynamische Steuerung basierend auf der Rolle
if nutzer_rolle != "Bitte auswählen...":
    st.write("---")
    st.subheader("Schnellauswahl: Was möchtest du tun?")
    
    # Grid für die Konzept-Buttons
    col1, col2, col3 = st.columns(3)
    col4, col5 = st.columns(2)
    
    aktion = None
    with col1:
        if st.button("ℹ️ Ich brauche Hilfe."): aktion = "Hilfe"
    with col2:
        if st.button("⚠️ Es gibt eine Störung."): aktion = "Störung"
    with col3:
        if st.button("📊 Ich benötige einen Bericht."): aktion = "Bericht"
    with col4:
        if st.button("📝 Ich habe neue Informationen."): aktion = "Information"
    with col5:
        if st.button("🛠️ Ich möchte eine Änderung am XLS vornehmen."): aktion = "Änderung"

    # Dropdown-Verständnishilfe nach Konzept
    kategorie_auswahl = st.selectbox(
        "Bereich einschränken (optional zur Verbesserung des Eingabeverständnisses):",
        ["Keine Einschränkung", "Geräte / Ausst. innen", "Geräte / Ausst. außen", "Systeme"]
    )

    # Logikverarbeitung der Konzept-Buttons
    if aktion:
        prompt_text = ""
        if aktion == "Hilfe":
            prompt_text = "Hilfe"
        elif aktion == "Störung":
            prompt_text = f"Frage: Es gibt eine Störung im Bereich '{kategorie_auswahl}'. Was sollte ich jetzt tun?"
        elif aktion == "Bericht":
            prompt_text = f"Frage: Gib mir einen Bericht zum Thema '{kategorie_auswahl}' für die letzte Zeit."
        elif aktion == "Information":
            prompt_text = f"Information: [Bitte hier deine neuen Infos eintragen] (Bereich: {kategorie_auswahl})"
        elif aktion == "Änderung":
            prompt_text = f"Information: Ich möchte eine Änderung am XLS vornehmen im Bereich '{kategorie_auswahl}': "
        
        st.info(f"Vorschlag generiert! Kopiere diesen Text oder nutze ihn als Vorlage für das Eingabefeld unten: \n\n**`{prompt_text}`**")

# Manueller Chat-Input (Smartphone-optimiert)
if prompt := st.chat_input("Wie kann ich helfen? (z.B. 'Frage: Wann war die letzte Wartung?')"):
    if nutzer_rolle == "Bitte auswählen...":
        st.warning("Bitte wähle oben zuerst aus, wer du bist!")
    else:
        with st.chat_message("user"):
            st.markdown(prompt)
        st.session_state.messages.append({"role": "user", "content": prompt})
        
        # Falls es sich um einen Schreibbefehl handelt (Information: )
        if prompt.strip().lower().startswith("information:") and df_wissen is not None:
            reiner_text = prompt.split(":", 1)[1].strip()
            kat = kategorie_auswahl if 'kategorie_auswahl' in locals() else "Allgemein"
            erfolg = append_info_to_drive(df_wissen, reiner_text, nutzer_rolle, kat)
            if erfolg:
                st.success("Eintrag erfolgreich in der Google Drive Excel-Wissensbasis gespeichert!")
                df_wissen, _ = load_data_from_drive()

        # KI-Antwort generieren mit Live-Tabellenkontext
        kontext = f"\n\nAktuelle Daten aus der Wissensbasis:\n{df_wissen.to_string(index=False)}" if df_wissen is not None else ""
        with st.chat_message("assistant"):
            try:
                response = model.generate_content(f"Nutzer-Rolle: {nutzer_rolle}\nGewählter Bereich: {kategorie_auswahl if 'kategorie_auswahl' in locals() else 'Keiner'}\nAnfrage: {prompt} {kontext}")
                st.markdown(response.text)
                st.session_state.messages.append({"role": "assistant", "content": response.text})
            except Exception as e:
                st.error(f"Fehler bei der KI-Verarbeitung: {e}")
