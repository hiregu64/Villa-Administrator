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
# 3. KI-GEHIRN INITIALISIERUNG
# ==========================================
VILLA_PROMPT = """
Du bist „Villa“, der digitale Verwalter für die Bewohner und Helfer der Villa. Deine Aufgabe ist es, den Betrieb und Erhalt des Hauses so einfach wie möglich zu halten.
Beziehe dich bei allgemeinen Abläufen auf 'Villa Wissen_72.jfif' und bei der Wasserversorgung auf 'PXL_20260516_202437801_72.jpg'.
"""

if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

try:
    model = genai.GenerativeModel(
        model_name="gemini-1.5-flash",
        system_instruction=VILLA_PROMPT
    )
except Exception:
    model = genai.GenerativeModel(
        model_name="models/gemini-1.5-flash",
        system_instruction=VILLA_PROMPT
    )

def generate_ki_response(prompt_text):
    try:
        return model.generate_content(prompt_text).text
    except Exception as e:
        try:
            res = genai.generate_text(
                model="models/gemini-1.5-flash",
                prompt=f"{VILLA_PROMPT}\n\nAnfrage:\n{prompt_text}"
            )
            return res.result
        except Exception:
            raise e

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
        "content": "Hallo! Ich bin „Villa Barsinghausen“ – dein digitaler Verwalter. ☀️ Nutze gerne dein Tastatur-Mikrofon!\n\nWähle oben deine Rolle aus, um zu beginnen."
    }]

# Chat-Verlauf anzeigen
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# Variablen initialisieren
button_prompt = None
gewaehlte_aktion = "Allgemein"

if nutzer_rolle != "Bitte auswählen...":
    st.write("---")
    
    # 1. SCHRITT: NUTZUNG DER WISSENSBASIS FÜR (BUTTONS)
    st.subheader("Nutzung der Wissensbasis für:")
    
    if nutzer_rolle == "Besucher":
        col1, col2 = st.columns(2)
        with col1:
            if st.button("ℹ️ Ich brauche Hilfe."):
                button_prompt = "Was möchtest du wissen?"
                gewaehlte_aktion = "Hilfe"
        with col2:
            if st.button("⚠️ Es gibt eine Störung."):
                button_prompt = "Was ist passiert?"
                gewaehlte_aktion = "Störung"
    else:
        col1, col2, col3 = st.columns(3)
        col4, col5 = st.columns(2)
        with col1:
            if st.button("ℹ️ Ich brauche Hilfe."):
                button_prompt = "Was möchtest du wissen?"
                gewaehlte_aktion = "Hilfe"
        with col2:
            if st.button("⚠️ Es gibt eine Störung."):
                button_prompt = "Was ist passiert?"
                gewaehlte_aktion = "Störung"
        with col3:
            if st.button("📊 Ich benötige einen Bericht."):
                button_prompt = "Nenne mir bitte den Zeitraum und das Thema."
                gewaehlte_aktion = "Bericht"
        with col4:
            if st.button("📝 Ich habe neue Informationen."):
                button_prompt = "Gern nehme ich deine Informationen auf und ordne sie in meiner Wissensbasis zu."
                gewaehlte_aktion = "Information"
        with col5:
            if st.button("🛠️ Ich möchte eine Änderung am XLS vornehmen."):
                button_prompt = "Beschreibe deine Änderung so genau wie möglich."
                gewaehlte_aktion = "Änderung"

    # 2. SCHRITT: DROP-DOWN AUSWAHL ZUR UNTERSTÜTZUNG (DARUNTER)
    st.write("")
    kategorie_auswahl = st.selectbox(
        "Verständnishilfe (Auswahl filtert die Liste der Bezeichnungen):",
        ["Alle Einträge", "Geräte / Ausst. innen", "Geräte / Ausst. außen", "Systeme"]
    )
    
    # 3. SCHRITT: ANZEIGE DER BEZEICHNUNGEN
    if df_wissen is not None and not df_wissen.empty:
        spalten_namen = df_wissen.columns.tolist()
        bez_spalte = spalten_namen[1] if len(spalten_namen) > 1 else spalten_namen[0]
        
        if kategorie_auswahl != "Alle Einträge":
            suchbegriff = kategorie_auswahl.replace(" ", "").lower()
            mask = df_wissen.astype(str).apply(lambda x: x.str.replace(" ", "").str.lower().str.contains(suchbegriff)).any(axis=1)
            df_gefiltert = df_wissen[mask]
        else:
            df_gefiltert = df_wissen
            
        if not df_gefiltert.empty:
            anzeige_df = df_gefiltert[[bez_spalte]].drop_duplicates()
            anzeige_df = anzeige_df[anzeige_df[bez_spalte].astype(str).str.lower() != str(bez_spalte).lower()]
            anzeige_df = anzeige_df.reset_index(drop=True)
            
            st.dataframe(anzeige_df, use_container_width=True, hide_index=True)
        else:
            st.info(f"Keine Einträge für '{kategorie_auswahl}' hinterlegt.")

    # Verarbeitung der Klicks
    if button_prompt:
        st.session_state.messages.append({"role": "user", "content": f"Aktion gewählt: {gewaehlte_aktion} (Bereich: {kategorie_auswahl})"})
        
        if gewaehlte_aktion in ["Information", "Änderung"] and df_wissen is not None:
            append_info_to_drive(df_wissen, f"Button-Aktion: {gewaehlte_aktion}", nutzer_rolle, kategorie_auswahl)
        
        kontext = f"\n\nAktuelle Daten aus der Wissensbasis:\n{df_wissen.to_string(index=False)}" if df_wissen is not None else ""
        try:
            antwort_text = generate_ki_response(
                f"SYSTEM-BEFEHL: Der Nutzer hat den Button für '{gewaehlte_aktion}' gedrückt. "
                f"Antworte ihm exakt mit der Spezifikations-Gegenfrage: '{button_prompt}'. "
                f"Gib keine weiteren Erklärungen ab, sondern warte auf seine Eingabe zum Bereich '{kategorie_auswahl}'. {kontext}"
            )
            st.session_state.messages.append({"role": "assistant", "content": antwort_text})
            st.rerun()
        except Exception as e:
            st.error(f"Fehler bei der Verarbeitung: {e}")

# Manueller Chat-Input
if prompt := st.chat_input("Wie kann ich helfen? (z.B. 'Frage: Wo ist der Hauptwasserhahn?')"):
    if nutzer_rolle == "Bitte auswählen...":
        st.warning("Bitte wähle oben zuerst aus, wer du bist!")
    else:
        with st.chat_message("user"):
            st.markdown(prompt)
        st.session_state.messages.append({"role": "user", "content": prompt})
        
        if prompt.strip().lower().startswith("information:") and df_wissen is not None:
            reiner_text = prompt.split(":", 1)[1].strip()
            kat = kategorie_auswahl if ('kategorie_auswahl' in locals() and kategorie_auswahl != "Alle Einträge") else "Allgemein"
            erfolg = append_info_to_drive(df_wissen, reiner_text, nutzer_rolle, kat)
            if erfolg:
                st.success("Eintrag erfolgreich in Google Drive gespeichert!")
                st.cache_data.clear()
                df_wissen, _ = load_data_from_drive()

        # KI-Antwort generieren via stabiler Wrapper-Funktion
        kontext = f"\n\nAktuelle Daten aus der Wissensbasis:\n{df_wissen.to_string(index=False)}" if df_wissen is not None else ""
        with st.chat_message("assistant"):
            try:
                antwort_text = generate_ki_response(
                    f"SYSTEM-KONTEXT: Der Nutzer tippt in der Rolle '{nutzer_rolle}'. "
                    f"Ausgewählter Bereich im HMI: '{kategorie_auswahl if 'kategorie_auswahl' in locals() else 'Alle Einträge'}'.\n"
                    f"Anfrage: {prompt} {kontext}"
                )
                st.markdown(antwort_text)
                st.session_state.messages.append({"role": "assistant", "content": antwort_text})
            except Exception as e:
                st.error(f"Fehler bei der KI-Verarbeitung: {e}")
