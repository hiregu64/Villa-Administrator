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

WICHTIG FÜR JEDE ANTWORT:
Dir wird bei jeder Anfrage die ausgewählte Rolle und der aktuell ausgewählte Bereich (Kategorie/System) des Nutzers übergeben. Nutze diese Information zwingend, um kurze, präzise und kontextbezogene Antworten zu geben. Wenn der Bereich 'Systeme' gewählt ist und der Nutzer eine vage Frage stellt, bezieht sich dies immer auf die in den Systemen hinterlegten Daten.
"""

if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
model = genai.GenerativeModel(model_name="gemini-1.5-flash", system_instruction=VILLA_PROMPT)

# ==========================================
# 4. BENUTZEROBERFLÄCHE (STREAMLIT UI)
# ==========================================
st.set_page_config(page_title="Villa Verwalter", page_icon="☀️", layout="centered")
st.title("☀️ Villa Wissensbasis")

# Rollen-Auswahl (Ganz oben)
nutzer_rolle = st.selectbox("Wer bist du?", ["Bitte auswählen...", "Besucher", "Eigentümer", "Administrator", "Handwerker/Helfer"])

if "messages" not in st.session_state:
    st.session_state.messages = [{
        "role": "assistant", 
        "content": "Hallo! Ich bin „Villa“ – dein digitaler Verwalter. ☀️ Nutze gerne dein Tastatur-Mikrofon!\n\nWähle oben deine Rolle aus, um zu beginnen."
    }]

# Chat-Verlauf anzeigen
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# Variable für Button-Klicks initialisieren
button_prompt = None
gewaehlte_aktion = "Allgemein"

# Wenn die Rolle ausgewählt ist, folgt die Struktur exakt der PPT-Reihenfolge von oben nach unten
if nutzer_rolle != "Bitte auswählen...":
    st.write("---")
    
    # 1. SCHRITT: NUTZUNG DER WISSENSBASIS FÜR (BUTTONS)
    st.subheader("Nutzung der Wissensbasis für:")
    
    if nutzer_rolle == "Besucher":
        col1, col2 = st.columns(2)
        with col1:
            if st.button("ℹ️ Ich brauche Hilfe."): 
                button_prompt = "Hilfe"
                gewaehlte_aktion = "Hilfe"
        with col2:
            if st.button("⚠️ Es gibt eine Störung."): 
                button_prompt = "Ich möchte eine Störung melden."
                gewaehlte_aktion = "Störung"
    else:
        col1, col2, col3 = st.columns(3)
        col4, col5 = st.columns(2)
        with col1:
            if st.button("ℹ️ Ich brauche Hilfe."): 
                button_prompt = "Hilfe"
                gewaehlte_aktion = "Hilfe"
        with col2:
            if st.button("⚠️ Es gibt eine Störung."): 
                button_prompt = "Ich möchte eine Störung melden."
                gewaehlte_aktion = "Störung"
        with col3:
            if st.button("📊 Ich benötige einen Bericht."): 
                button_prompt = "Ich benötige einen Bericht."
                gewaehlte_aktion = "Bericht"
        with col4:
            if st.button("📝 Ich habe neue Informationen."): 
                button_prompt = "Information: Ich möchte neue Informationen einpflegen."
                gewaehlte_aktion = "Information"
        with col5:
            if st.button("🛠️ Ich möchte eine Änderung am XLS vornehmen."): 
                button_prompt = "Information: Ich möchte eine Änderung am XLS vornehmen."
                gewaehlte_aktion = "Änderung"

    # 2. SCHRITT: DROP-DOWN AUSWAHL ZUR UNTERSTÜTZUNG (DARUNTER)
    st.write("")
    kategorie_auswahl = st.selectbox(
        "Verständnishilfe (Auswahl filtert die Liste der Bezeichnungen):",
        ["Alle Einträge", "Geräte / Ausst. innen", "Geräte / Ausst. außen", "Systeme"]
    )
    
    # 3. SCHRITT: ANZEIGE DER GEFUNDENEN AUSSTATTUNG / BEZEICHNUNGEN
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
            anzeige_df = df_gefiltert[[bez_spalte]].drop_duplicates().reset_index(drop=True)
            st.dataframe(anzeige_df, use_container_width=True, hide_index=True)
        else:
            st.info(f"Keine Einträge für '{kategorie_auswahl}' hinterlegt.")

    # Verarbeitung, falls ein Button gedrückt wurde (Gibt den Drop-down Kontext aktiv mit!)
    if button_prompt:
        final_prompt = f"{button_prompt} (Ausgewählter Kontext-Bereich: {kategorie_auswahl})"
        st.session_state.messages.append({"role": "user", "content": final_prompt})
        
        if gewaehlte_aktion in ["Information", "Änderung"] and df_wissen is not None:
            append_info_to_drive(df_wissen, button_prompt, nutzer_rolle, kategorie_auswahl)
            st.success("Aktion wurde im System vermerkt!")
        
        # KI erhält hier die expliziten Metadaten zur Rolle und dem Drop-Down
        kontext = f"\n\nAktuelle Daten aus der Wissensbasis:\n{df_wissen.to_string(index=False)}" if df_wissen is not None else ""
        try:
            response = model.generate_content(
                f"SYSTEM-KONTEXT: Der Nutzer agiert in der Rolle '{nutzer_rolle}' und hat im Drop-Down aktiv den Bereich '{kategorie_auswahl}' selektiert. "
                f"Nutze diese Einschränkung für deine Antwort!\nAnfrage: {final_prompt} {kontext}"
            )
            st.session_state.messages.append({"role": "assistant", "content": response.text})
            st.rerun()
        except Exception as e:
            st.error(f"Fehler bei der Verarbeitung: {e}")

# Manueller Chat-Input (Gibt den Drop-down Kontext ebenfalls aktiv mit!)
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

        # KI erhält auch beim freien Tippen/Sprechen die genaue Drop-Down Auswahl übermittelt
        kontext = f"\n\nAktuelle Daten aus der Wissensbasis:\n{df_wissen.to_string(index=False)}" if df_wissen is not None else ""
        with st.chat_message("assistant"):
            try:
                response = model.generate_content(
                    f"SYSTEM-KONTEXT: Der Nutzer tippt eine freie Frage in der Rolle '{nutzer_rolle}'. "
                    f"Im Drop-Down ist aktuell '{kategorie_auswahl if 'kategorie_auswahl' in locals() else 'Alle Einträge'}' ausgewählt. "
                    f"Beziehe seine Frage primär auf diesen ausgewählten Bereich!\nAnfrage: {prompt} {kontext}"
                )
                st.markdown(response.text)
                st.session_state.messages.append({"role": "assistant", "content": response.text})
            except Exception as e:
                st.error(f"Fehler bei der KI-Verarbeitung: {e}")
