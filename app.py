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
        df = pd.read_excel(fh)
        
        # VERZEICHNIS-LOGIK: Leere Zellen in Spalte A (Kategorie) automatisch auffüllen
        if df is not None and not df.empty:
            df.iloc[:, 0] = df.iloc[:, 0].ffill()
            
        return df, service
    except Exception as e:
        st.error(f"Fehler bei der Verbindung zur Google Drive Wissensbasis: {e}")
        return None, None

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
            config=types.GenerateContentConfig(system_instruction=VILLA_PROMPT)
        )
        return response.text
    except Exception as e:
        try:
            response = client.models.generate_content(
                model="gemini-1.5-flash",
                contents=prompt_text,
                config=types.GenerateContentConfig(system_instruction=VILLA_PROMPT)
            )
            return response.text
        except Exception:
            return f"Fehler bei der KI-Verarbeitung: {e}"

# ==========================================
# 4. BENUTZEROBERFLÄCHE (HMI) & MESSENGER-STYLING
# ==========================================
st.set_page_config(page_title="Villa Avatar", page_icon="☀️", layout="centered")

st.markdown("""
    <style>
    div[data-testid="stSelectbox"] div[data-baseweb="select"] { font-weight: bold; font-size: 15px; }
    
    div[data-testid="stChatMessage"]:has(div[aria-label="Chat message from user"]) {
        flex-direction: row-reverse !important;
        background-color: rgba(0, 0, 0, 0.03) !important;
        border-radius: 10px !important;
        padding: 10px !important;
    }
    div[data-testid="stChatMessage"]:has(div[aria-label="Chat message from user"]) div[data-testid="stChatMessageContent"] {
        text-align: right !important;
        width: 100% !important;
    }
    div[data-testid="stChatMessage"]:has(div[aria-label="Chat message from user"]) div[data-testid="stMarkdownContainer"] p {
        text-align: right !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("☀️ Villa Avatar")
st.markdown("Hallo! Ich bin Villa Avatar, dein digitaler **'Helfer'**! Wähle unten die Rolle aus, um zu beginnen.")

if "messages" not in st.session_state:
    st.session_state.messages = []
if "aktive_aktion" not in st.session_state:
    st.session_state.aktive_aktion = None
if "aktive_frage" not in st.session_state:
    st.session_state.aktive_frage = None
if "vorherige_rolle" not in st.session_state:
    st.session_state.vorherige_rolle = "Bitte auswählen..."

# Callback: Setzt alte Drop-down-Inhalte im Speicher zurück
def reset_dropdown_states():
    for key in list(st.session_state.keys()):
        if key.startswith("sub_cat_wahl_"):
            del st.session_state[key]

# Callback: Verarbeitet Klicks und zwingt das UI zum sofortigen, synchronen Neu-Rendern (Issue 33 & 34)
def on_button_click(aktion_name, frage_text):
    reset_dropdown_states()
    st.session_state.aktive_aktion = aktion_name
    st.session_state.aktive_frage = frage_text
    st.rerun()

nutzer_rolle = st.selectbox("Wer bist du?", ["Bitte auswählen...", "Besucher", "Eigentümer", "Administrator", "Handwerker/Helfer"])

if nutzer_rolle != st.session_state.vorherige_rolle:
    st.session_state.vorherige_rolle = nutzer_rolle
    st.session_state.aktive_aktion = None
    st.session_state.aktive_frage = None
    st.session_state.messages = []  
    reset_dropdown_states()
    st.rerun()

if nutzer_rolle != "Bitte auswählen...":
    st.write("---")
    st.subheader("Mein Anliegen:")
    
    if nutzer_rolle == "Besucher":
        col1, col2 = st.columns(2)
        with col1:
            st.button("Ich brauche Hilfe.", use_container_width=True, 
                      type="primary" if st.session_state.aktive_aktion == "Hilfe" else "secondary",
                      on_click=on_button_click, args=("Hilfe", "Wobei kann ich dir helfen?"))
        with col2:
            st.button("Es gibt eine Störung.", use_container_width=True, 
                      type="primary" if st.session_state.aktive_aktion == "Störung" else "secondary",
                      on_click=on_button_click, args=("Störung", "Was ist passiert?"))
    else:
        col1, col2, col3 = st.columns(3)
        col4, col5 = st.columns(2)
        with col1:
            st.button("Ich brauche Hilfe.", use_container_width=True, 
                      type="primary" if st.session_state.aktive_aktion == "Hilfe" else "secondary",
                      on_click=on_button_click, args=("Hilfe", "Wobei kann ich dir helfen?"))
        with col2:
            st.button("Ich habe neue Informationen.", use_container_width=True, 
                      type="primary" if st.session_state.aktive_aktion == "Information" else "secondary",
                      on_click=on_button_click, args=("Information", "Gern nehme ich deine Informationen auf und ordne sie in meiner Wissensbasis zu."))
        with col3:
            st.button("Es gibt eine Störung.", use_container_width=True, 
                      type="primary" if st.session_state.aktive_aktion == "Störung" else "secondary",
                      on_click=on_button_click, args=("Störung", "Was ist passiert?"))
        with col4:
            st.button("Ich benötigt einen Bericht.", use_container_width=True, 
                      type="primary" if st.session_state.aktive_aktion == "Bericht" else "secondary",
                      on_click=on_button_click, args=("Bericht", "Nenne mir bitte den Zeitraum und das Thema."))
        with col5:
            st.button("Ich möchte eine Änderung an der Wissensbasis vornehmen.", use_container_width=True, 
                      type="primary" if st.session_state.aktive_aktion == "Änderung" else "secondary",
                      on_click=on_button_click, args=("Änderung", "Beschreibe deine Änderung so genau wie möglich."))

    # ==========================================
    # 5. DROPDOWN-MENÜS (DYNAMISCHER KEY-RESET)
    # ==========================================
    if st.session_state.aktive_aktion:
        st.write("")
        # Issue 33 Fix: Zeigt jetzt garantiert synchron die exakte Frage an
        st.info(f"**Villa Avatar:** {st.session_state.aktive_frage}")
        
        kategorien_fuer_rolle = []
        if st.session_state.aktive_aktion == "Hilfe":
            kategorien_fuer_rolle = ["Ausstattung innen", "Ausstattung außen"]
        else:
            if nutzer_rolle == "Handwerker/Helfer":
                kategorien_fuer_rolle = ["Systeme", "Ausstattung außen"]
            else:
                kategorien_fuer_rolle = ["Systeme", "Ausstattung innen", "Ausstattung außen"]

        konkrete_auswahlen = {}
        
        if df_wissen is not None and not df_wissen.empty:
            kat_spalte = df_wissen.columns[0]
            bez_spalte = df_wissen.columns[1] if len(df_wissen.columns) > 1 else df_wissen.columns[0]

            for kat in kategorien_fuer_rolle:
                mask = df_wissen[kat_spalte].astype(str).str.strip() == kat
                verfuegbare_bezeichnungen = df_wissen[mask][bez_spalte].dropna().drop_duplicates().tolist()
                verfuegbare_bezeichnungen = sorted([str(b).strip() for b in verfuegbare_bezeichnungen])
                
                # Issue 34 Fix: Durch den dynamischen Key im Format f"..._{st.session_state.aktive_aktion}"
                # wird das Drop-down bei jedem Button-Wechsel komplett auf den Ursprung zurückgesetzt!
                wahl = st.selectbox(
                    label=f"Hidden_Label_{kat}", 
                    options=verfuegbare_bezeichnungen,
                    index=None,
                    placeholder=kat,
                    key=f"sub_cat_wahl_{kat}_{st.session_state.aktive_aktion}",
                    label_visibility="collapsed"
                )
                
                if wahl is not None:
                    konkrete_auswahlen[kat] = wahl

# ==========================================
# 6. CHAT-ANZEIGE UND MANUELLER INPUT
# ==========================================
st.write("---")
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

if prompt := st.chat_input("Wie kann ich helfen?"):
    if nutzer_rolle == "Bitte auswählen...":
        st.warning("Bitte wähle oben zuerst aus, wer du bist!")
    elif not st.session_state.aktive_aktion:
        st.warning("Bitte wähle oben zuerst ein Anliegen aus!")
    else:
        with st.chat_message("user"):
            st.markdown(prompt)
        st.session_state.messages.append({"role": "user", "content": prompt})
        
        gewaehltes_objekt = list(konkrete_auswahlen.values())[0] if konkrete_auswahlen else ""
        
        if st.session_state.aktive_aktion == "Information" and df_wissen is not None:
            kat_text = ", ".join(konkrete_auswahlen.keys()) if konkrete_auswahlen else "Allgemein"
            with st.spinner("Eintrag wird in Google Drive gespeichert..."):
                append_info_to_drive(df_wissen, prompt, nutzer_rolle, kat_text)
                st.cache_data.clear()
                df_wissen, _ = load_data_from_drive()

        kontext = f"\n\nAktuelle Daten aus der Wissensbasis:\n{df_wissen.to_string(index=False)}" if df_wissen is not None else ""
        
        with st.chat_message("assistant"):
            with st.spinner("Villa Avatar überlegt..."):
                antwort_text = generate_ki_response(
                    f"Rolle: {nutzer_rolle}\n"
                    f"Kontext-Aktion des Nutzers: {st.session_state.aktive_aktion}\n"
                    f"Gewähltes HMI-Objekt: {gewaehltes_objekt}\n"
                    f"Anfrage: {prompt} {kontext}"
                )
            st.markdown(antwort_text)
            st.session_state.messages.append({"role": "assistant", "content": antwort_text})
