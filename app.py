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
# 3. KI-GEHIRN (PPT-KONFORMER SYSTEM-PROMPT)
# ==========================================
VILLA_PROMPT = """
Du bist „Villa Avatar“, der digitale Helfer für die Bewohner, Eigentümer und Admins der Villa. Deine Aufgabe ist es, den Betrieb und Erhalt des Hauses so einfach wie möglich zu halten.

WICHTIGER KONTEXT & VERHALTEN:
- Antworte immer kurz, präzise und smartphone-optimiert.
- Nutze die vom HMI übergebene Rolle (Besucher, Eigentümer oder Admin) und die gewählte Kategorie/Bezeichnung zwingend als Arbeitsgrundlage.
- Beziehe dich exakt auf die übergebenen Daten aus der Wissensbasis.
- WICHTIG: Erwähne NIEMALS interne Dateinamen, Bildbezeichnungen (wie '.jfif' oder '.jpg') oder Tabellenstrukturen gegenüber dem Nutzer. Antworte so, als hättest du dieses Wissen einfach im Kopf.
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
        if "503" in str(e) or "UNAVAILABLE" in str(e):
            return "⏳ Die Google Gemini Server sind aktuell stark ausgelastet. Bitte warte einen kurzen Moment und sende deine Nachricht gleich noch einmal!"
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
# 4. BENUTZEROBERFLÄCHE (HMI) & STYLING
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

# ==========================================
# 5. DIE EXAKTE HMI-ZUSTANDSMATRIX (REIN PPT-KONFORM)
# ==========================================
HMI_MATRIX = {
    "Besucher": {
        "Hilfe": {"text": "Wobei kann ich dir helfen?", "dd": ["Ausstattung innen", "Ausstattung außen"]},
        "Störung": {"text": "Was ist passiert?", "dd": ["Ausstattung innen", "Ausstattung außen"]}
    },
    "Eigentümer": {
        "Hilfe": {"text": "Wobei kann ich dir helfen?", "dd": ["Ausstattung innen", "Ausstattung außen"]},
        "Information": {"text": "Gern nehme ich deine Informationen auf und ordne sie in meiner Wissensbasis zu.", "dd": ["Systeme", "Ausstattung innen", "Ausstattung außen"]},
        "Störung": {"text": "Was ist passiert?", "dd": ["Systeme", "Ausstattung innen", "Ausstattung außen"]},
        "Bericht": {"text": "Nenne mir bitte den Zeitraum und das Thema.", "dd": ["Systeme", "Ausstattung innen", "Ausstattung außen"]}
    },
    "Admin": {
        "Hilfe": {"text": "Wobei kann ich dir helfen?", "dd": ["Ausstattung innen", "Ausstattung außen"]},
        "Information": {"text": "Gern nehme ich deine Informationen auf und ordne sie in meiner Wissensbasis zu.", "dd": ["Systeme", "Ausstattung innen", "Ausstattung außen"]},
        "Störung": {"text": "Was ist passiert?", "dd": ["Systeme", "Ausstattung innen", "Ausstattung außen"]},
        "Bericht": {"text": "Nenne mir bitte den Zeitraum und das Thema.", "dd": ["Systeme", "Ausstattung innen", "Ausstattung außen"]},
        "Änderung": {"text": "Beschreibe deine Änderung so genau wie möglich.", "dd": ["Systeme", "Ausstattung innen", "Ausstattung außen"]}
    }
}

if "messages" not in st.session_state:
    st.session_state.messages = []
if "aktive_aktion" not in st.session_state:
    st.session_state.aktive_aktion = None
if "vorherige_rolle" not in st.session_state:
    st.session_state.vorherige_rolle = None

def handle_button_click(aktions_name):
    for key in list(st.session_state.keys()):
        if key.startswith("sub_cat_wahl_"):
            del st.session_state[key]
    st.session_state.aktive_aktion = aktions_name
    st.session_state.messages = []  # Chat löschen bei Button-Wechsel
    st.rerun()

# Kompakte Rollenauswahl (Besucher, Eigentümer, Admin)
nutzer_rolle = st.selectbox(
    label="Hidden_Rollen_Label",
    options=["Besucher", "Eigentümer", "Admin"],
    index=None,
    placeholder="Wer bist du?",
    label_visibility="collapsed"
)

# Dynamik bei Änderung der Rolle weiter oben umsetzen
if nutzer_rolle != st.session_state.vorherige_rolle:
    st.session_state.vorherige_rolle = nutzer_rolle
    st.session_state.aktive_aktion = None
    st.session_state.messages = []  
    for key in list(st.session_state.keys()):
        if key.startswith("sub_cat_wahl_"):
            del st.session_state[key]
    st.rerun()

if nutzer_rolle is not None:
    st.write("---")
    st.subheader("Mein Anliegen:")
    
    # Exaktes Zeichnen der Knöpfe je nach PPT-Rolle
    if nutzer_rolle == "Besucher":
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Ich brauche Hilfe.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Hilfe" else "secondary"):
                handle_button_click("Hilfe")
        with col2:
            if st.button("Es gibt eine Störung.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Störung" else "secondary"):
                handle_button_click("Störung")
                
    elif nutzer_rolle == "Admin":
        col1, col2, col3 = st.columns(3)
        col4, col5 = st.columns(2)
        with col1:
            if st.button("Ich brauche Hilfe.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Hilfe" else "secondary"):
                handle_button_click("Hilfe")
        with col2:
            if st.button("Ich habe neue Informationen.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Information" else "secondary"):
                handle_button_click("Information")
        with col3:
            if st.button("Es gibt eine Störung.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Störung" else "secondary"):
                handle_button_click("Störung")
        with col4:
            if st.button("Ich benötige einen Bericht.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Bericht" else "secondary"):
                handle_button_click("Bericht")
        with col5:
            if st.button("Ich möchte eine Änderung an der Wissensbasis vornehmen.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Änderung" else "secondary"):
                handle_button_click("Änderung")
                
    else:  # Eigentümer
        col1, col2 = st.columns(2)
        col3, col4 = st.columns(2)
        with col1:
            if st.button("Ich brauche Hilfe.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Hilfe" else "secondary"):
                handle_button_click("Hilfe")
        with col2:
            if st.button("Ich habe neue Informationen.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Information" else "secondary"):
                handle_button_click("Information")
        with col3:
            if st.button("Es gibt eine Störung.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Störung" else "secondary"):
                handle_button_click("Störung")
        with col4:
            if st.button("Ich benötige einen Bericht.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Bericht" else "secondary"):
                handle_button_click("Bericht")

    # ==========================================
    # 6. ABSOLUT SYNCHRONE GEGENFRAGEN & DROP-DOWNS
    # ==========================================
    if st.session_state.aktive_aktion and nutzer_rolle in HMI_MATRIX:
        aktiver_state = HMI_MATRIX[nutzer_rolle].get(st.session_state.aktive_aktion)
        
        if aktiver_state:
            st.write("")
            st.info(f"**Villa Avatar:** {aktiver_state['text']}")
            
            kategorien_fuer_rolle = aktiver_state["dd"]
            konkrete_auswahlen = {}
            
            if df_wissen is not None and not df_wissen.empty:
                kat_spalte = df_wissen.columns[0]
                bez_spalte = df_wissen.columns[1] if len(df_wissen.columns) > 1 else df_wissen.columns[0]

                for kat in kategorien_fuer_rolle:
                    mask = df_wissen[kat_spalte].astype(str).str.strip() == kat
                    verfuegbare_bezeichnungen = df_wissen[mask][bez_spalte].dropna().drop_duplicates().tolist()
                    verfuegbare_bezeichnungen = sorted([str(b).strip() for b in verfuegbare_bezeichnungen])
                    
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
# 7. CHAT-ANZEIGE UND MANUELLER INPUT
# ==========================================
st.write("---")
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

if prompt := st.chat_input("Bitte schreibe hier oder sprich mit mir 🎙️"):
    if nutzer_rolle is None:
        st.warning("Bitte wähle oben zuerst aus, wer du bist!")
    elif not st.session_state.aktive_aktion:
        st.warning("Bitte wähle oben zuerst ein Anliegen aus!")
    else:
        with st.chat_message("user"):
            st.markdown(prompt)
        st.session_state.messages.append({"role": "user", "content": prompt})
        
        gewaehltes_objekt = list(konkrete_auswahlen.values())[0] if 'konkrete_auswahlen' in locals() and konkrete_auswahlen else ""
        
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
