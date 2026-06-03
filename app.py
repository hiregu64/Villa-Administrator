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
        
        # Verzeichnis-Logik für Spalte B (Kategorie) automatisch auffüllen
        if df is not None and not df.empty and len(df.columns) > 1:
            df.iloc[:, 1] = df.iloc[:, 1].ffill()
            
        return df, service
    except Exception as e:
        st.error(f"Fehler bei der Verbindung zur Google Drive Wissensbasis: {e}")
        return None, None

with st.spinner("Verbindung zur Google Drive Wissensbasis wird hergestellt..."):
    df_wissen, drive_service = load_data_from_drive()

# ==========================================
# 2. LIVE-UPDATE (INFORMATION & FEEDBACK: Zeile anhängen)
# ==========================================
def append_info_to_drive(df, neuer_text, nutzername, kategorie="Nicht definiert"):
    try:
        neue_zeile = {
            df.columns[0]: f"Update: {neuer_text}", 
            df.columns[1]: kategorie,
            "Zeitstempel": datetime.datetime.now().strftime("%d.%m.%Y %H:%M"),
            "Nutzer": nutzername
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
# 2b. STÖRUNGSLOGIK (Spalten J & K aktualisieren)
# ==========================================
def update_stoerung_in_drive(df, stoerung_text, nutzername, objekt_name=None):
    try:
        # Absicherung: Falls die Excel weniger als 11 Spalten eingelesen hat (da leer), auffüllen bis K (Index 10)
        while len(df.columns) < 11:
            df[f"Spalte_{len(df.columns)+1}"] = ""
            
        row_idx = None
        
        # 1. Versuchen, das konkrete ausgewählte Objekt in Spalte A zu finden
        if objekt_name:
            match = df[df.iloc[:, 0].astype(str).str.strip().str.lower() == str(objekt_name).strip().lower()]
            if not match.empty:
                row_idx = match.index[0]
                
        # 2. Wenn kein Objekt gewählt wurde oder das Thema unzuordenbar ist -> Zeile "Nicht gefunden" nutzen
        if row_idx is None:
            match_nf = df[df.iloc[:, 0].astype(str).str.strip().str.lower() == "nicht gefunden"]
            if not match_nf.empty:
                row_idx = match_nf.index[0]
            else:
                # Fallback, falls die Zeile "Nicht gefunden" physisch in Excel fehlt, legen wir sie an
                neue_zeile = {df.columns[0]: "Nicht gefunden"}
                df = pd.concat([df, pd.DataFrame([neue_zeile])], ignore_index=True)
                row_idx = len(df) - 1

        # Daten für Spalte J (Index 9) und Spalte K (Index 10) vorbereiten & Historie wahren
        zeitstempel = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
        
        alt_j = str(df.iloc[row_idx, 9]) if pd.notna(df.iloc[row_idx, 9]) else ""
        alt_k = str(df.iloc[row_idx, 10]) if pd.notna(df.iloc[row_idx, 10]) else ""
        
        neu_j = f"{alt_j}\n- {stoerung_text}".strip() if alt_j else f"- {stoerung_text}"
        neu_k = f"{alt_k}\n- {zeitstempel} ({nutzername})".strip() if alt_k else f"- {zeitstempel} ({nutzername})"
        
        df.iloc[row_idx, 9] = neu_j
        df.iloc[row_idx, 10] = neu_k
        
        # Datei zurückschreiben
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        output.seek(0)
        
        media = MediaIoBaseUpload(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        drive_service.files().update(fileId=FILE_ID, media_body=media).execute()
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Fehler beim Loggen der Störung in Google Drive: {e}")
        return False

# ==========================================
# 3. KI-GEHIRN (SYSTEM-PROMPT)
# ==========================================
VILLA_PROMPT = """
Du bist „Villa Avatar“, der digitale Helfer für die Gäste (Gast), Gastgeber (Host) und Administratoren (Admin) der Villa. Deine Aufgabe ist es, den Aufenthalt, Betrieb und Erhalt des Hauses so einfach und angenehm wie möglich zu gestalten.

WICHTIGER KONTEXT & VERHALTEN:
- Antworte immer kurz, freundlich, präzise und smartphone-optimiert (nutze kurze Absätze oder Aufzählungspunkte).
- Passe deine Tonalität an die übergebene Rolle an: Zu Gästen (Gast) bist du einladend und herzlich, zu Hosts und Admins agierst du effizient und lösungsorientiert.
- Beziehe dich bei Antworten exakt auf die mitgegebenen Live-Daten aus der Wissensbasis. 
- WICHTIG (Wissenslücken): Wenn die mitgegebenen Daten keine Antwort auf die Frage des Nutzers enthalten, erfinde niemals Informationen! Antworte stattdessen: „Dazu liegen mir aktuell leider keine Informationen vor. Ich leite dein Anliegen aber gerne an das Team weiter.“
- ABSOLUTES VERBOT: Erwähne NIEMALS interne Dateinamen, Bildbezeichnungen (wie '.jfif' oder '.jpg') oder die Struktur der Excel-Tabelle (wie 'Spalte A', 'Spalte B', Überschriften oder Zeilen). Antworte so, als hättest du dieses Wissen einfach natürlich im Kopf.
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
        error_msg = str(e)
        if "429" in error_msg or "RESOURCE_EXHAUSTED" in error_msg or "503" in error_msg or "UNAVAILABLE" in error_msg:
            try:
                response = client.models.generate_content(
                    model="gemini-1.5-flash",
                    contents=prompt_text,
                    config=types.GenerateContentConfig(system_instruction=VILLA_PROMPT)
                )
                return response.text
            except Exception as e_fallback:
                if "429" in str(e_fallback) or "RESOURCE_EXHAUSTED" in str(e_fallback):
                    return "🛑 Das tägliche kostenlose Abfrage-Limit der Villa Avatar ist leider für heute aufgebraucht. Bitte versuche es morgen wieder!"
                return "⏳ Die KI-Server sind aktuell stark ausgelastet. Bitte warte einen kurzen Moment und sende deine Nachricht noch einmal."
        
        return f"Fehler bei der KI-Verarbeitung: {e}"

# ==========================================
# 4. BENUTZEROBERFLÄCHE (HMI) & STYLING
# ==========================================
st.set_page_config(page_title="Villa Avatar", page_icon="☀️", layout="centered")

st.markdown("""
    <style>
    div.stButton > button[kind="primary"] {
        background-color: #e3f2fd !important;
        color: #1565c0 !important;
        border: 1px solid #bbdefb !important;
        font-weight: bold !important;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #bbdefb !important;
        border: 1px solid #64b5f6 !important;
    }
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
# 5. DIE EXAKTE HMI-ZUSTANDSMATRIX
# ==========================================
STANDARD_DROPDOWNS = ["Ausstattung innen", "Ausstattung außen", "In der Nähe"]

HMI_MATRIX = {
    "Gast": {
        "Hilfe": {"text": "Wobei kann ich dir helfen?", "dd": STANDARD_DROPDOWNS},
        "Feedback": {"text": "Welches Feedback hast du?", "dd": STANDARD_DROPDOWNS},
        "Störung": {"text": "Was ist passiert?", "dd": STANDARD_DROPDOWNS}
    },
    "Host": {
        "Hilfe": {"text": "Wobei kann ich dir helfen?", "dd": STANDARD_DROPDOWNS},
        "Information": {"text": "Gern nehme ich deine Informationen auf und ordne sie in meiner Wissensbasis zu.", "dd": STANDARD_DROPDOWNS},
        "Feedback": {"text": "Welches Feedback hast du?", "dd": STANDARD_DROPDOWNS},
        "Störung": {"text": "Was ist passiert?", "dd": STANDARD_DROPDOWNS},
        "Bericht": {"text": "Nenne mir bitte den Zeitraum und das Thema.", "dd": []}
    },
    "Admin": {
        "Hilfe": {"text": "Wobei kann ich dir helfen?", "dd": STANDARD_DROPDOWNS},
        "Information": {"text": "Gern nehme ich deine Informationen auf und ordne sie in meiner Wissensbasis zu.", "dd": STANDARD_DROPDOWNS},
        "Feedback": {"text": "Welches Feedback hast du?", "dd": STANDARD_DROPDOWNS},
        "Störung": {"text": "Was ist passiert?", "dd": STANDARD_DROPDOWNS},
        "Bericht": {"text": "Nenne mir bitte den Zeitraum und das Thema.", "dd": []},
        "Änderung": {"text": "Beschreibe deine Änderung so genau wie möglich.", "dd": STANDARD_DROPDOWNS}
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
    st.session_state.messages = []  
    st.rerun()

nutzer_rolle = st.selectbox(
    label="Hidden_Rollen_Label",
    options=["Gast", "Host", "Admin"],
    index=None,
    placeholder="Wer bist du?",
    label_visibility="collapsed"
)

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
    
    with st.container():
        st.markdown(
            "<div style='display: flex; justify-content: flex-end; align-items: center; gap: 8px; margin-bottom: 10px;'>"
            "<span style='font-weight: bold; font-size: 1.2rem; font-family: inherit;'>Mein Anliegen:</span>"
            "<div style='width: 32px; height: 32px; background-color: rgb(255, 75, 75); border-radius: 8px; display: flex; align-items: center; justify-content: center;'>"
            "<svg viewBox='0 0 24 24' width='20' height='20' stroke='white' stroke-width='2' fill='none' stroke-linecap='round' stroke-linejoin='round'>"
            "<circle cx='12' cy='12' r='10'></circle>"
            "<path d='M8 14s1.5 2 4 2 4-2 4-2'></path>"
            "<line x1='9' y1='9' x2='9.01' y2='9'></line>"
            "<line x1='15' y1='9' x2='15.01' y2='9'></line>"
            "</svg>"
            "</div>"
            "</div>", 
            unsafe_allow_html=True
        )
    
    if nutzer_rolle == "Gast":
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("Ich brauche Hilfe.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Hilfe" else "secondary"):
                handle_button_click("Hilfe")
        with col2:
            if st.button("Ich möchte Feedback geben.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Feedback" else "secondary"):
                handle_button_click("Feedback")
        with col3:
            if st.button("Es gibt eine Störung.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Störung" else "secondary"):
                handle_button_click("Störung")
                
    elif nutzer_rolle == "Host":
        col1, col2, col3 = st.columns(3)
        col4, col5 = st.columns(2)
        with col1:
            if st.button("Ich brauche Hilfe.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Hilfe" else "secondary"):
                handle_button_click("Hilfe")
        with col2:
            if st.button("Ich habe neue Informationen.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Information" else "secondary"):
                handle_button_click("Information")
        with col3:
            if st.button("Ich möchte Feedback geben.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Feedback" else "secondary"):
                handle_button_click("Feedback")
        with col4:
            if st.button("Es gibt eine Störung.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Störung" else "secondary"):
                handle_button_click("Störung")
        with col5:
            if st.button("Ich benötige einen Bericht.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Bericht" else "secondary"):
                handle_button_click("Bericht")

    elif nutzer_rolle == "Admin":
        col1, col2, col3 = st.columns(3)
        col4, col5, col6 = st.columns(3)
        with col1:
            if st.button("Ich brauche Hilfe.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Hilfe" else "secondary"):
                handle_button_click("Hilfe")
        with col2:
            if st.button("Ich habe neue Informationen.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Information" else "secondary"):
                handle_button_click("Information")
        with col3:
            if st.button("Ich möchte Feedback geben.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Feedback" else "secondary"):
                handle_button_click("Feedback")
        with col4:
            if st.button("Es gibt eine Störung.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Störung" else "secondary"):
                handle_button_click("Störung")
        with col5:
            if st.button("Ich benötige einen Bericht.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Bericht" else "secondary"):
                handle_button_click("Bericht")
        with col6:
            if st.button("Ich möchte eine Änderung an der Wissensbasis vornehmen.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Änderung" else "secondary"):
                handle_button_click("Änderung")

    if st.session_state.aktive_aktion and nutzer_rolle in HMI_MATRIX:
        aktiver_state = HMI_MATRIX[nutzer_rolle].get(st.session_state.aktive_aktion)
        
        if aktiver_state:
            st.write("")
            with st.chat_message("assistant"):
                st.markdown(aktiver_state['text'])
            
            kategorien_fuer_rolle = aktiver_state["dd"]
            
            if df_wissen is not None and not df_wissen.empty:
                bez_spalte = df_wissen.columns[0]  
                kat_spalte = df_wissen.columns[1]  

                for kat in kategorien_fuer_rolle:
                    if "innen" in kat.lower():
                        mask = df_wissen[kat_spalte].astype(str).str.contains("innen", case=False, na=False)
                    elif "außen" in kat.lower() or "aussen" in kat.lower():
                        mask = df_wissen[kat_spalte].astype(str).str.contains("außen|aussen", case=False, na=False)
                    elif "nähe" in kat.lower() or "naehe" in kat.lower():
                        mask = df_wissen[kat_spalte].astype(str).str.contains("nähe|naehe|In der Nähe", case=False, na=False)
                    else:
                        mask = df_wissen[kat_spalte].astype(str).str.contains(kat, case=False, na=False)
                    
                    verfuegbare_bezeichnungen = df_wissen[mask][bez_spalte].dropna().drop_duplicates().tolist()
                    verfuegbare_bezeichnungen = sorted([str(b).strip() for b in verfuegbare_bezeichnungen])
                    
                    st.selectbox(
                        label=f"Hidden_Label_{kat}", 
                        options=verfuegbare_bezeichnungen,
                        index=None,
                        placeholder=f"📍 {kat} wählen...",
                        key=f"sub_cat_wahl_{kat}_{st.session_state.aktive_aktion}",
                        label_visibility="collapsed"
                    )

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
        
        # Auslesen der aktiven Drop-down Filter
        konkrete_auswahlen = {}
        if st.session_state.aktive_aktion and nutzer_rolle in HMI_MATRIX:
            for kat in HMI_MATRIX[nutzer_rolle][st.session_state.aktive_aktion]["dd"]:
                key = f"sub_cat_wahl_{kat}_{st.session_state.aktive_aktion}"
                if key in st.session_state and st.session_state[key] is not None:
                    konkrete_auswahlen[kat] = st.session_state[key]
        
        gewaehlte_objekte_str = ", ".join([f"{k}: {v}" for k, v in konkrete_auswahlen.items()]) if konkrete_auswahlen else "Keines ausgewählt"
        gewaehltes_objekt = list(konkrete_auswahlen.values())[0] if konkrete_auswahlen else None
        
        # 1. LOGIK FÜR INFORMATION / FEEDBACK (Zeile anhängen)
        if st.session_state.aktive_aktion in ["Information", "Feedback"] and df_wissen is not None:
            kat_text = ", ".join(konkrete_auswahlen.keys()) if konkrete_auswahlen else "Allgemein"
            with st.spinner("Eintrag wird in Google Drive gespeichert..."):
                append_info_to_drive(df_wissen, prompt, nutzer_rolle, kat_text)
                st.cache_data.clear()
                df_wissen, _ = load_data_from_drive()
                
        # 2. LOGIK FÜR STÖRUNGEN (Spalte J & K aktualisieren in Zielzeile oder "Nicht gefunden")
        elif st.session_state.aktive_aktion == "Störung" and df_wissen is not None:
            with st.spinner("Störung wird in der Wissensbasis (Spalten J/K) protokolliert..."):
                update_stoerung_in_drive(df_wissen, prompt, nutzer_rolle, gewaehltes_objekt)
                st.cache_data.clear()
                df_wissen, _ = load_data_from_drive()

        kontext = f"\n\nAktuelle Daten aus der Wissensbasis:\n{df_wissen.to_string(index=False)}" if df_wissen is not None else ""
        
        with st.chat_message("assistant"):
            with st.spinner("Villa Avatar überlegt..."):
                antwort_text = generate_ki_response(
                    f"Rolle: {nutzer_rolle}\n"
                    f"Kontext-Aktion des Nutzers: {st.session_state.aktive_aktion}\n"
                    f"Gewählte(s) HMI-Objekt(e): {gewaehlte_objekte_str}\n"
                    f"Anfrage: {prompt} {kontext}"
                )
            st.markdown(antwort_text)
            st.session_state.messages.append({"role": "assistant", "content": antwort_text})
