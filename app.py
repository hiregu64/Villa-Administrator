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
import openpyxl
from openpyxl.styles import Font  # Import für die farbliche Kennzeichnung

# Google File ID der Excel-Tabelle
FILE_ID = '1FzhWZuO6aRZkdRuQBzaojhkq7bQDyprl'

# Globale Definition deiner Spaltenübersichten (A = 1, B = 2 ... Y = 25)
SPALTEN_MAP = {
    "Bezeichnung": 1, "Wo?": 2, "Relevanz Gast": 3, "System": 4, "Marke/ Typ": 5,
    "Besonderheit": 6, "Quelle Handwerker/ Verkäufer": 7, "Details Nutzung": 8, "Details Steuerung": 9,
    "Störung": 10, "Störung Status": 11, "Wartung (Betrieb)": 12, "Vorsorge": 13,
    "Ersatzteile": 14, "Ersatzteil Quelle": 15, "Ersatzteil Lagerort": 16, "Details Vorsorge Wartung": 17,
    "Wartung (Historie)": 18, "Schlüssel (HW, SW)": 19, "Schlüssel Gast": 20,
    "Dokumente/ Link zur Anleitung": 21, "Kontakt": 22, "Kosten": 23, "Feedback": 24, "Feedback Status": 25
}

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
            _, done = downloader.next_chunk()
            
        fh.seek(0)
        df = pd.read_excel(fh)
        
        if df is not None and not df.empty and len(df.columns) > 1:
            df.iloc[:, 1] = df.iloc[:, 1].ffill()
            
        return df, service
    except Exception as e:
        st.error(f"Fehler bei der Verbindung zur Google Drive Wissensbasis: {e}")
        return None, None

with st.spinner("Verbindung zur Google Drive Wissensbasis wird hergestellt..."):
    df_wissen, drive_service = load_data_from_drive()

# ==========================================
# 2. DIE VIER SCHREIB-AKTIVITÄTEN
# ==========================================

# AKTIVITÄT 1: Störung und Feedback erfassen (Gast, Host, Admin)
def schreibe_input(service, text, nutzername, objekt_name, aktions_typ):
    try:
        request = service.files().get_media(fileId=FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done: _, done = downloader.next_chunk()
        fh.seek(0)
        
        wb = openpyxl.load_workbook(fh)
        ws = wb.active
        
        row_idx = None
        if objekt_name and str(objekt_name).strip().lower() != "nicht gefunden":
            for r in range(2, ws.max_row + 1):
                val = ws.cell(row=r, column=1).value
                if val and str(val).strip().lower() == str(objekt_name).strip().lower():
                    row_idx = r
                    break
        if row_idx is None:
            for r in range(2, ws.max_row + 1):
                val = ws.cell(row=r, column=1).value
                if val and "nicht gefunden" in str(val).strip().lower():
                    row_idx = r
                    break
        if row_idx is None:
            row_idx = ws.max_row + 1
            ws.cell(row=row_idx, column=1, value="Nicht gefunden")

        # Spalten-Zuweisung laut HMI-Matrix
        if aktions_typ == "Störung":
            t_col, status_col, status_val = 10, 11, "aktiv"  # Spalte J & K
        else:
            t_col, status_col, status_val = 24, 25, "offen"  # Spalte X & Y

        zeitstempel = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
        alt_text = str(ws.cell(row=row_idx, column=t_col).value) if ws.cell(row=row_idx, column=t_col).value is not None else ""
        alt_status = str(ws.cell(row=row_idx, column=status_col).value) if ws.cell(row=row_idx, column=status_col).value is not None else ""
        
        # Text & Status schreiben
        cell_text = ws.cell(row=row_idx, column=t_col, value=f"{alt_text}\n- {text}".strip() if alt_text else f"- {text}")
        cell_status = ws.cell(row=row_idx, column=status_col, value=f"{alt_status}\n- {zeitstempel} ({nutzername}): {status_val}".strip() if alt_status else f"- {zeitstempel} ({nutzername}): {status_val}")
        
        # Blau einfärben (Hex-Code: 0000FF für reines Blau)
        cell_text.font = Font(color="0000FF")
        cell_status.font = Font(color="0000FF")
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        media = MediaIoBaseUpload(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        service.files().update(fileId=FILE_ID, media_body=media).execute()
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Fehler in schreibe_input(): {e}")
        return False

# AKTIVITÄT 2: Neue Informationen in Stammdaten einpflegen (Host, Admin)
def aktualisiere_stammdaten(service, text, nutzername, objekt_name, zielspalte_name):
    try:
        request = service.files().get_media(fileId=FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done: _, done = downloader.next_chunk()
        fh.seek(0)
        
        wb = openpyxl.load_workbook(fh)
        ws = wb.active
        
        row_idx = None
        if objekt_name and str(objekt_name).strip().lower() != "nicht gefunden":
            for r in range(2, ws.max_row + 1):
                val = ws.cell(row=r, column=1).value
                if val and str(val).strip().lower() == str(objekt_name).strip().lower():
                    row_idx = r
                    break
        if row_idx is None:
            for r in range(2, ws.max_row + 1):
                val = ws.cell(row=r, column=1).value
                if val and "nicht gefunden" in str(val).strip().lower():
                    row_idx = r
                    break

        if row_idx is None:
            st.error("Objekt für Stammdatenpflege nicht gefunden.")
            return False

        col_idx = SPALTEN_MAP.get(zielspalte_name, 8)  # Fallback auf Details Nutzung (H)
        
        alt_val = str(ws.cell(row=row_idx, column=col_idx).value) if ws.cell(row=row_idx, column=col_idx).value is not None else ""
        
        # Wert aktualisieren
        cell_info = ws.cell(row=row_idx, column=col_idx, value=f"{alt_val}\n{text}".strip() if alt_val else text)
        
        # Blau einfärben
        cell_info.font = Font(color="0000FF")
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        media = MediaIoBaseUpload(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        service.files().update(fileId=FILE_ID, media_body=media).execute()
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Fehler in aktualisiere_stammdaten(): {e}")
        return False

# AKTIVITÄT 3: Neue Objekte/Zeilen in Struktur anlegen (Nur Admin)
def ändere_struktur(service, text, nutzername, kategorie_text):
    try:
        request = service.files().get_media(fileId=FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done: _, done = downloader.next_chunk()
        fh.seek(0)
        
        wb = openpyxl.load_workbook(fh)
        ws = wb.active
        
        new_row = ws.max_row + 1
        clean_name = text.replace("Füge", "").replace("hinzu", "").replace("Objekt", "").strip()
        if len(clean_name) > 60: clean_name = clean_name[:57] + "..."
        
        # Neue Zeile befüllen
        cell_name = ws.cell(row=new_row, column=1, value=clean_name)
        cell_kat = ws.cell(row=new_row, column=2, value=kategorie_text if kategorie_text else "Ausstattung innen")
        cell_gast = ws.cell(row=new_row, column=3, value="X")  # Standardmäßig für Gäste aktivierbar
        
        # Die komplette neue Zeile (A, B, C) blau einfärben
        cell_name.font = Font(color="0000FF")
        cell_kat.font = Font(color="0000FF")
        cell_gast.font = Font(color="0000FF")
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        media = MediaIoBaseUpload(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        service.files().update(fileId=FILE_ID, media_body=media).execute()
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Fehler in ändere_struktur(): {e}")
        return False

# ==========================================
# 3. KI-GEHIRN & FALLBACK-MANAGEMENT
# ==========================================
VILLA_PROMPT = """
Du bist „Villa Avatar“, der digitale Helfer für die Gäste (Gast), Gastgeber (Host) und Administratoren (Admin) der Villa. Deine Aufgabe ist es, den Aufenthalt, Betrieb und Erhalt des Hauses so einfach und angenehm wie möglich zu gestalten.

WICHTIGER KONTEXT & VERHALTEN:
- Antworte immer kurz, freundlich, präzise und smartphone-optimiert.
- Passe deine Tonalität an die übergebene Rolle an.
- Beziehe dich bei Antworten exakt auf die mitgegebenen Live-Daten aus der Wissensbasis. 
- WICHTIG (Wissenslücken): Wenn die Daten keine Antwort enthalten, erfinde niemals Informationen! Antworte stattdessen: „Dazu liegen mir aktuell leider keine Informationen vor. Ich leite dein Anliegen aber gerne an das Team weiter.“
- ABSOLUTES VERBOT: Erwähne NIEMALS interne Dateinamen, Spaltenbuchstaben oder Tabellenstrukturen.
"""

@st.cache_resource
def get_ki_client():
    if "GEMINI_API_KEY" in st.secrets:
        return genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
    return None

client = get_ki_client()

def ai_waehle_stammdaten_spalte(user_text):
    if client is None: return "Details Nutzung"
    try:
        prompt = f"""
        Analysiere den folgenden Text eines Host/Admins und bestimme, in welche Tabellenspalte diese Information eingetragen werden soll.
        Text: "{user_text}"
        
        Mögliche Zielspalten:
        - Wo?
        - Relevanz Gast
        - System
        - Marke/ Typ
        - Besonderheit
        - Quelle Handwerker/ Verkäufer
        - Details Nutzung
        - Details Steuerung
        - Wartung (Betrieb)
        - Vorsorge
        - Ersatzteile
        - Ersatzteil Quelle
        - Ersatzteil Lagerort
        - Details Vorsorge Wartung
        - Wartung (Historie)
        - Schlüssel (HW, SW)
        - Schlüssel Gast
        - Dokumente/ Link zur Anleitung
        - Kontakt
        - Kosten
        
        Antworte NUR with dem exakten Namen der Spalte aus dieser Liste, ohne weiteren Text. Wenn unklar, antworte mit: Details Nutzung
        """
        response = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        return response.text.strip().replace('"', '').replace("'", "")
    except:
        return "Details Nutzung"

def generate_ki_response(prompt_text):
    if client is None: 
        return "🛑 KI-Dienst nicht konfiguriert: Der API-Key fehlt komplett in den Streamlit Secrets."
    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash", 
            contents=prompt_text, 
            config=types.GenerateContentConfig(system_instruction=VILLA_PROMPT)
        )
        return response.text
    except Exception as e:
        error_msg = str(e)
        if any(err in error_msg for err in ["429", "RESOURCE_EXHAUSTED", "503", "UNAVAILABLE", "504", "TIMEOUT"]):
            try:
                response = client.models.generate_content(
                    model="gemini-2.0-flash",
                    contents=prompt_text,
                    config=types.GenerateContentConfig(system_instruction=VILLA_PROMPT)
                )
                return response.text
            except Exception as e_fallback:
                if "429" in str(e_fallback) or "RESOURCE_EXHAUSTED" in str(e_fallback):
                    return "🛑 Das Limit für kostenlose Anfragen (Quota) bei Google ist aktuell aufgebraucht. Bitte warte eine Minute oder versuche es gleich noch einmal."
                return "🌐 Die Google-KI-Server sind aktuell überlastet. Bitte warte einen kurzen Moment und sende deine Nachricht noch einmal."
        elif any(err in error_msg for err in ["403", "API_KEY_INVALID", "PERMISSION_DENIED"]):
            return "🔑 Zugriff verweigert: Der Google API-Key ist ungültig oder abgelaufen."
        return f"⚠️ Fehler bei der KI-Verarbeitung: {e}"

# ==========================================
# 4. BENUTZEROBERFLÄCHE (HMI) & STYLING
# ==========================================
st.set_page_config(page_title="Villa Avatar", page_icon="☀️", layout="centered")

st.markdown("""
    <style>
    div.stButton > button[kind="primary"] { background-color: #e3f2fd !important; color: #1565c0 !important; border: 1px solid #bbdefb !important; font-weight: bold !important; }
    div.stButton > button[kind="primary"]:hover { background-color: #bbdefb !important; border: 1px solid #64b5f6 !important; }
    div[data-testid="stSelectbox"] div[data-baseweb="select"] { font-weight: bold; font-size: 15px; }
    div[data-testid="stChatMessage"]:has(div[aria-label="Chat message from user"]) { flex-direction: row-reverse !important; background-color: rgba(0, 0, 0, 0.03) !important; border-radius: 10px !important; padding: 10px !important; }
    div[data-testid="stChatMessage"]:has(div[aria-label="Chat message from user"]) div[data-testid="stChatMessageContent"] { text-align: right !important; width: 100% !important; }
    div[data-testid="stChatMessage"]:has(div[aria-label="Chat message from user"]) div[data-testid="stMarkdownContainer"] p { text-align: right !important; }
    </style>
""", unsafe_allow_html=True)

st.title("☀️ Villa Avatar")
st.markdown("Hallo! Ich bin Villa Avatar, dein digitaler **'Helfer'**! Wähle unten die Rolle aus, um zu beginnen.")

STANDARD_DROPDOWNS = ["Ausstattung innen", "Ausstattung außen", "In der Nähe"]

HMI_MATRIX = {
    "Gast": {
        "Hilfe": {"text": "Wobei kann ich dir helfen?", "dd": STANDARD_DROPDOWNS},
        "Feedback": {"text": "Welches Feedback hast du?", "dd": STANDARD_DROPDOWNS},
        "Störung": {"text": "Was ist passiert?", "dd": STANDARD_DROPDOWNS}
    },
    "Host": {
        "Hilfe": {"text": "Wobei kann ich dir helfen?", "dd": STANDARD_DROPDOWNS},
        "Information": {"text": "Gern nehme ich deine Informationen auf und ordne sie zu.", "dd": STANDARD_DROPDOWNS},
        "Feedback": {"text": "Welches Feedback hast du?", "dd": STANDARD_DROPDOWNS},
        "Störung": {"text": "Was ist passiert?", "dd": STANDARD_DROPDOWNS},
        "Bericht": {"text": "Nenne mir bitte den Zeitraum und das Thema.", "dd": []}
    },
    "Admin": {
        "Hilfe": {"text": "Wobei kann ich dir helfen?", "dd": STANDARD_DROPDOWNS},
        "Information": {"text": "Gern nehme ich deine Informationen auf und ordne sie zu.", "dd": STANDARD_DROPDOWNS},
        "Feedback": {"text": "Welches Feedback hast du?", "dd": STANDARD_DROPDOWNS},
        "Störung": {"text": "Was ist passiert?", "dd": STANDARD_DROPDOWNS},
        "Bericht": {"text": "Nenne mir bitte den Zeitraum und das Thema.", "dd": []},
        "Änderung": {"text": "Beschreibe dein neues Objekt, welches der Struktur hinzugefügt werden soll.", "dd": STANDARD_DROPDOWNS}
    }
}

if "messages" not in st.session_state: st.session_state.messages = []
if "aktive_aktion" not in st.session_state: st.session_state.aktive_aktion = None
if "vorherige_rolle" not in st.session_state: st.session_state.vorherige_rolle = None

def handle_button_click(aktions_name):
    for key in list(st.session_state.keys()):
        if key.startswith("sub_cat_wahl_"): del st.session_state[key]
    st.session_state.aktive_aktion = aktions_name
    st.session_state.messages = []  
    st.rerun()

nutzer_rolle = st.selectbox(label="Hidden_Rollen_Label", options=["Gast", "Host", "Admin"], index=None, placeholder="Wer bist du?", key="haupt_nutzer_rolle", label_visibility="collapsed")

if nutzer_rolle != st.session_state.vorherige_rolle:
    st.session_state.vorherige_rolle = nutzer_rolle
    st.session_state.aktive_aktion = None
    st.session_state.messages = []  
    for key in list(st.session_state.keys()):
        if key.startswith("sub_cat_wahl_"): del st.session_state[key]
    st.rerun()

if nutzer_rolle is not None:
    st.write("---")
    
    with st.container():
        st.markdown(
            "<div style='display: flex; justify-content: flex-end; align-items: center; gap: 8px; margin-bottom: 10px;'>"
            "<span style='font-weight: bold; font-size: 1.2rem; font-family: inherit;'>Mein Anliegen:</span>"
            "<div style='width: 32px; height: 32px; background-color: rgb(255, 75, 75); border-radius: 8px; display: flex; align-items: center; justify-content: center;'>"
            "<svg viewBox='0 0 24 24' width='20' height='20' stroke='white' stroke-width='2' fill='none' stroke-linecap='round' stroke-linejoin='round'>"
            "<circle cx='12' cy='12' r='10'></circle><path d='M8 14s1.5 2 4 2 4-2 4-2'></path><line x1='9' y1='9' x2='9.01' y2='9'></line><line x1='15' y1='9' x2='15.01' y2='9'></line>"
            "</svg></div></div>", 
            unsafe_allow_html=True
        )
    
    if nutzer_rolle == "Gast":
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("Ich brauche Hilfe.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Hilfe" else "secondary"): handle_button_click("Hilfe")
        with col2:
            if st.button("Ich möchte Feedback geben.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Feedback" else "secondary"): handle_button_click("Feedback")
        with col3:
            if st.button("Es gibt eine Störung.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Störung" else "secondary"): handle_button_click("Störung")
                
    elif nutzer_rolle in ["Host", "Admin"]:
        col1, col2, col3 = st.columns(3)
        col4, col5, col6 = st.columns(3)
        with col1:
            if st.button("Ich brauche Hilfe.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Hilfe" else "secondary"): handle_button_click("Hilfe")
        with col2:
            if st.button("Ich habe neue Informationen.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Information" else "secondary"): handle_button_click("Information")
        with col3:
            if st.button("Ich möchte Feedback geben.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Feedback" else "secondary"): handle_button_click("Feedback")
        with col4:
            if st.button("Es gibt eine Störung.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Störung" else "secondary"): handle_button_click("Störung")
        with col5:
            if st.button("Ich benötigt einen Bericht.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Bericht" else "secondary"): handle_button_click("Bericht")
        with col6:
            if nutzer_rolle == "Admin":
                if st.button("Ich möchte eine Änderung vornehmen.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Änderung" else "secondary"): handle_button_click("Änderung")
            else:
                st.button("Änderung (Nur Admin)", use_container_width=True, disabled=True)

    if st.session_state.aktive_aktion and nutzer_rolle in HMI_MATRIX:
        aktiver_state = HMI_MATRIX[nutzer_rolle].get(st.session_state.aktive_aktion)
        if aktiver_state:
            st.write("")
            with st.chat_message("assistant"): st.markdown(aktiver_state['text'])
            
            kategorien_fuer_rolle = aktiver_state["dd"]
            if df_wissen is not None and not df_wissen.empty:
                bez_spalte = df_wissen.columns[0]  
                kat_spalte = df_wissen.columns[1]  

                for kat in kategorien_fuer_rolle:
                    if "innen" in kat.lower(): mask = df_wissen[kat_spalte].astype(str).str.contains("innen", case=False, na=False)
                    elif "außen" in kat.lower() or "aussen" in kat.lower(): mask = df_wissen[kat_spalte].astype(str).str.contains("außen|aussen", case=False, na=False)
                    elif "nähe" in kat.lower() or "naehe" in kat.lower(): mask = df_wissen[kat_spalte].astype(str).str.contains("nähe|naehe|In der Nähe", case=False, na=False)
                    else: mask = df_wissen[kat_spalte].astype(str).str.contains(kat, case=False, na=False)
                    
                    if nutzer_rolle == "Gast" and len(df_wissen.columns) > 2:
                        mask = mask & (df_wissen.iloc[:, 2].astype(str).str.strip().str.upper() == "X")
                    
                    verfuegbare_bezeichnungen = df_wissen[mask][bez_spalte].dropna().drop_duplicates().tolist()
                    verfuegbare_bezeichnungen = sorted([str(b).strip() for b in verfuegbare_bezeichnungen])
                    if "Nicht gefunden" in verfuegbare_bezeichnungen: verfuegbare_bezeichnungen.remove("Nicht gefunden")
                    verfuegbare_bezeichnungen.append("Nicht gefunden")
                    
                    st.selectbox(label=f"Hidden_{kat}", options=verfuegbare_bezeichnungen, index=None, placeholder=f"📍 {kat} wählen...", key=f"sub_cat_wahl_{kat}_{st.session_state.aktive_aktion}", label_visibility="collapsed")

st.write("---")
for message in st.session_state.messages:
    with st.chat_message(message["role"]): st.markdown(message["content"])

if prompt := st.chat_input("Bitte schreibe hier oder sprich mit mir 🎙️"):
    if nutzer_rolle is None: st.warning("Bitte wähle oben zuerst aus, wer du bist!")
    elif not st.session_state.aktive_aktion: st.warning("Bitte wähle oben zuerst ein Anliegen aus!")
    else:
        with st.chat_message("user"): st.markdown(prompt)
        st.session_state.messages.append({"role": "user", "content": prompt})
        
        konkrete_auswahlen = {}
        if st.session_state.aktive_aktion and nutzer_rolle in HMI_MATRIX:
            for kat in HMI_MATRIX[nutzer_rolle][st.session_state.aktive_aktion]["dd"]:
                key = f"sub_cat_wahl_{kat}_{st.session_state.aktive_aktion}"
                if key in st.session_state and st.session_state[key] is not None:
                    konkrete_auswahlen[kat] = st.session_state[key]
        
        gewaehlte_objekte_str = ", ".join([f"{k}: {v}" for k, v in konkrete_auswahlen.items()]) if konkrete_auswahlen else "Keines ausgewählt"
        gewaehltes_objekt = list(konkrete_auswahlen.values())[0] if konkrete_auswahlen else None
        
        # AUSFÜHRUNG DER SCHREIB-AKTIVITÄTEN
        if drive_service is not None:
            if st.session_state.aktive_aktion in ["Störung", "Feedback"]:
                with st.spinner("Eintrag wird formatsicher protokolliert..."):
                    if schreibe_input(drive_service, prompt, nutzer_rolle, gewaehltes_objekt, st.session_state.aktive_aktion):
                        df_wissen, _ = load_data_from_drive()

            elif st.session_state.aktive_aktion == "Information":
                with st.spinner("KI analysiert Zielspalte..."):
                    zielspalte = ai_waehle_stammdaten_spalte(prompt)
                    if aktualisiere_stammdaten(drive_service, prompt, nutzer_rolle, gewaehltes_objekt, zielspalte):
                        df_wissen, _ = load_data_from_drive()

            elif st.session_state.aktive_aktion == "Änderung" and nutzer_rolle == "Admin":
                with st.spinner("Struktur wird erweitert..."):
                    kat_text = list(konkrete_auswahlen.keys())[0] if konkrete_auswahlen else "Ausstattung innen"
                    if ändere_struktur(drive_service, prompt, nutzer_rolle, kat_text):
                        df_wissen, _ = load_data_from_drive()

        # CONTEXT-FILTER (TOKEN-SPARGANG) FÜR DIE LESE-AKTIVITÄTEN
        if df_wissen is not None and not df_wissen.empty:
            df_gefiltert = df_wissen.copy()
            if gewaehltes_objekt and gewaehltes_objekt != "Nicht gefunden":
                df_gefiltert = df_gefiltert[df_gefiltert.iloc[:, 0].astype(str).str.strip().str.lower() == str(gewaehltes_objekt).strip().lower()]
            elif konkrete_auswahlen:
                mask = pd.Series(False, index=df_gefiltert.index)
                for kat in konkrete_auswahlen.keys():
                    if "innen" in kat.lower(): mask = mask | df_gefiltert.iloc[:, 1].astype(str).str.contains("innen", case=False, na=False)
                    elif "außen" in kat.lower() or "aussen" in kat.lower(): mask = mask | df_gefiltert.iloc[:, 1].astype(str).str.contains("außen|aussen", case=False, na=False)
                    elif "nähe" in kat.lower() or "naehe" in kat.lower(): mask = mask | df_gefiltert.iloc[:, 1].astype(str).str.contains("nähe|naehe|In der Nähe", case=False, na=False)
                    else: mask = mask | df_gefiltert.iloc[:, 1].astype(str).str.contains(kat, case=False, na=False)
                df_gefiltert = df_gefiltert[mask]
            
            if df_gefiltert.empty: df_gefiltert = df_wissen
            kontext = f"\n\nAktuelle Daten aus der Wissensbasis:\n{df_gefiltert.to_string(index=False)}"
        else:
            kontext = ""
        
        with st.chat_message("assistant"):
            with st.spinner("Villa Avatar überlegt..."):
                antwort_text = generate_ki_response(
                    f"Rolle: {nutzer_rolle}\nKontext-Aktion des Nutzers: {st.session_state.aktive_aktion}\n"
                    f"Gewählte(s) HMI-Objekt(e): {gewaehlte_objekte_str}\nAnfrage: {prompt} {kontext}"
                )
            st.markdown(antwort_text)
            st.session_state.messages.append({"role": "assistant", "content": antwort_text})
