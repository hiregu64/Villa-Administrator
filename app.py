import streamlit as st
from google import genai
from google.genai import types
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import pandas as pd
import io
import datetime
import openpyxl
from openpyxl.styles import Font

# ==========================================
# KONFIGURATION
# ==========================================
st.set_page_config(page_title="Villa Avatar", page_icon="☀️", layout="centered")
FILE_ID = '1FzhWZuO6aRZkdRuQBzaojhkq7bQDyprl'

# ==========================================
# 1. DYNAMISCHES LADEN: 2-BLATT-ARCHITEKTUR
# ==========================================
@st.cache_data(ttl=30)  
def load_dynamic_data():
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
        # Lese beide Tabellenblätter
        df_wissen = pd.read_excel(fh, sheet_name="Wissensbasis")
        fh.seek(0)
        df_lexikon = pd.read_excel(fh, sheet_name="Spalten_Lexikon")
        
        # 'Wo?' (Kategorie) vorwärts auffüllen für Dropdowns
        if df_wissen is not None and not df_wissen.empty and "Wo?" in df_wissen.columns:
            df_wissen["Wo?"] = df_wissen["Wo?"].ffill()
            
        return df_wissen, df_lexikon, service
    except Exception as e:
        st.error(f"Fehler beim Laden der Excel-Matrix: {e}")
        return None, None, None

with st.spinner("Initialisiere dynamische Matrix aus Google Drive..."):
    df_wissen, df_lexikon, drive_service = load_dynamic_data()

# ==========================================
# 2. DYNAMISCHE SCHREIB-ENGINE (NAMENSBASIERT)
# ==========================================
def dynamic_write_to_excel(service, text, nutzername, objekt_name, action_type, df_lexikon_current):
    try:
        request = service.files().get_media(fileId=FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done: _, done = downloader.next_chunk()
        fh.seek(0)
        
        wb = openpyxl.load_workbook(fh)
        ws = wb["Wissensbasis"]
        
        # Header auslesen um Spalten dynamisch zu finden
        headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
        
        # Zeile (Objekt) finden
        row_idx = None
        col_bez_idx = headers.index("Bezeichnung") + 1 if "Bezeichnung" in headers else 1
        
        if objekt_name and str(objekt_name).strip().lower() != "nicht gefunden":
            for r in range(2, ws.max_row + 1):
                val = ws.cell(row=r, column=col_bez_idx).value
                if val and str(val).strip().lower() == str(objekt_name).strip().lower():
                    row_idx = r
                    break
                    
        # Fallback auf "Nicht gefunden"
        if row_idx is None:
            for r in range(2, ws.max_row + 1):
                val = ws.cell(row=r, column=col_bez_idx).value
                if val and "nicht gefunden" in str(val).strip().lower():
                    row_idx = r
                    break
        if row_idx is None:
            row_idx = ws.max_row + 1
            ws.cell(row=row_idx, column=col_bez_idx, value="Nicht gefunden")

        # Zielspalten ermitteln je nach Action Type
        t_col, status_col, status_val = None, None, None
        
        if action_type == "Störung":
            t_col = headers.index("Störung [Input]") + 1 if "Störung [Input]" in headers else None
            status_col = headers.index("Störung Status") + 1 if "Störung Status" in headers else None
            status_val = "aktiv"
        elif action_type == "Feedback":
            t_col = headers.index("Feedback [Input]") + 1 if "Feedback [Input]" in headers else None
            status_col = headers.index("Feedback Status") + 1 if "Feedback Status" in headers else None
            status_val = "offen"
        elif action_type == "Keine Information":
            t_col = headers.index("Keine Information") + 1 if "Keine Information" in headers else None
            status_col = headers.index("Keine Information Status") + 1 if "Keine Information Status" in headers else None
            status_val = "offen"
        elif action_type == "Information":
            # Spezifisch für Host: Nutzt AI um Zielspalte zu finden
            mögliche_spalten = df_lexikon_current['Spaltenname'].tolist() if df_lexikon_current is not None else []
            zielspalte_name = ai_waehle_stammdaten_spalte(text, mögliche_spalten)
            t_col = headers.index(zielspalte_name) + 1 if zielspalte_name in headers else (headers.index("Details Nutzung [Output]") + 1 if "Details Nutzung [Output]" in headers else 8)
        
        if t_col is None:
            st.error(f"Konnte die Zielspalte für '{action_type}' in der Excel-Tabelle nicht finden.")
            return False

        zeitstempel = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
        
        # Wert in Hauptspalte schreiben
        alt_text = str(ws.cell(row=row_idx, column=t_col).value) if ws.cell(row=row_idx, column=t_col).value is not None else ""
        neu_text = f"{alt_text}\n- [{zeitstempel} | {nutzername}]: {text}".strip() if alt_text else f"- [{zeitstempel} | {nutzername}]: {text}"
        cell_text = ws.cell(row=row_idx, column=t_col, value=neu_text)
        cell_text.font = Font(color="0000FF")

        # Status schreiben (Falls vorhanden, bei 'Information' z.B. nicht)
        if status_col is not None and status_val is not None:
            alt_status = str(ws.cell(row=row_idx, column=status_col).value) if ws.cell(row=row_idx, column=status_col).value is not None else ""
            neu_status = f"{alt_status}\n- [{zeitstempel} | {nutzername}]: {status_val}".strip() if alt_status else f"- [{zeitstempel} | {nutzername}]: {status_val}"
            cell_status = ws.cell(row=row_idx, column=status_col, value=neu_status)
            cell_status.font = Font(color="0000FF")
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        media = MediaIoBaseUpload(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        service.files().update(fileId=FILE_ID, media_body=media).execute()
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Fehler in dynamic_write_to_excel(): {e}")
        return False

# ==========================================
# 3. KI-GEHIRN & FALLBACK
# ==========================================
VILLA_PROMPT = """
Du bist „Villa Avatar“, der digitale Helfer für Gäste (Gast) und Gastgeber (Host) der Villa. 

VERHALTEN:
- Antworte immer kurz, freundlich und smartphone-optimiert.
- Passe deine Tonalität der Rolle an.
- Nutze für Antworten AUSSCHLIESSLICH die übergebenen Tabellendaten und beachte die Regeln im 'Spalten_Lexikon'.
- Wenn die Daten keine Antwort enthalten: Erfinde nichts! Antworte: „Dazu liegen mir aktuell leider keine Informationen vor.“
- ABSOLUTES VERBOT: Erwähne niemals Tabellenstrukturen, Dateinamen, Zeilen oder Spaltennamen gegenüber dem Nutzer. Verhalte dich wie ein Mensch, der dieses Wissen im Kopf hat.
"""

@st.cache_resource
def get_ki_client():
    if "GEMINI_API_KEY" in st.secrets:
        return genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
    return None

client = get_ki_client()

def ai_waehle_stammdaten_spalte(user_text, verfuegbare_spalten):
    if client is None or not verfuegbare_spalten: return "Details Nutzung [Output]"
    try:
        spalten_str = "\n".join([f"- {s}" for s in verfuegbare_spalten])
        prompt = f"""
        Analysiere diesen Host-Text: "{user_text}"
        In welche dieser Excel-Spalten gehört die Info am ehesten?
        {spalten_str}
        Antworte NUR mit dem exakten Spaltennamen.
        """
        response = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        return response.text.strip().replace('"', '').replace("'", "")
    except:
        return "Details Nutzung [Output]"

def generate_ki_response(prompt_text):
    if client is None: 
        return "🛑 KI-Dienst nicht konfiguriert."
    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash", 
            contents=prompt_text, 
            config=types.GenerateContentConfig(system_instruction=VILLA_PROMPT)
        )
        return response.text
    except Exception as e:
        if any(err in str(e) for err in ["429", "RESOURCE_EXHAUSTED", "503", "UNAVAILABLE", "504", "TIMEOUT"]):
            try:
                res = client.models.generate_content(model="gemini-2.0-flash", contents=prompt_text, config=types.GenerateContentConfig(system_instruction=VILLA_PROMPT))
                return res.text
            except:
                return "🛑 KI Server überlastet oder Limit erreicht. Bitte kurz warten."
        return f"⚠️ Fehler: {e}"

# ==========================================
# 4. UI & HMI-MATRIX (STRIKT 2 ROLLEN)
# ==========================================
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
st.markdown("Hallo! Ich bin Villa Avatar. Wähle unten deine Rolle aus, um zu beginnen.")

STANDARD_DROPDOWNS = ["Ausstattung innen", "Ausstattung außen", "In der Nähe"]

HMI_MATRIX = {
    "Gast": {
        "Hilfe": {"text": "Wobei kann ich dir helfen?", "dd": STANDARD_DROPDOWNS},
        "Feedback": {"text": "Welches Feedback hast du für uns?", "dd": STANDARD_DROPDOWNS},
        "Störung": {"text": "Was ist passiert oder defekt?", "dd": STANDARD_DROPDOWNS},
        "Keine Information": {"text": "Welche Information hat dir hier gefehlt?", "dd": STANDARD_DROPDOWNS}
    },
    "Host": {
        "Hilfe": {"text": "Wobei kann ich dir helfen?", "dd": STANDARD_DROPDOWNS},
        "Information": {"text": "Gern nehme ich deine Notizen auf und ordne sie zu.", "dd": STANDARD_DROPDOWNS},
        "Feedback": {"text": "Welches Feedback dokumentieren wir?", "dd": STANDARD_DROPDOWNS},
        "Störung": {"text": "Was ist passiert?", "dd": STANDARD_DROPDOWNS},
        "Keine Information": {"text": "Welche Wissenslücke sollen wir protokollieren?", "dd": STANDARD_DROPDOWNS},
        "Bericht": {"text": "Nenne mir bitte den Zeitraum und das Thema.", "dd": []}
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

nutzer_rolle = st.selectbox("Rolle", options=["Gast", "Host"], index=None, placeholder="Wer bist du?", label_visibility="collapsed")

if nutzer_rolle != st.session_state.vorherige_rolle:
    st.session_state.vorherige_rolle = nutzer_rolle
    st.session_state.aktive_aktion = None
    st.session_state.messages = []  
    st.rerun()

if nutzer_rolle is not None:
    st.write("---")
    
    with st.container():
        st.markdown(
            "<div style='display: flex; justify-content: flex-end; align-items: center; gap: 8px; margin-bottom: 10px;'>"
            "<span style='font-weight: bold; font-size: 1.2rem;'>Mein Anliegen:</span>"
            "<div style='width: 32px; height: 32px; background-color: rgb(255, 75, 75); border-radius: 8px; display: flex; align-items: center; justify-content: center;'>"
            "<svg viewBox='0 0 24 24' width='20' height='20' stroke='white' stroke-width='2' fill='none' stroke-linecap='round' stroke-linejoin='round'>"
            "<circle cx='12' cy='12' r='10'></circle><path d='M8 14s1.5 2 4 2 4-2 4-2'></path><line x1='9' y1='9' x2='9.01' y2='9'></line><line x1='15' y1='9' x2='15.01' y2='9'></line>"
            "</svg></div></div>", unsafe_allow_html=True
        )
    
    if nutzer_rolle == "Gast":
        col1, col2 = st.columns(2)
        col3, col4 = st.columns(2)
        with col1:
            if st.button("Ich brauche Hilfe.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Hilfe" else "secondary"): handle_button_click("Hilfe")
        with col2:
            if st.button("Es gibt eine Störung.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Störung" else "secondary"): handle_button_click("Störung")
        with col3:
            if st.button("Ich möchte Feedback geben.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Feedback" else "secondary"): handle_button_click("Feedback")
        with col4:
            if st.button("Mir fehlt eine Information.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Keine Information" else "secondary"): handle_button_click("Keine Information")
                
    elif nutzer_rolle == "Host":
        col1, col2, col3 = st.columns(3)
        col4, col5, col6 = st.columns(3)
        with col1:
            if st.button("Ich brauche Hilfe.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Hilfe" else "secondary"): handle_button_click("Hilfe")
        with col2:
            if st.button("Ich habe neue Infos.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Information" else "secondary"): handle_button_click("Information")
        with col3:
            if st.button("Bericht erstellen.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Bericht" else "secondary"): handle_button_click("Bericht")
        with col4:
            if st.button("Störung melden.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Störung" else "secondary"): handle_button_click("Störung")
        with col5:
            if st.button("Feedback erfassen.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Feedback" else "secondary"): handle_button_click("Feedback")
        with col6:
            if st.button("Wissenslücke loggen.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Keine Information" else "secondary"): handle_button_click("Keine Information")

    # Dropdowns generieren
    if st.session_state.aktive_aktion and nutzer_rolle in HMI_MATRIX:
        aktiver_state = HMI_MATRIX[nutzer_rolle].get(st.session_state.aktive_aktion)
        if aktiver_state:
            st.write("")
            with st.chat_message("assistant"): st.markdown(aktiver_state['text'])
            
            kategorien_fuer_rolle = aktiver_state["dd"]
            if df_wissen is not None and not df_wissen.empty:
                bez_spalte = "Bezeichnung" if "Bezeichnung" in df_wissen.columns else df_wissen.columns[0]
                kat_spalte = "Wo?" if "Wo?" in df_wissen.columns else df_wissen.columns[1]

                for kat in kategorien_fuer_rolle:
                    if "innen" in kat.lower(): mask = df_wissen[kat_spalte].astype(str).str.contains("innen", case=False, na=False)
                    elif "außen" in kat.lower() or "aussen" in kat.lower(): mask = df_wissen[kat_spalte].astype(str).str.contains("außen|aussen", case=False, na=False)
                    elif "nähe" in kat.lower() or "naehe" in kat.lower(): mask = df_wissen[kat_spalte].astype(str).str.contains("nähe|naehe|In der Nähe", case=False, na=False)
                    else: mask = df_wissen[kat_spalte].astype(str).str.contains(kat, case=False, na=False)
                    
                    # 2D-Matrix Vor-Filter für Dropdown (Gast sieht nur freigegebene Zeilen)
                    if nutzer_rolle == "Gast" and "Relevanz Gast" in df_wissen.columns:
                        mask = mask & (df_wissen["Relevanz Gast"].astype(str).str.strip().str.lower() == "x")
                    
                    verfuegbare_bez = df_wissen[mask][bez_spalte].dropna().drop_duplicates().tolist()
                    verfuegbare_bez = sorted([str(b).strip() for b in verfuegbare_bez])
                    if "Nicht gefunden" in verfuegbare_bez: verfuegbare_bez.remove("Nicht gefunden")
                    verfuegbare_bez.append("Nicht gefunden")
                    
                    st.selectbox(label=f"Hidden_{kat}", options=verfuegbare_bez, index=None, placeholder=f"📍 {kat} wählen...", key=f"sub_cat_wahl_{kat}_{st.session_state.aktive_aktion}", label_visibility="collapsed")

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
        
        # ==========================================
        # AUSFÜHRUNG DER SCHREIB-AKTIVITÄTEN
        # ==========================================
        if drive_service is not None and st.session_state.aktive_aktion in ["Störung", "Feedback", "Keine Information", "Information"]:
            with st.spinner("Eintrag wird formatsicher in Excel protokolliert..."):
                if dynamic_write_to_excel(drive_service, prompt, nutzer_rolle, gewaehltes_objekt, st.session_state.aktive_aktion, df_lexikon):
                    df_wissen, df_lexikon, _ = load_dynamic_data()

        # ==========================================
        # 2D-MATRIX-FILTER & TOKEN-SPARGANG FÜR KI
        # ==========================================
        if df_wissen is not None and not df_wissen.empty:
            df_gefiltert = df_wissen.copy()
            
            # 1. ACHSE: Zeilen-Filter (Welches Objekt / Welche Kategorie?)
            bez_spalte = "Bezeichnung" if "Bezeichnung" in df_gefiltert.columns else df_gefiltert.columns[0]
            kat_spalte = "Wo?" if "Wo?" in df_gefiltert.columns else df_gefiltert.columns[1]
            
            if gewaehltes_objekt and gewaehltes_objekt != "Nicht gefunden":
                df_gefiltert = df_gefiltert[df_gefiltert[bez_spalte].astype(str).str.strip().str.lower() == str(gewaehltes_objekt).strip().lower()]
            elif konkrete_auswahlen:
                mask = pd.Series(False, index=df_gefiltert.index)
                for kat in konkrete_auswahlen.keys():
                    if "innen" in kat.lower(): mask = mask | df_gefiltert[kat_spalte].astype(str).str.contains("innen", case=False, na=False)
                    elif "außen" in kat.lower() or "aussen" in kat.lower(): mask = mask | df_gefiltert[kat_spalte].astype(str).str.contains("außen|aussen", case=False, na=False)
                    elif "nähe" in kat.lower() or "naehe" in kat.lower(): mask = mask | df_gefiltert[kat_spalte].astype(str).str.contains("nähe|naehe|In der Nähe", case=False, na=False)
                    else: mask = mask | df_gefiltert[kat_spalte].astype(str).str.contains(kat, case=False, na=False)
                df_gefiltert = df_gefiltert[mask]
                
            # 2. ACHSE: Der strikte Gast-Matrix-Filter (Zeile + Spalte)
            if nutzer_rolle == "Gast":
                # Zeilen: Nur mit 'x'
                if "Relevanz Gast" in df_gefiltert.columns:
                    df_gefiltert = df_gefiltert[df_gefiltert["Relevanz Gast"].astype(str).str.strip().str.lower() == "x"]
                
                # Spalten: Nur wenn im Lexikon 'Sichtbar für Gast' == 'ja'
                if df_lexikon is not None and "Sichtbar für Gast" in df_lexikon.columns:
                    erlaubt = df_lexikon[df_lexikon["Sichtbar für Gast"].astype(str).str.strip().str.lower() == "ja"]["Spaltenname"].tolist()
                    erlaubte_spalten = [c for c in df_gefiltert.columns if str(c).strip() in erlaubt]
                    df_gefiltert = df_gefiltert[erlaubte_spalten]

            # Lexikon als Instruktion für die KI aufbereiten
            lexikon_text = ""
            if df_lexikon is not None and not df_lexikon.empty:
                lexikon_text = "REGELN FÜR SPALTEN (Spalten_Lexikon):\n"
                for _, row in df_lexikon.iterrows():
                    spaltenname = str(row.get("Spaltenname", ""))
                    if spaltenname in df_gefiltert.columns:
                        bedeutung = str(row.get("Bedeutung / Beschreibung", ""))
                        format_regel = str(row.get("Erwartetes Format / Regel", ""))
                        lexikon_text += f"- '{spaltenname}': {bedeutung} (Format: {format_regel})\n"

            if df_gefiltert.empty: 
                df_gefiltert = pd.DataFrame(["Keine passenden/freigegebenen Daten gefunden."])
                
            kontext = f"\n\n{lexikon_text}\nAktuelle Daten aus der Wissensbasis:\n{df_gefiltert.to_string(index=False)}"
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
