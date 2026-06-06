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

# ==============================================================================
# CONFIGURATION & INFRASTRUCTURE (Sicht G: Cache & Deployment)
# ==============================================================================
st.set_page_config(page_title="Villa Avatar", page_icon="☀️", layout="centered")
FILE_ID = '1FzhWZuO6aRZkdRuQBzaojhkq7bQDyprl'

# Unverrückbarer Fallback-Satz aus der Master-PPT
FALLBACK_SATZ = "Ich habe dazu leider keine Informationen, Ich gebe das aber gern an die Hosts weiter."

@st.cache_data(ttl=30)  # Sicht G: 30-Sekunden-Taktung für optimalen Performance-Kompromiss
def load_dynamic_data():
    try:
        creds_dict = st.secrets["GOOGLE_CREDENTIALS"]
        creds = service_account.Credentials.from_service_account_info(creds_dict)
        service = build('drive', 'v3', credentials=creds)
        
        request = service.files().get_media(fileId=FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False: _, done = downloader.next_chunk()
            
        fh.seek(0)
        xl = pd.ExcelFile(fh)
        df_wissen = pd.read_excel(fh, sheet_name="Wissensbasis")
        fh.seek(0)
        df_lexikon = pd.read_excel(fh, sheet_name="Spalten_Lexikon")
        
        # Fehlertolerante Datenbasis: Automatische Vererbung der Kategorien
        if df_wissen is not None and not df_wissen.empty and "Wo?" in df_wissen.columns:
            df_wissen["Wo?"] = df_wissen["Wo?"].ffill()
            
        return df_wissen, df_lexikon, service
    except Exception as e:
        st.error(f"Fehler beim Laden der Excel-Matrix: {e}")
        return None, None, None

with st.spinner("Initialisiere dynamische Matrix..."):
    df_wissen, df_lexikon, drive_service = load_dynamic_data()

# ==============================================================================
# DATA ENGINE - WRITING SCRIPT (Sicht A: Mini-Workflows)
# ==============================================================================
def find_column_by_fuzzy_name(headers, target_name):
    cleaned_headers = [str(h).strip().lower().replace("\n", " ") for h in headers]
    search = str(target_name).strip().lower()
    if search in cleaned_headers:
        return cleaned_headers.index(search) + 1
    for idx, h in enumerate(cleaned_headers):
        if search in h:
            return idx + 1
    return None

def dynamic_write_to_excel(service, text, nutzername, objekt_name, system_action, df_lexikon_current):
    try:
        request = service.files().get_media(fileId=FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done: _, done = downloader.next_chunk()
        fh.seek(0)
        
        wb = openpyxl.load_workbook(fh)
        ws = wb["Wissensbasis"]
        headers = [str(c.value) if c.value else "" for c in ws[1]]
        
        col_bez_idx = find_column_by_fuzzy_name(headers, "Bezeichnung") or 1
        row_idx = None
        
        if objekt_name and str(objekt_name).strip().lower() != "nicht gefunden":
            for r in range(2, ws.max_row + 1):
                val = ws.cell(row=r, column=col_bez_idx).value
                if val and str(val).strip().lower() == str(objekt_name).strip().lower():
                    row_idx = r
                    break
                    
        if row_idx is None:
            for r in range(2, ws.max_row + 1):
                val = ws.cell(row=r, column=col_bez_idx).value
                if val and "nicht gefunden" in str(val).strip().lower():
                    row_idx = r
                    break
                    
        if row_idx is None:
            row_idx = ws.max_row + 1
            ws.cell(row=row_idx, column=col_bez_idx, value="Nicht gefunden")

        t_col, status_col, status_val = None, None, None
        
        # Sicht A: Mini-Workflow Zuordnungen und Statusbefehle
        if system_action == "Störung":
            t_col = find_column_by_fuzzy_name(headers, "Störung")
            status_col = find_column_by_fuzzy_name(headers, "Störung Status")
            status_val = "aktiv"
        elif system_action == "Feedback":
            t_col = find_column_by_fuzzy_name(headers, "Feedback")
            status_col = find_column_by_fuzzy_name(headers, "Feedback Status")
            status_val = "offen"
        elif system_action == "Keine Information":
            t_col = find_column_by_fuzzy_name(headers, "Keine Information")
            status_col = find_column_by_fuzzy_name(headers, "Keine Information Status")
            status_val = "offen"
        elif system_action == "Information":
            mögliche_spalten = df_lexikon_current['Spaltenname'].tolist() if df_lexikon_current is not None else []
            zielspalte_name = ai_waehle_stammdaten_spalte(text, mögliche_spalten)
            t_col = headers.index(zielspalte_name) + 1 if zielspalte_name in headers else 8
        
        if t_col is None: return False

        zeitstempel = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
        alt_text = str(ws.cell(row=row_idx, column=t_col).value) if ws.cell(row=row_idx, column=t_col).value is not None else ""
        neu_text = f"{alt_text}\n- [{zeitstempel} | {nutzername}]: {text}".strip() if alt_text else f"- [{zeitstempel} | {nutzername}]: {text}"
        ws.cell(row=row_idx, column=t_col, value=neu_text).font = Font(color="0000FF")

        if status_col is not None and status_val is not None:
            alt_status = str(ws.cell(row=row_idx, column=status_col).value) if ws.cell(row=row_idx, column=status_col).value is not None else ""
            neu_status = f"{alt_status}\n- [{zeitstempel} | {nutzername}]: {status_val}".strip() if alt_status else f"- [{zeitstempel} | {nutzername}]: {status_val}"
            ws.cell(row=row_idx, column=status_col, value=neu_status).font = Font(color="0000FF")
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        media = MediaIoBaseUpload(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        service.files().update(fileId=FILE_ID, media_body=media).execute()
        return True
    except Exception as e:
        st.error(f"Fehler in Daten-Engine: {e}")
        return False

# ==============================================================================
# KI-CORE ENGINE (Sicht E: Globaler Prompt & Systemidentität)
# ==============================================================================
VILLA_PROMPT = """Du bist „Villa Avatar“, der digitale Helfer für die Gäste (Gast), Gastgeber (Host) und Administratoren (Admin) der Villa. 
Antworte immer kurz, freundlich, präzise und smartphone-optimiert. Beziehe dich exakt auf die mitgegebenen Live-Daten. 
Falls im Kontext keine Daten vorhanden sind oder eine Wissenslücke vorliegt, erfinde NIEMALS Fakten.
ABSOLUTES VERBOT: Erwähne NIEMALS interne Dateinamen, Spaltenbezeichnungen oder Excel-Strukturen."""

@st.cache_resource
def get_ki_client():
    if "GEMINI_API_KEY" in st.secrets: return genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
    return None
client = get_ki_client()

def ai_waehle_stammdaten_spalte(user_text, verfuegbare_spalten):
    if client is None or not verfuegbare_spalten: return "Details Nutzung [Output]"
    try:
        spalten_str = "\n".join([f"- {s}" for s in verfuegbare_spalten])
        prompt = f'Analysiere: "{user_text}"\nIn welche Spalte gehört das?\n{spalten_str}\nAntworte NUR mit dem exakten Spaltennamen.'
        response = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        return response.text.strip().replace('"', '').replace("'", "")
    except: return "Details Nutzung [Output]"

def generate_ki_response(prompt_text):
    if client is None: return "🛑 KI-Schnittstelle nicht konfiguriert."
    try:
        return client.models.generate_content(model="gemini-2.5-flash", contents=prompt_text, config=types.GenerateContentConfig(system_instruction=VILLA_PROMPT)).text
    except:
        return "🛑 KI temporär nicht erreichbar."

# ==============================================================================
# HMI PRESENTATION LAYER (Sicht E: Asymmetrisches App-Design)
# ==============================================================================
st.markdown("""
    <style>
    div.stButton > button[kind="primary"] { background-color: #e3f2fd !important; color: #1565c0 !important; border: 1px solid #bbdefb !important; font-weight: bold !important; }
    div.stButton > button[kind="primary"]:hover { background-color: #bbdefb !important; border: 1px solid #64b5f6 !important; }
    
    /* Sicht E: Asymmetrischer Chatverlauf (Nutzer rechtsbündig/grau, KI linksbündig) */
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
st.markdown("Hallo! Ich bin Villa Avatar. Wähle unten deine Rolle aus, um zu beginnen.")

STANDARD_DROPDOWNS = ["Ausstattung innen", "Ausstattung außen", "In der Nähe"]

HMI_MATRIX = {
    "Gast": {
        "Ich brauche Hilfe.": {"sys_action": "Hilfe", "reply": "Wobei kann ich dir helfen?", "dd": STANDARD_DROPDOWNS},
        "Ich möchte eine Störung melden.": {"sys_action": "Störung", "reply": "Was ist passiert oder defekt? Ich kümmere mich darum.", "dd": STANDARD_DROPDOWNS},
        "Ich möchte Feedback geben.": {"sys_action": "Feedback", "reply": "Welches Feedback hast du für uns? Erzähl mir davon.", "dd": STANDARD_DROPDOWNS}
    },
    "Host": {
        "Ich brauche Hilfe.": {"sys_action": "Hilfe", "reply": "Wobei kann ich dir helfen?", "dd": STANDARD_DROPDOWNS},
        "Ich habe neue Informationen.": {"sys_action": "Information", "reply": "Gern nehme ich deine Notizen auf und ordne sie zu.", "dd": STANDARD_DROPDOWNS},
        "Ich benötigt einen Bericht.": {"sys_action": "Bericht", "reply": "Welchen Bericht möchtest du einsehen? Wähle per Button:", "dd": []}, # Sicht D: Keine Standard-Dropdowns
        "Ich möchte eine Störung melden.": {"sys_action": "Störung", "reply": "Was ist passiert? Ich halte es fest.", "dd": STANDARD_DROPDOWNS},
        "Ich möchte Feedback geben.": {"sys_action": "Feedback", "reply": "Welches Feedback dokumentieren wir?", "dd": STANDARD_DROPDOWNS}
    }
}

if "messages" not in st.session_state: st.session_state.messages = []
if "aktive_aktion" not in st.session_state: st.session_state.aktive_aktion = None
if "vorherige_rolle" not in st.session_state: st.session_state.vorherige_rolle = None
if "bericht_filter" not in st.session_state: st.session_state.bericht_filter = None

def handle_button_click(aktions_satz):
    for key in list(st.session_state.keys()):
        if key.startswith("sub_cat_wahl_"): del st.session_state[key]
    st.session_state.aktive_aktion = aktions_satz
    st.session_state.bericht_filter = None
    st.session_state.messages = []  
    st.rerun()

nutzer_rolle = st.selectbox("Rolle", options=["Gast", "Host"], index=None, placeholder="Wer bist du?", label_visibility="collapsed")

if nutzer_rolle != st.session_state.vorherige_rolle:
    st.session_state.vorherige_rolle = nutzer_rolle
    st.session_state.aktive_aktion = None
    st.session_state.bericht_filter = None
    st.session_state.messages = []  
    st.rerun()

if nutzer_rolle is not None:
    st.write("---")
    
    # HMI Hauptanliegen
    if nutzer_rolle == "Gast":
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("Ich brauche Hilfe.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Ich brauche Hilfe." else "secondary"): handle_button_click("Ich brauche Hilfe.")
        with col2:
            if st.button("Ich möchte eine Störung melden.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Ich möchte eine Störung melden." else "secondary"): handle_button_click("Ich möchte eine Störung melden.")
        with col3:
            if st.button("Ich möchte Feedback geben.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Ich möchte Feedback geben." else "secondary"): handle_button_click("Ich möchte Feedback geben.")
                
    elif nutzer_rolle == "Host":
        col1, col2, col3 = st.columns(3)
        col4, col5 = st.columns(2)
        with col1:
            if st.button("Ich brauche Hilfe.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Ich brauche Hilfe." else "secondary"): handle_button_click("Ich brauche Hilfe.")
        with col2:
            if st.button("Ich habe neue Informationen.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Ich habe neue Informationen." else "secondary"): handle_button_click("Ich habe neue Informationen.")
        with col3:
            if st.button("Ich benötigt einen Bericht.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Ich benötigt einen Bericht." else "secondary"): handle_button_click("Ich benötigt einen Bericht.")
        with col4:
            if st.button("Ich möchte eine Störung melden.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Ich möchte eine Störung melden." else "secondary"): handle_button_click("Ich möchte eine Störung melden.")
        with col5:
            if st.button("Ich möchte Feedback geben.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Ich möchte Feedback geben." else "secondary"): handle_button_click("Ich möchte Feedback geben.")

    if st.session_state.aktive_aktion and nutzer_rolle in HMI_MATRIX:
        cfg = HMI_MATRIX[nutzer_rolle].get(st.session_state.aktive_aktion)
        if cfg:
            st.write("")
            with st.chat_message("assistant"): st.markdown(cfg['reply'])
            
            # ==================================================================
            # SICHT D: HANDLING USE CASE "BERICHT" (Buttons statt Dropdowns)
            # ==================================================================
            if cfg["sys_action"] == "Bericht":
                b_col1, b_col2 = st.columns(2)
                b_col3, b_col4 = st.columns(2)
                b_col5, b_col6 = st.columns(2)
                
                with b_col1:
                    if st.button("⚠️ Offene Störungen", use_container_width=True): st.session_state.bericht_filter = "offene_stoerungen"
                with b_col2:
                    if st.button("✅ Behobene Störungen", use_container_width=True): st.session_state.bericht_filter = "behobene_stoerungen"
                with b_col3:
                    if st.button("💡 Offenes Feedback", use_container_width=True): st.session_state.bericht_filter = "offenes_feedback"
                with b_col4:
                    if st.button("❌ Ignoriertes Feedback", use_container_width=True): st.session_state.bericht_filter = "ignoriertes_feedback"
                with b_col5:
                    if st.button("🔍 Offene Wissenslücken", use_container_width=True): st.session_state.bericht_filter = "offene_luecken"
                with b_col6:
                    if st.button("📋 Gesamtübersicht", use_container_width=True): st.session_state.bericht_filter = "gesamtuebersicht"
            
            # ==================================================================
            # SICHT B: DROPDOWN LOGIK (Für Hilfe & Inputs optional)
            # ==================================================================
            else:
                kategorien_fuer_rolle = cfg["dd"]
                if df_wissen is not None and not df_wissen.empty:
                    bez_spalte = "Bezeichnung" if "Bezeichnung" in df_wissen.columns else df_wissen.columns[0]
                    kat_spalte = "Wo?" if "Wo?" in df_wissen.columns else df_wissen.columns[1]

                    for kat in kategorien_fuer_rolle:
                        if "innen" in kat.lower(): mask = df_wissen[kat_spalte].astype(str).str.contains("innen", case=False, na=False)
                        elif "außen" in kat.lower() or "aussen" in kat.lower(): mask = df_wissen[kat_spalte].astype(str).str.contains("außen|aussen", case=False, na=False)
                        else: mask = df_wissen[kat_spalte].astype(str).str.contains("nähe|naehe|In der Nähe", case=False, na=False)
                        
                        if nutzer_rolle == "Gast" and "Relevanz Gast" in df_wissen.columns:
                            mask = mask & (df_wissen["Relevanz Gast"].astype(str).str.strip().str.lower() == "x")
                        
                        verfuegbare_bez = df_wissen[mask][bez_spalte].dropna().drop_duplicates().tolist()
                        verfuegbare_bez = sorted([str(b).strip() for b in verfuegbare_bez])
                        if "Nicht gefunden" in verfuegbare_bez: verfuegbare_bez.remove("Nicht gefunden")
                        verfuegbare_bez.append("Nicht gefunden")
                        
                        st.selectbox(label=f"Hidden_{kat}", options=verfuegbare_bez, index=None, placeholder=f"📍 {kat} wählen...", key=f"sub_cat_wahl_{kat}_{st.session_state.aktive_aktion}", label_visibility="collapsed")

# ==============================================================================
# CHAT & PROCESSING LAYER (Sicht F: Integrierte Spracheingabe-Philosophie)
# ==============================================================================
st.write("---")
for message in st.session_state.messages:
    with st.chat_message(message["role"]): st.markdown(message["content"])

# Trigger-Weiche: Entweder Freitext-Eingabe ODER Klick auf einen Bericht-Button
prompt = st.chat_input("Bitte schreibe hier oder sprich mit mir 🎙️")
bericht_trigger = (nutzer_rolle == "Host" and st.session_state.aktive_aktion == "Ich benötigt einen Bericht." and st.session_state.bericht_filter is not None)

if prompt or bericht_trigger:
    if nutzer_rolle is None or not st.session_state.aktive_aktion: 
        st.warning("Bitte wähle zuerst deine Rolle und dein Anliegen aus!")
    else:
        cfg = HMI_MATRIX[nutzer_rolle][st.session_state.aktive_aktion]
        aktuelle_sys_action = cfg["sys_action"]
        
        # ----------------------------------------------------------------------
        # HANDLER: USE CASE "BERICHT" (Sicht D)
        # ----------------------------------------------------------------------
        if aktuelle_sys_action == "Bericht":
            with st.chat_message("assistant"):
                with st.spinner("Generiere Host-Bericht aus Zentralmatrix..."):
                    df_report = df_wissen.copy() if df_wissen is not None else pd.DataFrame()
                    report_text = ""
                    
                    if not df_report.empty:
                        # Spalten-Fuzzy-Abgleiche
                        s_status_col = next((c for c in df_report.columns if "störung status" in c.lower()), None)
                        s_input_col = next((c for c in df_report.columns if "störung" in c.lower() and "status" not in c.lower()), None)
                        f_status_col = next((c for c in df_report.columns if "feedback status" in c.lower()), None)
                        f_input_col = next((c for c in df_report.columns if "feedback" in c.lower() and "status" not in c.lower()), None)
                        l_status_col = next((c for c in df_report.columns if "keine information status" in c.lower()), None)
                        l_input_col = next((c for c in df_report.columns if "keine information" in c.lower() and "status" not in c.lower()), None)
                        bez_col = "Bezeichnung" if "Bezeichnung" in df_report.columns else df_report.columns[0]
                        
                        # Filter-Auswertung gemäß Spezifikation
                        f_type = st.session_state.bericht_filter
                        if f_type == "offene_stoerungen" and s_status_col:
                            df_report = df_report[df_report[s_status_col].astype(str).str.contains("aktiv", case=False, na=False)]
                        elif f_type == "behobene_stoerungen" and s_status_col:
                            df_report = df_report[df_report[s_status_col].astype(str).str.contains("OK", case=False, na=False)]
                        elif f_type == "offenes_feedback" and f_status_col:
                            df_report = df_report[df_report[f_status_col].astype(str).str.contains("offen", case=False, na=False)]
                        elif f_type == "ignoriertes_feedback" and f_status_col:
                            df_report = df_report[df_report[f_status_col].astype(str).str.contains("Nein", case=False, na=False)]
                        elif f_type == "offene_luecken" and l_status_col:
                            df_report = df_report[df_report[l_status_col].astype(str).str.contains("offen", case=False, na=False)]
                        elif f_type == "gesamtuebersicht":
                            # Behalte alle Zeilen mit Historien-Inhalten
                            cols_to_check = [c for c in [s_input_col, f_input_col, l_input_col] if c is not None]
                            if cols_to_check:
                                mask = pd.Series(False, index=df_report.index)
                                for c in cols_to_check: mask = mask | df_report[c].notna()
                                df_report = df_report[mask]
                        
                        if not df_report.empty:
                            report_text = f"Gefilterte Rohdaten für Host-Bericht ({f_type}):\n" + df_report[[bez_col] + [c for c in [s_input_col, f_input_col, l_input_col] if c is not None]].to_string(index=False)
                    
                    if not report_text or df_report.empty:
                        antwort_text = "Aktuell liegen keine Einträge für diesen Bericht vor. Alles läuft einwandfrei! ☀️"
                    else:
                        ki_prompt = f"Generiere einen strukturierten, chronologischen Bericht für den Host basierend auf diesen Daten:\n{report_text}\nOrdne die Punkte klar nach der Bezeichnung zu und nutze die Zeitstempel."
                        antwort_text = generate_ki_response(ki_prompt)
                    
                    st.markdown(antwort_text)
                    st.session_state.messages.append({"role": "assistant", "content": antwort_text})
            st.session_state.bericht_filter = None

        # ----------------------------------------------------------------------
        # HANDLER: REINE INPUT-SZENARIEN (Störung, Feedback, Information)
        # ----------------------------------------------------------------------
        elif aktuelle_sys_action in ["Störung", "Feedback", "Information"]:
            with st.chat_message("user"): st.markdown(prompt)
            st.session_state.messages.append({"role": "user", "content": prompt})
            
            konkrete_auswahlen = {}
            for kat in cfg["dd"]:
                key = f"sub_cat_wahl_{kat}_{st.session_state.aktive_aktion}"
                if key in st.session_state and st.session_state[key] is not None:
                    konkrete_auswahlen[kat] = st.session_state[key]
            gewaehlte_objekte_str = ", ".join([f"{k}: {v}" for k, v in konkrete_auswahlen.items()]) if konkrete_auswahlen else "Keines ausgewählt"
            gewaehltes_objekt = list(konkrete_auswahlen.values())[0] if konkrete_auswahlen else None

            with st.chat_message("assistant"):
                with st.spinner("Villa Avatar registriert Daten..."):
                    ki_prompt = (
                        f"Bestätige dem {nutzer_rolle} kurz, höflich und smartphone-optimiert, dass sein "
                        f"Input '{prompt}' bezüglich '{gewaehlte_objekte_str}' erfolgreich registriert wurde. "
                        f"Erteile ausdrücklich KEINE inhaltlichen Ratschläge oder Erklärungen aus der Matrix."
                    )
                    antwort_text = generate_ki_response(ki_prompt)
                
                st.markdown(antwort_text)
                st.session_state.messages.append({"role": "assistant", "content": antwort_text})
                
                if drive_service is not None:
                    with st.spinner("Synchronisiere Excel..."):
                        dynamic_write_to_excel(drive_service, prompt, nutzer_rolle, gewaehltes_objekt, aktuelle_sys_action, df_lexikon)
                        st.cache_data.clear()

        # ----------------------------------------------------------------------
        # HANDLER: SUCH- & HILFE-SZENARIEN (Sicht B & Sicht C)
        # ----------------------------------------------------------------------
        else:
            with st.chat_message("user"): st.markdown(prompt)
            st.session_state.messages.append({"role": "user", "content": prompt})
            
            konkrete_auswahlen = {}
            for kat in cfg["dd"]:
                key = f"sub_cat_wahl_{kat}_{st.session_state.aktive_aktion}"
                if key in st.session_state and st.session_state[key] is not None:
                    konkrete_auswahlen[kat] = st.session_state[key]
            gewaehltes_objekt = list(konkrete_auswahlen.values())[0] if intrauterine_device_wahl_ else None # Fallback-Sicherheit
            gewaehltes_objekt = list(konkrete_auswahlen.values())[0] if konkrete_auswahlen else None

            kontext = ""
            wissensluecke_erkannt = False
            aktiv_stoerung_warnung = ""
            
            if df_wissen is not None and not df_wissen.empty:
                df_gefiltert = df_wissen.copy()
                bez_spalte = "Bezeichnung" if "Bezeichnung" in df_gefiltert.columns else df_gefiltert.columns[0]
                kat_spalte = "Wo?" if "Wo?" in df_gefiltert.columns else df_gefiltert.columns[1]
                
                # Sicht B: Die drei Filter-Szenarien
                # Szenario a: Konkretes Item gewählt
                if gewaehltes_objekt and gewaehltes_objekt != "Nicht gefunden":
                    df_gefiltert = df_gefiltert[df_gefiltert[bez_spalte].astype(str).str.strip().str.lower() == str(gewaehltes_objekt).strip().lower()]
                # Szenario b & c: "Nicht gefunden" oder leere Auswahl -> Offene Tabellensuche
                elif konkrete_auswahlen:
                    mask = pd.Series(False, index=df_gefiltert.index)
                    for kat in konkrete_auswahlen.keys():
                        if "innen" in kat.lower(): mask = mask | df_gefiltert[kat_spalte].astype(str).str.contains("innen", case=False, na=False)
                        elif "außen" in kat.lower() or "aussen" in kat.lower(): mask = mask | df_gefiltert[kat_spalte].astype(str).str.contains("außen|aussen", case=False, na=False)
                        else: mask = mask | df_gefiltert[kat_spalte].astype(str).str.contains("nähe|naehe|In der Nähe", case=False, na=False)
                    df_gefiltert = df_gefiltert[mask]
                    
                # Sicht C: Harte Rollen- und Spaltenfilterung für Gäste
                if nutzer_rolle == "Gast":
                    if "Relevanz Gast" in df_gefiltert.columns:
                        df_gefiltert = df_gefiltert[df_gefiltert["Relevanz Gast"].astype(str).str.strip().str.lower() == "x"]
                    if df_lexikon is not None and "Sichtbar für Gast" in df_lexikon.columns:
                        erlaubt = df_lexikon[df_lexikon["Sichtbar für Gast"].astype(str).str.strip().str.lower() == "ja"]["Spaltenname"].tolist()
                        df_gefiltert = df_gefiltert[[c for c in df_gefiltert.columns if str(c).strip() in erlaubt]]

                # Sicht A: Proaktive Störungswarnung auslesen
                s_status_col = next((c for c in df_gefiltert.columns if "störung status" in c.lower()), None)
                if s_status_col and not df_gefiltert.empty:
                    if df_gefiltert[s_status_col].astype(str).str.contains("aktiv", case=False, na=False).any():
                        aktiv_stoerung_warnung = "⚠️ HINWEIS FÜR DIE KI: Für dieses Objekt liegt aktuell eine aktive Störung vor! Informiere den Nutzer zuerst proaktiv darüber!"

                # Deterministischer Wortfilter gegen Halluzinationen
                gesamter_daten_text = df_gefiltert.to_string().lower()
                such_woerter = [w.strip().lower() for w in prompt.split() if len(w.strip()) > 3]
                
                if such_woerter and not any(w in gesamter_daten_text for w in such_woerter):
                    wissensluecke_erkannt = True

                kontext = f"\nAktuelle verifizierte Daten aus der Wissensbasis:\n{df_gefiltert.to_string(index=False)}"
            else:
                wissensluecke_erkannt = True

            with st.chat_message("assistant"):
                # Szenario: Definitiv leere Datenbasis oder Wort-Filter fehlgeschlagen
                if wissensluecke_erkannt or gewaehltes_objekt == "Nicht gefunden":
                    # Sicht C (Typ B): Versuch freies Weltwissen anzuwenden, falls sinnvoll
                    weltwissen_prompt = f"Der Nutzer fragt: '{prompt}'. Wir haben dazu keine Hausordnung-Daten. Kannst du diese Frage mit allgemeinem, nützlichen Weltwissen kurz beantworten (z.B. Wetter, allgemeine Regionstipps)? Wenn JA, antworte. Wenn es eine hausspezifische Frage ist, antworte NUR mit dem Wort 'FALLBACK'."
                    welt_antwort = generate_ki_response(weltwissen_prompt)
                    
                    if "FALLBACK" in welt_antwort.upper() or len(welt_antwort).strip() < 3:
                        # Harter Lücken-Bruch mit Protokollierung
                        st.markdown(FALLBACK_SATZ)
                        st.session_state.messages.append({"role": "assistant", "content": FALLBACK_SATZ})
                        if drive_service is not None:
                            with st.spinner("Dokumentiere Wissenslücke..."):
                                dynamic_write_to_excel(drive_service, prompt, nutzer_rolle, gewaehltes_objekt, "Keine Information", df_lexikon)
                                st.cache_data.clear()
                    else:
                        # Sicht C: Gekennzeichnete freie KI-Information
                        gekennzeichnete_antwort = f"🤖 Freie KI-Information (nicht aus der Hausordnung verifiziert):\n{welt_antwort}"
                        st.markdown(gekennzeichnete_antwort)
                        st.session_state.messages.append({"role": "assistant", "content": gekennzeichnete_antwort})
                
                # Szenario: Daten in Excel gefunden (Typ A)
                else:
                    with st.spinner("Villa Avatar durchsucht Wissensbasis..."):
                        ki_prompt = f"Rolle: {nutzer_rolle}\nAnfrage: {prompt}\n{aktiv_stoerung_warnung}\nKontext:{kontext}"
                        antwort_text = generate_ki_response(ki_prompt)
                    st.markdown(antwort_text)
                    st.session_state.messages.append({"role": "assistant", "content": antwort_text})
