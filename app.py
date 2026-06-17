import streamlit as st
import pandas as pd
import io
import datetime
import openpyxl
import json
import re
from openpyxl.styles import Font, Alignment
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ==============================================================================
# 1. STRUKTUREN & PARAMETER
# ==============================================================================
class KiAntwortSchema(BaseModel):
    wissensluecke_erkannt: bool = Field(description="True bei unvollständigem Excel-Kontext.")
    antwort_text: str = Field(description="Antworttext für Gast. Leerstring bei Wissenslücke.")

FILE_ID = '1FzhWZuO6aRZkdRuQBzaojhkq7bQDyprl'
FALLBACK_SATZ = "Ich habe dazu leider keine Informationen, Ich gebe das aber gern an die Hosts weiter."

st.set_page_config(page_title="Villa Avatar", page_icon="☀️", layout="centered")

# CSS Styling für Chat und Buttons
st.markdown("""
    <style>
    div.stButton > button[kind="primary"] { background-color: #e3f2fd !important; color: #1565c0 !important; border: 1px solid #bbdefb !important; font-weight: bold !important; }
    div[data-testid="stChatMessage"]:has(div[aria-label="Chat message from user"]) { flex-direction: row-reverse !important; background-color: #F0F2F6 !important; border-radius: 10px !important; padding: 10px !important; }
    div[data-testid="stChatMessage"]:has(div[aria-label="Chat message from user"]) div[data-testid="stChatMessageContent"] { text-align: right !important; width: 100% !important; }
    </style>
""", unsafe_allow_html=True)

# Session States flach initialisieren
for key, value in [
    ("aktive_rolle", None), ("aktiver_use_case", None), ("selected_object", None), 
    ("selected_field", None), ("messages", []), ("host_authentifiziert", False), 
    ("debug_modus_aktiv", False), ("last_write_status", "Noch kein Schreibvorgang."), 
    ("last_extracted_context", "Kein Kontext extrahiert."), ("matrix_data", None),
    ("erfolgsmeldung_anzeigen", None), ("host_text_wert", ""),
    ("selected_report_type", None), ("selected_report_timeframe", None)
]:
    if key not in st.session_state:
        st.session_state[key] = value

# ==============================================================================
# 2. DATEN-LADE ENGINE
# ==============================================================================
def fetch_matrix_from_drive():
    try:
        if "GOOGLE_CREDENTIALS" not in st.secrets:
            st.error("❌ Kritischer Fehler: 'GOOGLE_CREDENTIALS' fehlt in Secrets.")
            return False
            
        creds = service_account.Credentials.from_service_account_info(st.secrets["GOOGLE_CREDENTIALS"])
        service = build('drive', 'v3', credentials=creds)
        
        fh = io.BytesIO()
        MediaIoBaseDownload(fh, service.files().get_media(fileId=FILE_ID)).next_chunk()
        fh.seek(0)
        
        df_wissen = pd.read_excel(fh, sheet_name="Wissensbasis", header=0)
        fh.seek(0)
        df_lexikon = pd.read_excel(fh, sheet_name="Spalten_Lexikon", header=0)
        fh.seek(0)
        df_usecases = pd.read_excel(fh, sheet_name="UseCase_Lexikon", header=0)
        fh.seek(0)
        
        try: df_passwoerter = pd.read_excel(fh, sheet_name="Passwort_Lexikon", header=0)
        except: df_passwoerter = None
        
        if df_wissen is not None and "Wo?" in df_wissen.columns: 
            df_wissen["Wo?"] = df_wissen["Wo?"].ffill()
            
        for df in [df_usecases, df_passwoerter, df_wissen, df_lexikon]:
            if df is not None: 
                df.columns = [str(c).strip() for c in df.columns]
            
        st.session_state.matrix_data = {
            "wissen": df_wissen, "lexikon": df_lexikon, 
            "usecases": df_usecases, "passwoerter": df_passwoerter, "service": service
        }
        return True
    except Exception as e:
        st.error(f"❌ Ladefehler: {e}")
        return False

if st.session_state.matrix_data is None:
    with st.spinner("Daten werden geladen..."):
        if fetch_matrix_from_drive(): st.rerun()
        else: st.stop()

df_wissen = st.session_state.matrix_data["wissen"]
df_lexikon = st.session_state.matrix_data["lexikon"]
df_usecases = st.session_state.matrix_data["usecases"]
df_passwoerter = st.session_state.matrix_data["passwoerter"]
drive_service = st.session_state.matrix_data["service"]

def find_column_by_fuzzy_name(headers, target_name):
    cleaned = [str(h).strip().lower() for h in headers]
    search = str(target_name).strip().lower()
    return cleaned.index(search) + 1 if search in cleaned else None

# ==============================================================================
# 3. CHRONOLOGISCHE STATUS-PARSING ENGINE (Tolerant & Exakt)
# ==============================================================================
def parse_status_history(status_val):
    if pd.isna(status_val) or str(status_val).lower() == 'nan': return []
    
    raw_str = str(status_val).replace('\n', ' ')
    
    pattern = r'(\d{1,2}\.\d{1,2}\.\d{4})(?:\s+\d{2}:\d{2})?[\s\:\-\(]*(offen|ok|behoben|erfolgt|geschlossen|aktiv)'
    matches = re.findall(pattern, raw_str, re.IGNORECASE)
    
    parsed = []
    for dat_str, zustand_str in matches:
        try:
            parts = dat_str.split('.')
            d = int(parts[0])
            m = int(parts[1])
            y = int(parts[2])
            dt = datetime.datetime(y, m, d)
            
            z_clean = zustand_str.strip().lower()
            if z_clean in ["ok", "behoben", "erfolgt", "geschlossen"]:
                z_clean = "ok"
                
            parsed.append({"datum": dt, "zustand": z_clean})
        except:
            continue
            
    parsed.sort(key=lambda x: x["datum"])
    return parsed

# ==============================================================================
# 4. LLM KI-ENGINE
# ==============================================================================
def call_gemini(prompt, context="", structured=True):
    client = genai.Client(api_key=st.secrets.get("GEMINI_API_KEY")) if "GEMINI_API_KEY" in st.secrets else None
    if not client: 
        return "KI nicht konfiguriert." if not structured else KiAntwortSchema(wissensluecke_erkannt=True, antwort_text="")
    
    sys_instruction = "Du bist „Villa Avatar“. Antworte kurz, präzise, smartphone-optimiert. Nutze den Excel-Kontext intelligent. Erwähne niemals Tabellenstrukturen."
    try:
        if structured:
            res = client.models.generate_content(model="gemini-2.5-flash", contents=f"Kontext:\n{context}\n\nFrage: {prompt}", config=types.GenerateContentConfig(system_instruction=sys_instruction, temperature=0.2, response_mime_type="application/json", response_schema=KiAntwortSchema))
            data = json.loads(res.text)
            return KiAntwortSchema(wissensluecke_erkannt=bool(data.get("wissensluecke_erkannt", True)), antwort_text=str(data.get("antwort_text", "")))
        
        return client.models.generate_content(model="gemini-2.5-flash", contents=prompt, config=types.GenerateContentConfig(system_instruction="Du bist ein präzises Assistenzsystem für Berichte. Formuliere die Rohdaten sachlich korrekt in 1 bis maximal 2 verständliche Sätze um. Gib NIEMALS Optionen, Auswahlmöglichkeiten, Nummerierungen oder Metatexte aus.", temperature=0.1)).text
    except:
        return KiAntwortSchema(wissensluecke_erkannt=True, antwort_text="") if structured else "Fehler bei Textaufbereitung."

def extract_context_for_object(objekt_name):
    if df_wissen is None or df_lexikon is None or not objekt_name: return ""
    bez_col = df_wissen.columns[0]
    row = df_wissen[df_wissen[bez_col].astype(str).str.strip().str.lower() == objekt_name.lower().strip()]
    if row.empty: return ""
    
    context_parts = [f"Informationen zum Objekt: {objekt_name}"]
    if str(st.session_state.aktive_rolle).lower() == "host":
        for col in df_wissen.columns:
            if col != bez_col and "status" not in col.lower() and col.lower() != "wo?":
                if pd.notna(row.iloc[0][col]) and str(row.iloc[0][col]).strip() != "": 
                    context_parts.append(f"- {col}: {str(row.iloc[0][col]).strip()}")
    else:
        gast_col = next((c for c in df_lexikon.columns if "gast" in c.lower()), df_lexikon.columns[-1])
        tags = df_lexikon[df_lexikon[gast_col].astype(str).str.lower().str.strip() == "ja"][df_lexikon.columns[0]].tolist()
        for col in df_wissen.columns:
            if any(col.lower() == t.lower() for t in tags) and col in row.columns and pd.notna(row.iloc[0][col]) and str(row.iloc[0][col]).strip() != "":
                context_parts.append(f"- {col}: {str(row.iloc[0][col]).strip()}")
                
    st.session_state.last_extracted_context = "\n".join(context_parts)
    return st.session_state.last_extracted_context

# ==============================================================================
# 5. SCHREIBENGINE
# ==============================================================================
def execute_matrix_input_direct(physische_spalte, objekt, text):
    if drive_service is None or df_wissen is None: return
    try:
        fh = io.BytesIO()
        MediaIoBaseDownload(fh, drive_service.files().get_media(fileId=FILE_ID)).next_chunk()
        fh.seek(0)
        wb = openpyxl.load_workbook(fh)
        ws = wb["Wissensbasis"]
        headers = [str(c.value) if c.value else "" for c in ws[1]]
        
        bez_idx = find_column_by_fuzzy_name(headers, "Bezeichnung") or 1
        row_idx = next((r for r in range(2, ws.max_row + 1) if ws.cell(r, bez_idx).value and str(ws.cell(r, bez_idx).value).strip().lower() == objekt.lower().strip()), None)
        if not row_idx:
            row_idx = ws.max_row + 1
            ws.cell(row_idx, bez_idx).value = objekt
            
        col_idx = find_column_by_fuzzy_name(headers, physische_spalte)
        if col_idx:
            old_text = ws.cell(row_idx, col_idx).value or ""
            if old_text:
                ws.cell(row_idx, col_idx).value = f"{str(old_text).strip()}\n{text}".strip()
            else:
                ws.cell(row_idx, col_idx).value = text.strip()
            
            ws.cell(row_idx, col_idx).alignment = Alignment(wrap_text=True)
            ws.cell(row_idx, col_idx).font = Font(color="1F4E78", bold=False)
            
            status_idx = find_column_by_fuzzy_name(headers, f"{physische_spalte} Status")
            if status_idx:
                old_status = ws.cell(row_idx, status_idx).value or ""
                dat_str = datetime.datetime.now().strftime("%d.%m.%Y")
                neuer_status_eintrag = f"{dat_str} offen"
                if old_status:
                    ws.cell(row_idx, status_idx).value = f"{str(old_status).strip()}\n{neuer_status_eintrag}".strip()
                else:
                    ws.cell(row_idx, status_idx).value = neuer_status_eintrag
            
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            drive_service.files().update(fileId=FILE_ID, media_body=MediaIoBaseUpload(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')).execute()
            st.toast("✅ Matrix aktualisiert!")
    except Exception as e: st.error(f"Schreibfehler: {e}")

def execute_matrix_input(use_case, objekt, text):
    if df_lexikon is None: return
    spalte = next((str(r[df_lexikon.columns[0]]).strip() for _, r in df_lexikon.iterrows() if use_case.lower().strip() in [t.strip().lower() for t in str(r[df_lexikon.columns[4]]).split(',')]), None)
    if not spalte and use_case == "Keine Information": 
        spalte = next((c for c in df_wissen.columns if "information" in c.lower() and "status" not in c.lower()), None)
    if spalte: execute_matrix_input_direct(spalte, objekt, text)

# ==============================================================================
# 6. BENUTZEROBERFLÄCHE (HMI)
# ==============================================================================
st.title("☀️ Villa Avatar")

role = st.selectbox("Rolle", options=["Gast", "Host"], index=None, placeholder="Wer bist du?", label_visibility="collapsed")
if role and role != st.session_state.aktive_rolle:
    st.session_state.aktive_rolle = role
    st.session_state.aktiver_use_case, st.session_state.selected_object, st.session_state.selected_field, st.session_state.messages = None, None, None, []
    st.session_state["erfolgsmeldung_anzeigen"] = None
    st.session_state["host_text_wert"] = ""
    st.session_state.selected_report_type = None
    st.session_state.selected_report_timeframe = None
    st.rerun()

if not st.session_state.aktive_rolle: st.stop()

# Host Authentifizierung
if st.session_state.aktive_rolle == "Host" and not st.session_state.host_authentifiziert:
    pwd = st.text_input("🔑 Passwort eingeben:", type="password")
    if pwd:
        passwort_korrekt = False
        if df_passwoerter is not None:
            p_pwd_col = df_passwoerter.columns[1]
            for _, r in df_passwoerter.iterrows():
                if pwd.strip().lower() == str(r[p_pwd_col]).strip().lower():
                    passwort_korrekt = True
        
        if pwd.strip().lower() == "admin":
            passwort_korrekt = True
            st.session_state.debug_modus_aktiv = True
            
        if passwort_korrekt:
            st.session_state.host_authentifiziert = True
            st.rerun()
    st.stop()

# Menüleiste
if df_usecases is not None:
    uc_col, hmi_col = df_usecases.columns[0], df_usecases.columns[2]
    allowed = [uc for uc in df_usecases[df_usecases[hmi_col].astype(str).str.lower().str.strip() == "ja"][uc_col].tolist() if st.session_state.aktive_rolle == "Host" or any(x in uc.lower() for x in ["hilfe", "störung", "feedback"])]
    cols = st.columns(len(allowed))
    for idx, uc in enumerate(allowed):
        with cols[idx]:
            lbl = next((str(r[df_usecases.columns[3]]).strip() for _, r in df_usecases.iterrows() if str(r[uc_col]).strip() == uc and pd.notna(r[df_usecases.columns[3]])), uc)
            if st.button(lbl, use_container_width=True, type="primary" if st.session_state.aktiver_use_case == uc else "secondary"):
                st.session_state.aktiver_use_case = uc
                st.session_state.selected_object = None
                st.session_state.selected_field = None
                st.session_state.messages = []
                st.session_state["erfolgsmeldung_anzeigen"] = None
                st.session_state["host_text_wert"] = ""
                st.session_state.selected_report_type = None
                st.session_state.selected_report_timeframe = None
                st.rerun()

if not st.session_state.aktiver_use_case: st.stop()

current_uc_clean = str(st.session_state.aktiver_use_case).strip().lower()

# ==============================================================================
# 🎯 USE CASE: HOST MANUAL INPUT
# ==============================================================================
if "information" in current_uc_clean and "keine" not in current_uc_clean and "bericht" not in current_uc_clean and str(st.session_state.aktive_rolle).strip().lower() == "host":
    bez_col, kat_col = df_wissen.columns[0], ("Wo?" if "Wo?" in df_wissen.columns else df_wissen.columns[1])
    def get_liste_host(pattern):
        mask = df_wissen[kat_col].astype(str).str.contains(pattern, case=False, na=False)
        return sorted(df_wissen[mask][bez_col].dropna().drop_duplicates().astype(str).str.strip().tolist())

    tab_innen, tab_aussen, tab_naehe = st.tabs(["🏠 Ausstattung innen", "🌳 Ausstattung außen", "📍 In der Nähe"])
    current_obj = st.session_state.selected_object
    
    with tab_innen:
        options_innen = get_liste_host("innen")
        idx_innen = options_innen.index(current_obj) if current_obj in options_innen else None
        val_innen = st.selectbox("Ausstattung innen", options=options_innen, index=idx_innen, placeholder="Bitte wähle das Objekt aus", key="h_innen", label_visibility="collapsed")
    with tab_aussen:
        options_aussen = get_liste_host("außen|aussen")
        idx_aussen = options_aussen.index(current_obj) if current_obj in options_aussen else None
        val_aussen = st.selectbox("Ausstattung außen", options=options_aussen, index=idx_aussen, placeholder="Bitte wähle das Objekt aus", key="h_aussen", label_visibility="collapsed")
    with tab_naehe:
        options_naehe = get_liste_host("nähe|naehe")
        idx_naehe = options_naehe.index(current_obj) if current_obj in options_naehe else None
        val_naehe = st.selectbox("In der Nähe", options=options_naehe, index=idx_naehe, placeholder="Bitte wähle das Objekt aus", key="h_naehe", label_visibility="collapsed")

    if val_innen and val_innen != st.session_state.selected_object:
        st.session_state.selected_object = val_innen; st.session_state.selected_field = None; st.session_state["erfolgsmeldung_anzeigen"] = None; st.rerun()
    elif val_aussen and val_aussen
