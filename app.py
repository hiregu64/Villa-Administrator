import streamlit as st
import pandas as pd
import io
import datetime
import openpyxl
import json
from openpyxl.styles import Font, Alignment
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ==============================================================================
# 1. STRUCTURATED OUTPUT SCHEMA
# ==============================================================================
class KiAntwortSchema(BaseModel):
    wissensluecke_erkannt: bool = Field(
        description="Muss zwingend True sein, wenn der bereitgestellte Excel-Kontext die Frage nicht direkt beantwortet oder unvollständig ist. False, wenn die Antwort im Text existiert."
    )
    antwort_text: str = Field(
        description="Die kurze Antwort an den Gast. Bleibe streng bei den Fakten. ACHTUNG: Wenn wissensluecke_erkannt True ist, MUSS dieses Feld absolut LEER bleiben (Leerstring ''). Formuliere NIEMALS eine eigene Absage!"
    )

# ==============================================================================
# 2. GLOBAL CONFIGURATION & HMI PRESENTATION LAYER
# ==============================================================================
st.set_page_config(page_title="Villa Avatar", page_icon="☀️", layout="centered")
FILE_ID = '1FzhWZuO6aRZkdRuQBzaojhkq7bQDyprl'
FALLBACK_SATZ = "Ich habe dazu leider keine Informationen, Ich gebe das aber gern an die Hosts weiter."

st.markdown("""
    <style>
    div.stButton > button[kind="primary"] { background-color: #e3f2fd !important; color: #1565c0 !important; border: 1px solid #bbdefb !important; font-weight: bold !important; }
    div[data-testid="stChatMessage"]:has(div[aria-label="Chat message from user"]) { 
        flex-direction: row-reverse !important; background-color: #F0F2F6 !important; border-radius: 10px !important; padding: 10px !important; 
    }
    div[data-testid="stChatMessage"]:has(div[aria-label="Chat message from user"]) div[data-testid="stChatMessageContent"] { text-align: right !important; width: 100% !important; }
    </style>
""", unsafe_allow_html=True)

# Session State Initialisierung
if "aktive_rolle" not in st.session_state: st.session_state.aktive_rolle = None
if "aktiver_use_case" not in st.session_state: st.session_state.aktiver_use_case = None
if "selected_object" not in st.session_state: st.session_state.selected_object = None
if "messages" not in st.session_state: st.session_state.messages = []
if "host_authentifiziert" not in st.session_state: st.session_state.host_authentifiziert = False
if "debug_modus_aktiv" not in st.session_state: st.session_state.debug_modus_aktiv = False
if "last_write_status" not in st.session_state: st.session_state.last_write_status = "Noch kein Schreibvorgang ausgelöst."
if "last_extracted_context" not in st.session_state: st.session_state.last_extracted_context = "Kein Kontext extrahiert."

def get_datenspalten_options(df):
    if df is None: return []
    geschuetzt = ["bezeichnung", "wo?", "id", "kategorie", "relevanz gast"]
    return [str(col).strip() for col in df.columns if "status" not in str(col).lower() and str(col).strip().lower() not in geschuetzt]

# ==============================================================================
# 3. DATEN-LADE ENGINE
# ==============================================================================
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
        while not done: _, done = downloader.next_chunk()
            
        fh.seek(0)
        df_wissen = pd.read_excel(fh, sheet_name="Wissensbasis", header=0)
        fh.seek(0)
        df_lexikon = pd.read_excel(fh, sheet_name="Spalten_Lexikon", header=0)
        fh.seek(0)
        df_usecases = pd.read_excel(fh, sheet_name="UseCase_Lexikon", header=0)
        fh.seek(0)
        try:
            df_passwoerter = pd.read_excel(fh, sheet_name="Passwort_Lexikon", header=0)
        except Exception:
            df_passwoerter = None
        
        if df_wissen is not None and not df_wissen.empty and "Wo?" in df_wissen.columns:
            df_wissen["Wo?"] = df_wissen["Wo?"].ffill()
            
        return df_wissen, df_lexikon, df_usecases, df_passwoerter, service
    except Exception as e:
        st.error(f"Kritischer Fehler beim Laden der Datenspezifikation: {e}")
        return None, None, None, None, None

with st.spinner("Synchronisiere mit der Excel-Zentralmatrix..."):
    df_wissen, df_lexikon, df_usecases, df_passwoerter, drive_service = load_dynamic_data()

if df_usecases is not None: df_usecases.columns = [str(c).strip() for c in df_usecases.columns]
if df_passwoerter is not None: df_passwoerter.columns = [str(c).strip() for c in df_passwoerter.columns]

# ==============================================================================
# 4. API-CORE & STRUCTURATED ROUTING ENGINE
# ==============================================================================
@st.cache_resource
def get_ki_client():
    if "GEMINI_API_KEY" in st.secrets: return genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
    return None

def find_column_by_fuzzy_name(headers, target_name):
    cleaned = [str(h).strip().lower().replace("\n", " ") for h in headers]
    search = str(target_name).strip().lower()
    if search in cleaned: return cleaned.index(search) + 1
    for idx, h in enumerate(cleaned):
        if search in h: return idx + 1
    return None

def call_gemini_api_structured(prompt, context="", system_context=None):
    client = get_ki_client()
    if client is None: return KiAntwortSchema(wissensluecke_erkannt=True, antwort_text="🛑 KI nicht konfiguriert.")
    sys_instruction = system_context or (
        "Du bist „Villa Avatar“, der digitale Helfer. Antworte immer kurz, freundlich, präzise und smartphone-optimiert. "
        "Analysiere den bereitgestellten Excel-Kontext intelligent. Befindet sich die Information zu einer Handlungsfrage implizit im Text, "
        "übersetze dies in eine direkte Anweisung für den Gast und setze wissensluecke_erkannt = False.\n"
        "Wenn das Thema im Kontext überhaupt nicht behandelt wird oder unvollständig ist, setze wissensluecke_erkannt = True.\n"
        "ABSOLUTES VERBOT: Erwähne NIEMALS interne Dateinamen, Spaltenüberschriften oder die Struktur der Excel-Tabelle."
    )
    full_prompt = f"Kontext aus der verifizierten Wissensbasis:\n{context}\n\nNutzerfrage: {prompt}" if context else prompt
    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash", contents=full_prompt,
            config=types.GenerateContentConfig(system_instruction=sys_instruction, temperature=0.2, response_mime_type="application/json", response_schema=KiAntwortSchema)
        )
        data = json.loads(response.text)
        return KiAntwortSchema(wissensluecke_erkannt=bool(data.get("wissensluecke_erkannt", True)), antwort_text=str(data.get("antwort_text", "")))
    except Exception:
        return KiAntwortSchema(wissensluecke_erkannt=True, antwort_text="")

def call_gemini_api_raw(prompt, system_context=None):
    client = get_ki_client()
    if client is None: return "🛑 KI nicht konfiguriert."
    try:
        response = client.models.generate_content(model="gemini-2.5-flash", contents=prompt, config=types.GenerateContentConfig(system_instruction=system_context, temperature=0.2))
        return response.text
    except Exception as e:
        return f"🛑 KI-Fehler: {e}"

def extract_context_for_object(objekt_name):
    if df_wissen is None or df_lexikon is None or objekt_name is None: return ""
    df_wissen.columns = [str(c).strip() for c in df_wissen.columns]
    df_lexikon.columns = [str(c).strip() for c in df_lexikon.columns]
    
    bez_col = df_wissen.columns[0] if "Bezeichnung" not in df_wissen.columns else "Bezeichnung"
    row_match = df_wissen[df_wissen[bez_col].astype(str).str.strip().str.lower() == objekt_name.lower().strip()]
    if row_match.empty: return ""
    
    aktuelle_rolle = str(st.session_state.get("aktive_rolle", "Gast")).strip().lower()
    context_parts = [f"Informationen zum Objekt: {objekt_name}"]

    if aktuelle_rolle == "host":
        for col in df_wissen.columns:
            if col != bez_col and "status" not in col.lower() and col.lower() != "wo?":
                val = row_match.iloc[0][col]
                if pd.notna(val) and str(val).strip() != "": context_parts.append(f"- {col}: {str(val).strip()}")
    else:
        lex_spalten_name = df_lexikon.columns[0]
        gast_freigabe_spalte = next((col for col in df_lexikon.columns if "gast" in col.lower()), df_lexikon.columns[-1])
        mask_ja = df_lexikon[gast_freigabe_spalte].astype(str).str.lower().str.strip() == "ja"
        freigegebene_tags = df_lexikon[mask_ja][lex_spalten_name].astype(str).str.strip().tolist()
        
        for col in df_wissen.columns:
            if any(col.lower() == tag.lower() for tag in freigegebene_tags) and col in row_match.columns:
                val = row_match.iloc[0][col]
                if pd.notna(val) and str(val).strip() != "": context_parts.append(f"- {col}: {str(val).strip()}")
                    
    final_context = "\n".join(context_parts)
    st.session_state.last_extracted_context = final_context
    return final_context

# ==============================================================================
# 5. MATRIZEN-SCHREIBENGINE
# ==============================================================================
def execute_matrix_input_direct(physische_zielspalte, objekt_name, freitext):
    if drive_service is None or df_wissen is None: return
    try:
        ziel_objekt = "Nicht gefunden" if not objekt_name else objekt_name
        request = drive_service.files().get_media(fileId=FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done: _, done = downloader.next_chunk()
        
        fh.seek(0)
        wb = openpyxl.load_workbook(fh)
        ws = wb["Wissensbasis"]
        headers = [str(c.value) if c.value else "" for c in ws[1]]
        
        col_bez_idx = find_column_by_fuzzy_name(headers, "Bezeichnung") or 1
        ziel_row_idx = next((row for row in range(2, ws.max_row + 1) if ws.cell(row=row, column=col_bez_idx).value and str(ws.cell(row=row, column=col_bez_idx).value).strip().lower() == ziel_objekt.lower().strip()), None)
        
        if not ziel_row_idx:
            ziel_row_idx = ws.max_row + 1
            ws.cell(row=ziel_row_idx, column=col_bez_idx).value = ziel_objekt
            
        ziel_col_idx = find_column_by_fuzzy_name(headers, physische_zielspalte)
        if not ziel_col_idx: return 
            
        zeitstempel = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
        nutzer = st.session_state.aktive_rolle or "System"
        alter_inhalt = ws.cell(row=ziel_row_idx, column=ziel_col_idx).value or ""
        
        neuer_eintrag = f"[{zeitstempel} | {nutzer}]: {freitext}"
        ws.cell(row=ziel_row_idx, column=ziel_col_idx).value = f"{alter_inhalt}\n{neuer_eintrag}" if alter_inhalt else neuer_eintrag
        ws.cell(row=ziel_row_idx, column=ziel_col_idx).font = Font(color="1F4E78")
        ws.cell(row=ziel_row_idx, column=ziel_col_idx).alignment = Alignment(wrap_text=True)
        
        status_col_idx = find_column_by_fuzzy_name(headers, f"{physische_zielspalte} Status")
        if status_col_idx:
            alter_status = ws.cell(row=ziel_row_idx, column=status_col_idx).value or ""
            ws.cell(row=ziel_row_idx, column=status_col_idx).value = f"{alter_status}\n[{zeitstempel}]: offen" if alter_status else f"[{zeitstempel}]: offen"
            ws.cell(row=ziel_row_idx, column=status_col_idx).font = Font(color="1F4E78")
                
        output_stream = io.BytesIO()
        wb.save(output_stream)
        output_stream.seek(0)
        media = MediaIoBaseUpload(output_stream, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        drive_service.files().update(fileId=FILE_ID, media_body=media).execute()
        
        st.session_state.last_write_status = f"✅ ERFOLG: Zeile {ziel_row_idx}, Spalte '{physische_zielspalte}' beschrieben."
        st.toast("✅ Matrix aktualisiert!")
    except Exception as e:
        st.error(f"Schreibfehler: {e}")

def execute_matrix_input(use_case_name, objekt_name, freitext):
    if df_lexikon is None: return
    tag_col_name, lex_spalten_name = df_lexikon.columns[4], df_lexikon.columns[0]
    physische_zielspalte = None
    for _, row in df_lexikon.iterrows():
        if use_case_name.lower().strip() in [t.strip().lower() for t in str(row[tag_col_name]).split(',')]:
            physische_zielspalte = str(row[lex_spalten_name]).strip()
            break
    if not physische_zielspalte and use_case_name == "Keine Information":
        physische_zielspalte = next((col for col in df_wissen.columns if "information" in col.lower() and "status" not in col.lower()), None)
    if physische_zielspalte:
        execute_matrix_input_direct(physische_zielspalte, objekt_name, freitext)

def execute_transitional_routing(user_input, objekt_name=None):
    st.session_state.messages.append({"role": "assistant", "content": FALLBACK_SATZ})
    execute_matrix_input("Keine Information", objekt_name or "Nicht gefunden", user_input)
    st.cache_data.clear()

def generate_raw_report_context(filter_type):
    if df_wissen is None: return "Keine Einträge verfügbar."
    report_lines = []
    for col in df_wissen.columns:
        verarbeiten = False
        if filter_type == "offene_stoerungen" and "störung" in col.lower() and "status" not in col.lower(): verarbeiten = True
        elif filter_type == "behobene_stoerungen" and "störung" in col.lower() and "status" not in col.lower(): verarbeiten = True
        elif filter_type == "offenes_feedback" and "feedback" in col.lower() and "status" not in col.lower(): verarbeiten = True
        elif filter_type == "ignoriertes_feedback" and "feedback" in col.lower() and "status" not in col.lower(): verarbeiten = True
        elif filter_type == "offene_luecken" and "information" in col.lower() and "status" not in col.lower(): verarbeiten = True
        elif filter_type == "gesamtuebersicht" and ("störung" in col.lower() or "feedback" in col.lower() or "information" in col.lower()) and "status" not in col.lower(): verarbeiten = True
            
        if verarbeiten:
            bez_col = df_wissen.columns[0]
            for idx, row in df_wissen.iterrows():
                cell_val = row[col]
                ist_gueltig = True
                status_col = f"{col} Status"
                if status_col in df_wissen.columns:
                    status_val = str(row[status_col]).lower()
                    if filter_type == "offene_stoerungen" and "aktiv" not in status_val: ist_gueltig = False
                    if filter_type == "behobene_stoerungen" and "ok" not in status_val and "beheben" not in status_val: ist_gueltig = False
                    if filter_type == "offenes_feedback" and "offen" not in status_val: ist_gueltig = False
                    if filter_type == "ignoriertes_feedback" and "nein" not in status_val and "ignorier" not in status_val: ist_gueltig = False
                    if filter_type == "offene_luecken" and "offen" not in status_val: ist_gueltig = False
                    
                if pd.notna(cell_val) and str(cell_val).strip() != "" and ist_gueltig:
                    report_lines.append(f"Objekt: {row[bez_col]} | Kat: {col}\nEintrag: {cell_val}\n---")
    return "\n".join(report_lines) if report_lines else "Keine passenden Einträge gefunden."

# ==============================================================================
# 6. HMI PRESENTATION LAYER
# ==============================================================================
st.title("☀️ Villa Avatar")

selected_role = st.selectbox("Rolle", options=["Gast", "Host"], index=None, placeholder="Wer bist du?", label_visibility="collapsed")
if selected_role and selected_role != st.session_state.aktive_rolle:
    st.session_state.aktive_rolle = selected_role
    st.session_state.aktiver_use_case = None
    st.session_state.selected_object = None
    st.session_state.messages = []
    st.rerun()

if not st.session_state.aktive_rolle: st.stop()

# Host-Passwort-Gate
if st.session_state.aktive_rolle == "Host" and not st.session_state.host_authentifiziert:
    st.write("---")
    pwd_input = st.text_input("🔑 Bitte Passwort für Host-Sicht eingeben:", type="password")
    if pwd_input and df_passwoerter is not None and not df_passwoerter.empty:
        p_rolle_col, p_pwd_col = df_passwoerter.columns[0], df_passwoerter.columns[1]
        host_rows = df_passwoerter[df_passwoerter[p_rolle_col].astype(str).str.strip().str.lower() == str(df_passwoerter.iloc[0][p_rolle_col]).strip().lower()]
        for _, row in host_rows.iterrows():
            if pwd_input.strip() == str(row[p_pwd_col]).strip():
                st.session_state.host_authentifiziert = True
                if len(df_passwoerter.columns) > 2 and str(row[df_passwoerter.columns[2]]).strip().lower() == "debug":
                    st.session_state.debug_modus_aktiv = True
                st.success("Erfolgreich eingeloggt!")
                st.rerun()
        st.error("❌ Falsches Passwort.")
    st.stop()

# Use Case Selektion
st.write("---")
if df_usecases is not None:
    uc_col, dir_col, hmi_col = df_usecases.columns[0], df_usecases.columns[1], df_usecases.columns[2]
    btn_col = df_usecases.columns[3] if len(df_usecases.columns) > 3 else None
    
    all_uc = df_usecases[df_usecases[hmi_col].astype(str).str.lower().str.strip() == "ja"][uc_col].tolist()
    erlaubte_buttons = [uc for uc in all_uc if any(x in uc.lower() for x in ["hilfe", "störung", "feedback"])] if st.session_state.aktive_rolle == "Gast" else all_uc
    
    cols = st.columns(len(erlaubte_buttons))
    for idx, uc_name in enumerate(erlaubte_buttons):
        with cols[idx]:
            button_label = uc_name
            if btn_col:
                btn_match = df_usecases[df_usecases[uc_col].astype(str).str.strip() == str(uc_name).strip()]
                if not btn_match.empty and pd.notna(btn_match.iloc[0][btn_col]):
                    button_label = str(btn_match.iloc[0][btn_col]).strip()
            
            if st.button(button_label, use_container_width=True, type="primary" if st.session_state.aktiver_use_case == uc_name else "secondary"):
                st.session_state.aktiver_use_case = uc_name
                st.session_state.selected_object = None
                st.session_state.messages = []
                st.rerun()

if not st.session_state.aktiver_use_case: st.stop()

# Metadaten aus dem UseCase_Lexikon
uc_row = df_usecases[df_usecases[df_usecases.columns[0]].astype(str).str.lower().str.strip() == st.session_state.aktiver_use_case.lower().strip()]
aktuelle_richtung = str(uc_row.iloc[0][df_usecases.columns[1]]).strip().upper() if not uc_row.empty else "OUTPUT"
chat_abfrage_text = str(uc_row.iloc[0][df_usecases.columns[4]]).strip() if not uc_row.empty and len(df_usecases.columns) > 4 and pd.notna(uc_row.iloc[0][df_usecases.columns[4]]) else "Wie kann ich dir helfen?"
danke_template = str(uc_row.iloc[0][df_usecases.columns[5]]).strip() if not uc_row.empty and len(df_usecases.columns) > 5 and pd.notna(uc_row.iloc[0][df_usecases.columns[5]]) else "Vielen Dank!"

# ==============================================================================
# WELT 1: DETERMINISTISCHER FORMULAR-MODUS (Bericht / Neue Information)
# ==============================================================================
if "bericht" in st.session_state.aktiver_use_case.lower():
    st.subheader("📋 System-Berichte abrufen")
    b_typ = st.selectbox("1. Welche Art von Bericht benötigst du?", options=["Offene Störungen", "Behobene Störungen", "Offenes Feedback", "Offene Wissenslücken", "Gesamtübersicht"], index=None, placeholder="Bitte wählen...")
    zeitraum = st.selectbox("2. Zeitraum auswählen:", options=["Letzte 24 Stunden", "Letzte 7 Tage", "Alle Einträge"], index=None, placeholder="Bitte wählen...")
    if b_typ and zeitraum:
        if st.button("📊 Bericht jetzt generieren", type="primary", use_container_width=True):
            with st.spinner("Generiere Übersicht..."):
                filter_map = {"Offene Störungen": "offene_stoerungen", "Behobene Störungen": "behobene_stoerungen", "Offenes Feedback": "offenes_feedback", "Offene Wissenslücken": "offene_luecken", "Gesamtübersicht": "gesamtuebersicht"}
                raw_context = generate_raw_report_context(filter_map.get(b_typ, "gesamtuebersicht"))
                if "Keine passenden Einträge" in raw_context:
                    st.info(f"Aktuell liegen keine Einträge für '{b_typ}' vor. ☀️")
                else:
                    st.markdown(call_gemini_api_raw(f"Strukturiere diese Matrix-Daten professionell und chronologisch für den Host:\n\n{raw_context}", system_context="Liste Fakten auf, nutze Bulletpoints, bleibe sachlich."))
    st.stop()

elif st.session_state.aktiver_use_case == "Neue Information" and st.session_state.aktive_rolle == "Host":
    st.subheader("📍 Neue Information in Matrix einpflegen")
    alle_objekte = sorted(df_wissen[df_wissen.columns[0]].dropna().astype(str).str.strip().unique().tolist())
    if "Nicht gefunden" in alle_objekte: alle_objekte.remove("Nicht gefunden")
    alle_objekte.append("Nicht gefunden")
    
    obj_wahl = st.selectbox("1. Bitte wähle das Objekt aus:", options=alle_objekte, index=None, placeholder="Objekt suchen...")
    feld_wahl = st.selectbox("2. Bitte wähle die Art der Information (Ziel-Datenfeld):", options=get_datenspalten_options(df_wissen), index=None, placeholder="Spalte wählen...")
    
    if obj_wahl and feld_wahl:
        st.write("---")
        with st.form("neue_info_form", clear_on_submit=True):
            info_text = st.text_area(f"Gib hier den neuen Informationstext für '{obj_wahl}' ➔ '{feld_wahl}' ein:", height=120)
            if st.form_submit_button("💾 Eintrag in Excel-Zentralmatrix speichern", type="primary") and info_text.strip():
                with st.spinner("Schreibe in Zentralmatrix..."):
                    execute_matrix_input_direct(feld_wahl, obj_wahl, info_text.strip())
                    st.success(f"Erfolgreich eingetragen! Das Thema '{feld_wahl}' wurde für das Objekt '{obj_wahl}' aktualisiert.")
    st.stop()

# ==============================================================================
# WELT 2: FLEXIBLER KI-CHAT-MODUS
# ==============================================================================
else:
    STANDARD_DROPDOWNS = ["Ausstattung innen", "Ausstattung außen", "In der Nähe"]
    bez_spalte, kat_spalte = df_wissen.columns[0], (df_wissen.columns[1] if "Wo?" not in df_wissen.columns else "Wo?")
    
    st.write("")
    
    # Rendern der voneinander isolierten Kaskaden-Dropdowns mit eindeutigen Labels
    for kat in STANDARD_DROPDOWNS:
        if "innen" in kat.lower(): mask = df_wissen[kat_spalte].astype(str).str.contains("innen", case=False, na=False)
        elif "außen" in kat.lower() or "aussen" in kat.lower(): mask = df_wissen[kat_spalte].astype(str).str.contains("außen|aussen", case=False, na=False)
        else: mask = df_wissen[kat_spalte].astype(str).str.contains("nähe|naehe|In der Nähe", case=False, na=False)
        
        if st.session_state.aktive_rolle == "Gast" and "Relevanz Gast" in df_wissen.columns:
            mask = mask & (df_wissen["Relevanz Gast"].astype(str).str.strip().str.lower() == "x")
        
        verfuegbare_bez = df_wissen[mask][bez_spalte].dropna().drop_duplicates().tolist()
        verfuegbare_bez = sorted([str(b).strip() for b in verfuegbare_bez])
        if "Nicht gefunden" in verfuegbare_bez: verfuegbare_bez.remove("Nicht gefunden")
        verfuegbare_bez.append("Nicht gefunden")
        
        dp_key = f"dropdown_{kat}_{st.session_state.aktiver_use_case}"
        
        aktueller_idx = None
        if st.session_state.selected_object in verfuegbare_bez:
            aktueller_idx = verfuegbare_bez.index(st.session_state.selected_object)
            
        # FIX: Eindeutige Label-Strings pro Dropdown verhindern das Verschwinden des Widgets
        auswahl = st.selectbox(label=f"Auswahl {kat}", options=verfuegbare_bez, index=aktueller_idx, placeholder=f"🔎 {kat} wählen...", key=dp_key, label_visibility="collapsed")
        if auswahl and auswahl != st.session_state.selected_object:
            st.session_state.selected_object = auswahl
            st.session_state.messages = []
            st.rerun()

    if not st.session_state.selected_object:
        st.stop()
        
    # DIAGNOSE MONITOR
    if st.session_state.debug_modus_aktiv:
        st.write("")
        with st.expander("🔍 SYSTEM-DIAGNOSE MONITOR (Laufzeit-Metriken)", expanded=True):
            d_col1, d_col2 = st.columns(2)
            with d_col1:
                st.metric(label="1. Aktive Rolle", value=str(st.session_state.get("aktive_rolle", "None")))
                st.metric(label="3. Gewähltes Objekt", value=str(st.session_state.selected_object))
            with d_col2:
                st.metric(label="2. Use Case | Richtung", value=f"{st.session_state.get('aktiver_use_case')} | {aktuelle_richtung}")
            
            st.write("**4. Letzter Matrix-Schreibstatus:**")
            st.info(st.session_state.get("last_write_status", "Kein Status"))
            
            st.write("**5. Letzter Kontext-Extrakt (KI-Input):**")
            st.text_area(label="Matrix-Rohdaten", value=st.session_state.get("last_extracted_context", ""), height=120, disabled=True, label_visibility="collapsed")

    st.write("---")
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]): st.markdown(msg["content"])
        
    if user_input := st.chat_input(chat_abfrage_text):
        st.session_state.messages.append({"role": "user", "content": user_input})
        st.rerun()
        
    if st.session_state.messages and st.session_state.messages[-1]["role"] == "user":
        aktueller_nutzer_text = st.session_state.messages[-1]["content"]
        
        if aktuelle_richtung == "OUTPUT":
            if st.session_state.selected_object == "Nicht gefunden":
                execute_transitional_routing(aktueller_nutzer_text, "Nicht gefunden")
                st.rerun()
            else:
                with st.spinner("Durchsuche Matrix..."):
                    context_str = extract_context_for_object(st.session_state.selected_object)
                    res = call_gemini_api_structured(aktueller_nutzer_text, context_str)
                    
                    luecken_phrasen = ["keine information", "weiß ich nicht", "nicht hinterlegt", "leider nein", "nicht bekannt", "fehlen mir details"]
                    if res.wissensluecke_erkannt or res.antwort_text == "" or any(p in res.antwort_text.lower() for p in luecken_phrasen):
                        execute_transitional_routing(aktueller_nutzer_text, st.session_state.selected_object)
                    else:
                        st.session_state.messages.append({"role": "assistant", "content": res.antwort_text})
                    st.rerun()
                    
        elif aktuelle_richtung == "INPUT":
            with st.spinner("Protokolliere Eintrag in Zentralmatrix..."):
                execute_matrix_input(st.session_state.aktiver_use_case, st.session_state.selected_object, aktueller_nutzer_text)
                danke_satz = danke_template.replace("{use_case}", st.session_state.aktiver_use_case)
                st.session_state.messages.append({"role": "assistant", "content": danke_satz})
                st.cache_data.clear()
                st.rerun()
