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
# 1. STRUCTURATED OUTPUT SCHEMA (Laut Spezifikation Kapitel 4.4)
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
        flex-direction: row-reverse !important; 
        background-color: #F0F2F6 !important; 
        border-radius: 10px !important; 
        padding: 10px !important; 
    }
    div[data-testid="stChatMessage"]:has(div[aria-label="Chat message from user"]) div[data-testid="stChatMessageContent"] { text-align: right !important; width: 100% !important; }
    </style>
""", unsafe_allow_html=True)

if "last_write_status" not in st.session_state: st.session_state.last_write_status = "Noch kein Schreibvorgang ausgelöst."
if "last_extracted_context" not in st.session_state: st.session_state.last_extracted_context = "Kein Kontext extrahiert."

def get_datenspalten_options(df):
    if df is None: return []
    geschuetzt = ["bezeichnung", "wo?", "id", "kategorie", "relevanz gast"]
    ergebnis = []
    for col in df.columns:
        col_clean = str(col).strip().lower()
        if "status" not in col_clean and col_clean not in geschuetzt:
            ergebnis.append(str(col).strip())
    return ergebnis

# ==============================================================================
# 3. DATEN-LADE ENGINE (Vier-Blatt-Modell mit header=0 laut Spezifikation)
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
        while done is False: _, done = downloader.next_chunk()
            
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

if df_usecases is not None and not df_usecases.empty:
    df_usecases.columns = [str(c).strip() for c in df_usecases.columns]
if df_passwoerter is not None and not df_passwoerter.empty:
    df_passwoerter.columns = [str(c).strip() for c in df_passwoerter.columns]

# ==============================================================================
# 4. API-CORE & STRUCTURATED ROUTING ENGINE
# ==============================================================================
@st.cache_resource
def get_ki_client():
    if "GEMINI_API_KEY" in st.secrets: 
        return genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
    return None

def find_column_by_fuzzy_name(headers, target_name):
    cleaned_headers = [str(h).strip().lower().replace("\n", " ") for h in headers]
    search = str(target_name).strip().lower()
    if search in cleaned_headers:
        return cleaned_headers.index(search) + 1
    for idx, h in enumerate(cleaned_headers):
        if search in h:
            return idx + 1
    return None

def call_gemini_api_structured(prompt, context="", system_context=None):
    client = get_ki_client()
    if client is None:
        return KiAntwortSchema(wissensluecke_erkannt=True, antwort_text="🛑 KI-Schnittstelle nicht konfiguriert.")
    
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
            model="gemini-2.5-flash", 
            contents=full_prompt, 
            config=types.GenerateContentConfig(
                system_instruction=sys_instruction,
                temperature=0.2,
                response_mime_type="application/json",
                response_schema=KiAntwortSchema
            )
        )
        data = json.loads(response.text)
        return KiAntwortSchema(
            wissensluecke_erkannt=bool(data.get("wissensluecke_erkannt", True)),
            antwort_text=str(data.get("antwort_text", ""))
        )
    except Exception as e:
        return KiAntwortSchema(wissensluecke_erkannt=True, antwort_text="")

def extract_context_for_object(objekt_name):
    if df_wissen is None or df_lexikon is None or objekt_name is None: return ""
    df_wissen.columns = [str(c).strip() for c in df_wissen.columns]
    df_lexikon.columns = [str(c).strip() for c in df_lexikon.columns]
    
    bez_col = df_wissen.columns[0] if "Bezeichnung" not in df_wissen.columns else "Bezeichnung"
    row_match = df_wissen[df_wissen[bez_col].astype(str).str.strip().str.lower() == objekt_name.lower().strip()]
    if row_match.empty: return ""
    
    aktuelle_rolle = str(st.session_state.get("aktive_rolle", "Gast")).strip().lower()
    context_parts = [f"Informationen zum Objekt: {objekt_name}"]

    if aktuelle_rolle in ["host", "host_verifiziert"]:
        for col in df_wissen.columns:
            if col != bez_col and "status" not in col.lower() and col.lower() != "wo?":
                val = row_match.iloc[0][col]
                if pd.notna(val) and str(val).strip() != "":
                    context_parts.append(f"- {col}: {str(val).strip()}")
    else:
        lex_spalten_name = df_lexikon.columns[0]
        gast_freigabe_spalte = None
        for col in df_lexikon.columns:
            if "gast" in col.lower():
                gast_freigabe_spalte = col
                break
        if not gast_freigabe_spalte:
            gast_freigabe_spalte = df_lexikon.columns[-1]
        
        mask_ja = df_lexikon[gast_freigabe_spalte].astype(str).str.lower().str.strip() == "ja"
        freigegebene_tags = df_lexikon[mask_ja][lex_spalten_name].astype(str).str.strip().tolist()
        
        for col in df_wissen.columns:
            if any(col.lower() == tag.lower() for tag in freigegebene_tags) and col in row_match.columns:
                val = row_match.iloc[0][col]
                if pd.notna(val) and str(val).strip() != "":
                    context_parts.append(f"- {col}: {str(val).strip()}")
                
    return "\n".join(context_parts)

# ==============================================================================
# 5. MATRIZEN-SCHREIBENGINE & REPORTING
# ==============================================================================
def execute_matrix_input_direct(physische_zielspalte, objekt_name, freitext):
    if drive_service is None or df_wissen is None: return
    try:
        ziel_objekt = "Nicht gefunden" if (objekt_name is None or objekt_name == "Nicht gefunden") else objekt_name
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
        
        ziel_row_idx = None
        for row in range(2, ws.max_row + 1):
            cell_val = ws.cell(row=row, column=col_bez_idx).value
            if cell_val and str(cell_val).strip().lower() == ziel_objekt.lower().strip():
                ziel_row_idx = row
                break
        if not ziel_row_idx:
            ziel_row_idx = ws.max_row + 1
            ws.cell(row=ziel_row_idx, column=col_bez_idx).value = ziel_objekt
            
        ziel_col_idx = find_column_by_fuzzy_name(headers, physische_zielspalte)
        if not ziel_col_idx: return 
            
        zeitstempel = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
        nutzer = st.session_state.aktive_rolle if st.session_state.aktive_rolle else "System"
        alter_inhalt = ws.cell(row=ziel_row_idx, column=ziel_col_idx).value or ""
        
        neuer_eintrag = f"[{zeitstempel} | {nutzer}]: {freitext}"
        ws.cell(row=ziel_row_idx, column=ziel_col_idx).value = f"{alter_inhalt}\n{neuer_eintrag}" if alter_inhalt else neuer_eintrag
        
        output_stream = io.BytesIO()
        wb.save(output_stream)
        output_stream.seek(0)
        media = MediaIoBaseUpload(output_stream, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        drive_service.files().update(fileId=FILE_ID, media_body=media).execute()
        st.toast("✅ Excel-Zentralmatrix aktualisiert!")
    except Exception as e:
        st.error(f"Schreibfehler: {e}")

def execute_matrix_input(use_case_name, objekt_name, freitext):
    if df_lexikon is None: return
    tag_col_name = df_lexikon.columns[4]
    lex_spalten_name = df_lexikon.columns[0]
    physische_zielspalte = None
    for _, row in df_lexikon.iterrows():
        if use_case_name.lower().strip() in [t.strip().lower() for t in str(row[tag_col_name]).split(',')]:
            physische_zielspalte = str(row[lex_spalten_name]).strip()
            break
    if not physische_zielspalte and use_case_name == "Keine Information":
        for col in df_wissen.columns:
            if "information" in col.lower() and "status" not in col.lower(): physische_zielspalte = col; break
    if physische_zielspalte: execute_matrix_input_direct(physische_zielspalte, objekt_name, freitext)

def execute_transitional_routing(user_input, objekt_name=None):
    st.session_state.messages.append({"role": "assistant", "content": FALLBACK_SATZ})
    execute_matrix_input("Keine Information", objekt_name if objekt_name else "Nicht gefunden", user_input)
    st.cache_data.clear()

def generate_raw_report_context(filter_type):
    if df_wissen is None: return "Keine Einträge verfügbar."
    report_lines, ft = [], str(filter_type).lower()
    for col in df_wissen.columns:
        verarbeiten = False
        if "störung" in ft and "störung" in col.lower() and "status" not in col.lower(): verarbeiten = True
        elif "feedback" in ft and "feedback" in col.lower() and "status" not in col.lower(): verarbeiten = True
        elif "lücken" in ft and "information" in col.lower() and "status" not in col.lower(): verarbeiten = True
        elif "gesamt" in ft and ("störung" in col.lower() or "feedback" in col.lower() or "information" in col.lower()) and "status" not in col.lower(): verarbeiten = True
        if verarbeiten:
            bez_col = df_wissen.columns[0]
            for idx, row in df_wissen.iterrows():
                cell_val = row[col]
                ist_gueltig = True
                status_col = f"{col} Status"
                if status_col in df_wissen.columns:
                    status_val = str(row[status_col]).lower()
                    if "offene störungen" in ft and "aktiv" not in status_val: ist_gueltig = False
                    if "behobene" in ft and "ok" not in status_val and "beheben" not in status_val: ist_gueltig = False
                    if "offenes feedback" in ft and "offen" not in status_val: ist_gueltig = False
                    if "ignoriertes" in ft and "nein" not in status_val and "ignorier" not in status_val: ist_gueltig = False
                    if "wissenslücken" in ft and "offen" not in status_val: ist_gueltig = False
                if pd.notna(cell_val) and str(cell_val).strip() != "" and ist_gueltig:
                    report_lines.append(f"Objekt: {row[bez_col]} | Kat: {col}\nEintrag: {cell_val}\n---")
    return "\n".join(report_lines) if report_lines else "Keine passenden Einträge gefunden."

# ==============================================================================
# NATIVE PASSPORT VALIDATION CALLBACK
# ==============================================================================
def check_password_callback():
    pwd_input = st.session_state.get("host_pwd_field", "").strip()
    if pwd_input == "" or df_passwoerter is None or df_passwoerter.empty: return
    p_rolle_col, p_pwd_col = df_passwoerter.columns[0], df_passwoerter.columns[1]
    host_rows = df_passwoerter[df_passwoerter[p_rolle_col].astype(str).str.strip().str.lower() == str(df_passwoerter.iloc[0][p_rolle_col]).strip().lower()]
    for _, row in host_rows.iterrows():
        if pwd_input == str(row[p_pwd_col]).strip():
            st.session_state.host_authentifiziert = True
            st.session_state["pwd_error_msg"] = None
            return
    st.session_state["pwd_error_msg"] = "❌ Falsches Passwort."

# ==============================================================================
# 6. HMI PRESENTATION LAYER
# ==============================================================================
st.title("☀️ Villa Avatar")

if "aktive_rolle" not in st.session_state: st.session_state.aktive_rolle = None
if "aktiver_use_case" not in st.session_state: st.session_state.aktiver_use_case = None
if "messages" not in st.session_state: st.session_state.messages = []
if "host_authentifiziert" not in st.session_state: st.session_state.host_authentifiziert = False
if "aktuell_gewaehltes_objekt" not in st.session_state: st.session_state.aktuell_gewaehltes_objekt = None

neue_rolle = st.selectbox("Rolle", options=["Gast", "Host"], index=None, placeholder="Wer bist du?", label_visibility="collapsed")

if neue_rolle is not None and neue_rolle != st.session_state.aktive_rolle:
    st.session_state.aktive_rolle = neue_rolle
    st.session_state.aktiver_use_case = None
    st.session_state.aktuell_gewaehltes_objekt = None
    st.session_state.messages = []
    for key in list(st.session_state.keys()):
        if any(x in key for x in ["dropdown_", "target_col_", "direct_", "host_"]): del st.session_state[key]
    st.rerun()

if st.session_state.aktive_rolle == "Host" and not st.session_state.host_authentifiziert:
    st.write("---")
    if st.session_state.get("pwd_error_msg"): st.error(st.session_state["pwd_error_msg"])
    st.text_input("🔑 Passwort eingeben:", type="password", key="host_pwd_field", on_change=check_password_callback)
    st.stop()

aktuelle_richtung = None
chat_abfrage_text = "Wie kann ich dir helfen?"
danke_text_template = "Vielen Dank! Ich habe deine Eingabe zum Thema '{use_case}' eingetragen."

if st.session_state.aktive_rolle in ["Host", "Gast"] and df_usecases is not None:
    st.write("---")
    uc_col, dir_col, hmi_col = df_usecases.columns[0], df_usecases.columns[1], df_usecases.columns[2]
    btn_col, prompt_col, danke_col = df_usecases.columns[3] if len(df_usecases.columns) > 3 else None, df_usecases.columns[4] if len(df_usecases.columns) > 4 else None, df_usecases.columns[5] if len(df_usecases.columns) > 5 else None
    
    mask_sichtbar = df_usecases[hmi_col].astype(str).str.lower().str.strip() == "ja"
    verfuegbare_uc = df_usecases[mask_sichtbar][uc_col].tolist()
    erlaubte_buttons = [uc for uc in verfuegbare_uc if any(x in uc.lower() for x in ["hilfe", "störung", "feedback"])] if st.session_state.aktive_rolle == "Gast" else verfuegbare_uc

    cols = st.columns(len(erlaubte_buttons))
    for idx, uc_name in enumerate(erlaubte_buttons):
        with cols[idx]:
            is_active = (st.session_state.aktiver_use_case == uc_name)
            lbl = uc_name
            if btn_col:
                btn_match = df_usecases[df_usecases[uc_col].astype(str).str.strip() == str(uc_name).strip()]
                if not btn_match.empty and pd.notna(btn_match.iloc[0][btn_col]): lbl = str(btn_match.iloc[0][btn_col]).strip()
            
            if st.button(lbl, use_container_width=True, type="primary" if is_active else "secondary", key=f"btn_{uc_name}"):
                st.session_state.aktiver_use_case = uc_name
                st.session_state.aktuell_gewaehltes_objekt = None
                st.session_state.messages = []
                for key in list(st.session_state.keys()):
                    if any(x in key for x in ["dropdown_", "target_col_", "direct_", "host_"]): del st.session_state[key]
                st.rerun()

    # ==============================================================================
    # CRITICAL INLINE GATING SYSTEM (Verhindert fehlerhaftes Absinken der Logik)
    # ==============================================================================
    if st.session_state.aktiver_use_case:
        uc_row = df_usecases[df_usecases[uc_col].astype(str).str.lower().str.strip() == st.session_state.aktiver_use_case.lower().strip()]
        if not uc_row.empty:
            aktuelle_richtung = str(uc_row.iloc[0][dir_col]).strip().upper()
            if prompt_col and pd.notna(uc_row.iloc[0][prompt_col]) and str(uc_row.iloc[0][prompt_col]).strip() != "":
                chat_abfrage_text = str(uc_row.iloc[0][prompt_col]).strip()
            if danke_col and pd.notna(uc_row.iloc[0][danke_col]):
                danke_text_template = str(uc_row.iloc[0][danke_col]).strip()
            
            if "bericht" in st.session_state.aktiver_use_case.lower():
                st.write("")
                bericht_optionen = ["⚠️ Offene Störungen", "✅ Behobene Störungen", "💡 Offenes Feedback", "❌ Ignoriertes Feedback", "🔍 Offene Wissenslücken", "📋 Gesamtübersicht"]
                rep_sel = st.selectbox("📋 Berichtstyp wählen:", options=bericht_optionen, index=None, label_visibility="collapsed")
                if rep_sel:
                    st.text_area("Berichtsinhalt", value=generate_raw_report_context(rep_sel), height=300, disabled=True, label_visibility="collapsed")
                st.stop() # Bericht-Modus beendet hier das Skript komplett (kein Chatfenster!)
            
            else:
                if df_wissen is not None and not df_wissen.empty:
                    bez_spalte = "Bezeichnung" if "Bezeichnung" in df_wissen.columns else df_wissen.columns[0]
                    kat_spalte = "Wo?" if "Wo?" in df_wissen.columns else df_wissen.columns[1]
                    
                    # --- HARTE STRUKTUR FÜR "NEUE INFORMATION" ---
                    if st.session_state.aktive_rolle == "Host" and st.session_state.aktiver_use_case == "Neue Information":
                        alle_objekte = sorted([str(b).strip() for b in df_wissen[bez_spalte].dropna().drop_duplicates().tolist()])
                        if "Nicht gefunden" in alle_objekte: alle_objekte.remove("Nicht gefunden")
                        alle_objekte.append("Nicht gefunden")
                        
                        chat_abfrage_text = "Bitte gib hier den neuen Informationstext ein:"
                        
                        # Dropdown 1: Objekt-Auswahl mit isoliertem Key Namespace
                        val_obj = st.selectbox("🔎 Objekt auswählen:", options=alle_objekte, index=None, key=f"direct_obj_{st.session_state.aktiver_use_case}")
                        st.session_state.aktuell_gewaehltes_objekt = val_obj
                        
                        # HARD GATE: Wenn kein Objekt gewählt ist, brechen wir HIER ab -> Verhindert das Chatfeld!
                        if not val_obj:
                            st.stop()
                        
                        # Dropdown 2: Rendert absolut isoliert erst NACH dem Gate
                        spalten_options = get_datenspalten_options(df_wissen)
                        st.write("")
                        val_col = st.selectbox("📍 Bitte wähle die Art der Information aus:", options=spalten_options, index=None, key=f"direct_col_{st.session_state.aktiver_use_case}")
                        st.session_state[f"target_col_{st.session_state.aktiver_use_case}"] = val_col
                        
                        # HARD GATE 2: Wenn die Spalte noch nicht gewählt ist, stoppen -> Keine Chat-Eingabe erlauben!
                        if not val_col:
                            st.stop()
                    
                    # --- STANDARDFILTER FÜR ANDERE MODI ---
                    else:
                        STANDARD_DROPDOWNS = ["Ausstattung innen", "Ausstattung außen", "In der Nähe"]
                        for kat in STANDARD_DROPDOWNS:
                            if "innen" in kat.lower(): mask = df_wissen[kat_spalte].astype(str).str.contains("innen", case=False, na=False)
                            elif "außen" in kat.lower() or "aussen" in kat.lower(): mask = df_wissen[kat_spalte].astype(str).str.contains("außen|aussen", case=False, na=False)
                            else: mask = df_wissen[kat_spalte].astype(str).str.contains("nähe|naehe", case=False, na=False)
                            if st.session_state.aktive_rolle == "Gast" and "Relevanz Gast" in df_wissen.columns:
                                mask = mask & (df_wissen["Relevanz Gast"].astype(str).str.strip().str.lower() == "x")
                            
                            verfuegbare_bez = sorted([str(b).strip() for b in df_wissen[mask][bez_spalte].dropna().drop_duplicates().tolist()])
                            if "Nicht gefunden" in verfuegbare_bez: verfuegbare_bez.remove("Nicht gefunden")
                            verfuegbare_bez.append("Nicht gefunden")
                            
                            val = st.selectbox(f"🔎 {kat} wählen...", options=verfuegbare_bez, index=None, key=f"dropdown_{kat}_{st.session_state.aktiver_use_case}", label_visibility="collapsed")
                            if val: st.session_state.aktuell_gewaehltes_objekt = val

# ==============================================================================
# 7. CHAT FLOW LAYER (Wird nur erreicht, wenn alle Gates oben passiert wurden)
# ==============================================================================
st.write("---")
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]): st.markdown(msg["content"])

if user_input := st.chat_input(chat_abfrage_text):
    st.session_state.messages.append({"role": "user", "content": user_input})
    st.rerun()

if st.session_state.messages and st.session_state.messages[-1]["role"] == "user":
    user_input = st.session_state.messages[-1]["content"]
    
    if aktuelle_richtung == "OUTPUT":
        if st.session_state.aktuell_gewaehltes_objekt in [None, "Nicht gefunden"]:
            execute_transitional_routing(user_input, "Nicht gefunden")
            st.rerun()
        else:
            context_str = extract_context_for_object(st.session_state.aktuell_gewaehltes_objekt)
            res = call_gemini_api_structured(user_input, context_str)
            if res.wissensluecke_erkannt or res.antwort_text == "" or any(p in res.antwort_text.lower() for p in ["keine information", "leider nein"]):
                execute_transitional_routing(user_input, st.session_state.aktuell_gewaehltes_objekt)
            else:
                st.session_state.messages.append({"role": "assistant", "content": res.antwort_text})
            st.rerun()
            
    elif aktuelle_richtung == "INPUT":
        akt_col = st.session_state.get(f"target_col_{st.session_state.aktiver_use_case}")
        if st.session_state.aktiver_use_case == "Neue Information" and akt_col:
            execute_matrix_input_direct(akt_col, st.session_state.aktuell_gewaehltes_objekt, user_input)
        else:
            ziel_obj = st.session_state.aktuell_gewaehltes_objekt if st.session_state.aktuell_gewaehltes_objekt else "Nicht gefunden"
            execute_matrix_input(st.session_state.aktiver_use_case, ziel_obj, user_input)
            
        st.session_state.messages.append({"role": "assistant", "content": danke_text_template.replace("{use_case}", st.session_state.aktiver_use_case)})
        st.cache_data.clear()
        st.rerun()
