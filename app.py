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

# Unverrückbarer Fallback-Satz (Kapitel 1.3)
FALLBACK_SATZ = "Ich habe dazu leider keine Informationen, Ich gebe das aber gern an die Hosts weiter."

# Asymmetrischer Chat & Smartphone-Optimierung via CSS Injektion
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

# ==============================================================================
# 3. DATEN-LADE ENGINE (Drei-Blatt-Modell mit header=0 laut Spezifikation)
# ==============================================================================
@st.cache_data(ttl=30)
def load_dynamic_data():
    try:
        # Authentifizierung über den gesetzlich vorgeschriebenen Key (Kapitel 6.4.1)
        creds_dict = st.secrets["GOOGLE_CREDENTIALS"]
        creds = service_account.Credentials.from_service_account_info(creds_dict)
        service = build('drive', 'v3', credentials=creds)
        
        request = service.files().get_media(fileId=FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False: _, done = downloader.next_chunk()
            
        fh.seek(0)
        # Spaltenköpfe fix in Zeile 1
        df_wissen = pd.read_excel(fh, sheet_name="Wissensbasis", header=0)
        fh.seek(0)
        df_lexikon = pd.read_excel(fh, sheet_name="Spalten_Lexikon", header=0)
        fh.seek(0)
        df_usecases = pd.read_excel(fh, sheet_name="UseCase_Lexikon", header=0)
        
        # Geografische Vererbung via Forward-Fill
        if df_wissen is not None and not df_wissen.empty and "Wo?" in df_wissen.columns:
            df_wissen["Wo?"] = df_wissen["Wo?"].ffill()
            
        return df_wissen, df_lexikon, df_usecases, service
    except Exception as e:
        st.error(f"Kritischer Fehler beim Laden der Datenspezifikation: {e}")
        return None, None, None, None

with st.spinner("Synchronisiere mit der Excel-Zentralmatrix..."):
    df_wissen, df_lexikon, df_usecases, drive_service = load_dynamic_data()

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
        
        # Sicheres Zurückparsen des JSON-Strings in das typisierte Python-Objekt
        data = json.loads(response.text)
        return KiAntwortSchema(
            wissensluecke_erkannt=bool(data.get("wissensluecke_erkannt", True)),
            antwort_text=str(data.get("antwort_text", ""))
        )
    except Exception as e:
        return KiAntwortSchema(wissensluecke_erkannt=True, antwort_text="")

def call_gemini_api_raw(prompt, system_context=None):
    client = get_ki_client()
    if client is None: return "🛑 KI-Schnittstelle nicht konfiguriert."
    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash", 
            contents=prompt, 
            config=types.GenerateContentConfig(system_instruction=system_context, temperature=0.2)
        )
        return response.text
    except Exception as e:
        return f"🛑 KI-Fehler: {e}"

def extract_context_for_object(objekt_name):
    if df_wissen is None or df_lexikon is None or objekt_name is None: return ""
    
    bez_col = df_wissen.columns[0] if "Bezeichnung" not in df_wissen.columns else "Bezeichnung"
    row_match = df_wissen[df_wissen[bez_col].astype(str).str.strip().str.lower() == objekt_name.lower().strip()]
    if row_match.empty: return ""
    
    lex_spalten_name = df_lexikon.columns[0]
    rolle_idx = 2 if st.session_state.aktive_rolle == "Gast" else 3 
    lex_rollen_freigabe = df_lexikon.columns[rolle_idx]
    
    freigegebene_tags = df_lexikon[df_lexikon[lex_rollen_freigabe].astype(str).str.lower().str.strip() == "ja"][lex_spalten_name].tolist()
    
    context_parts = [f"Informationen zum Objekt: {objekt_name}"]
    for col in df_wissen.columns:
        if col in freigegebene_tags and col in row_match.columns:
            val = row_match.iloc[0][col]
            if pd.notna(val) and str(val).strip() != "":
                context_parts.append(f"- {col}: {str(val).strip()}")
                
    return "\n".join(context_parts)

# ==============================================================================
# 5. MATRIZEN-SCHREIBENGINE (openpyxl / Text in Blau #1F4E78 laut Kapitel 6.4.3)
# ==============================================================================
def execute_matrix_input(use_case_name, objekt_name, freitext):
    if drive_service is None or df_lexikon is None: return
    try:
        ziel_objekt = "Nicht gefunden" if (objekt_name is None or objekt_name == "Nicht gefunden") else objekt_name
        
        tag_col_name = df_lexikon.columns[4]
        lex_spalten_name = df_lexikon.columns[0]
        
        match_lex = df_lexikon[df_lexikon[tag_col_name].astype(str).str.lower().str.strip() == use_case_name.lower().strip()]
        if match_lex.empty:
            st.error(f"Administrativer Fehler: Kein Tag im Spalten_Lexikon für '{use_case_name}' definiert.")
            return
            
        physische_zielspalte = match_lex.iloc[0][lex_spalten_name]
        
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
            ws.cell(row=ziel_row_idx, column=col_bez_idx).value = "Nicht gefunden"
            
        ziel_col_idx = find_column_by_fuzzy_name(headers, physische_zielspalte)
        if not ziel_col_idx: return 
            
        zeitstempel = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
        nutzer = st.session_state.aktive_rolle
        alter_inhalt = ws.cell(row=ziel_row_idx, column=ziel_col_idx).value or ""
        
        neuer_eintrag = f"[{zeitstempel} | {nutzer}]: {freitext}"
        kompletter_text = f"{alter_inhalt}\n{neuer_eintrag}" if alter_inhalt else neuer_eintrag
            
        ziel_zelle = ws.cell(row=ziel_row_idx, column=ziel_col_idx)
        ziel_zelle.value = kompletter_text
        # Gesetzliche Formatierungsvorgaben (Blau & Zeilenumbruch)
        ziel_zelle.font = Font(color="1F4E78")
        ziel_zelle.alignment = Alignment(wrap_text=True)
        
        status_col_name = f"{physische_zielspalte} Status"
        status_col_idx = find_column_by_fuzzy_name(headers, status_col_name)
        
        if status_col_idx:
            status_wert = "aktiv" if "störung" in use_case_name.lower() else "offen"
            status_zelle = ws.cell(row=ziel_row_idx, column=status_col_idx)
            alter_status = status_zelle.value or ""
            neuer_status = f"[{zeitstempel}]: {status_wert}"
            status_zelle.value = f"{alter_status}\n{neuer_status}" if alter_status else neuer_status
            status_zelle.font = Font(color="1F4E78")
                
        output_stream = io.BytesIO()
        wb.save(output_stream)
        output_stream.seek(0)
        media = MediaIoBaseUpload(output_stream, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        drive_service.files().update(fileId=FILE_ID, media_body=media).execute()
        
    except Exception as e:
        st.error(f"Fehler beim Schreiben in die Zentralmatrix: {e}")

def execute_transitional_routing(user_input, objekt_name=None):
    st.session_state.messages.append({"role": "assistant", "content": FALLBACK_SATZ})
    ziel_obj = objekt_name if objekt_name else "Nicht gefunden"
    execute_matrix_input("Keine Information", ziel_obj, user_input)
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

def reset_chat_flow():
    st.session_state.messages = []

# ==============================================================================
# 6. HMI PRESENTATION LAYER (Deterministische Kaskadenführung)
# ==============================================================================
st.title("☀️ Villa Avatar")

if "aktive_rolle" not in st.session_state: st.session_state.aktive_rolle = None
if "aktiver_use_case" not in st.session_state: st.session_state.aktiver_use_case = None
if "messages" not in st.session_state: st.session_state.messages = []
if "bericht_filter" not in st.session_state: st.session_state.bericht_filter = None

# KASKADE 1: Rollen-Zuweisung
neue_rolle = st.selectbox("Rolle", options=["Gast", "Host"], index=None, placeholder="Wer bist du?", label_visibility="collapsed")
if neue_rolle != st.session_state.aktive_rolle:
    st.session_state.aktive_rolle = neue_rolle
    st.session_state.aktiver_use_case = None
    st.session_state.bericht_filter = None
    st.session_state.messages = []
    for key in list(st.session_state.keys()):
        if key.startswith("dropdown_"): del st.session_state[key]
    st.rerun()

aktuelles_objekt = None
aktuelle_richtung = None

if st.session_state.aktive_rolle and df_usecases is not None:
    st.write("---")
    uc_col = df_usecases.columns[0]
    dir_col = df_usecases.columns[1]
    hmi_col = df_usecases.columns[2]
    
    mask_sichtbar = df_usecases[hmi_col].astype(str).str.lower().str.strip() == "ja"
    verfuegbare_uc = df_usecases[mask_sichtbar][uc_col].tolist()
    
    if st.session_state.aktive_rolle == "Gast":
        erlaubte_buttons = [uc for uc in verfuegbare_uc if any(x in uc.lower() for x in ["hilfe", "störung", "feedback"])]
    else:
        erlaubte_buttons = verfuegbare_uc

    # KASKADE 2: Use Cases (Aktions-Buttons)
    cols = st.columns(len(erlaubte_buttons))
    neuer_use_case = st.session_state.aktiver_use_case
    
    for idx, uc_name in enumerate(erlaubte_buttons):
        with cols[idx]:
            is_active = (st.session_state.aktiver_use_case == uc_name)
            if st.button(uc_name, use_container_width=True, type="primary" if is_active else "secondary"):
                neuer_use_case = uc_name

    if neuer_use_case != st.session_state.aktiver_use_case:
        st.session_state.aktiver_use_case = neuer_use_case
        st.session_state.bericht_filter = None
        st.session_state.messages = []
        for key in list(st.session_state.keys()):
            if key.startswith("dropdown_"): del st.session_state[key]
        st.rerun()

    # KASKADE 3: Kontextabhängige Dropdowns & Report UI
    if st.session_state.aktiver_use_case:
        uc_row = df_usecases[df_usecases[uc_col].astype(str).str.lower().str.strip() == st.session_state.aktiver_use_case.lower().strip()]
        if not uc_row.empty:
            aktuelle_richtung = str(uc_row.iloc[0][dir_col]).strip().upper()
            
            if "bericht" in st.session_state.aktiver_use_case.lower():
                st.write("")
                b_col1, b_col2 = st.columns(2)
                b_col3, b_col4 = st.columns(2)
                b_col5, b_col6 = st.columns(2)
                
                with b_col1:
                    if st.button("⚠️ Offene Störungen", use_container_width=True): st.session_state.bericht_filter = "offene_stoerungen"; st.rerun()
                with b_col2:
                    if st.button("✅ Behobene Störungen", use_container_width=True): st.session_state.bericht_filter = "behobene_stoerungen"; st.rerun()
                with b_col3:
                    if st.button("💡 Offenes Feedback", use_container_width=True): st.session_state.bericht_filter = "offenes_feedback"; st.rerun()
                with b_col4:
                    if st.button("❌ Ignoriertes Feedback", use_container_width=True): st.session_state.bericht_filter = "ignoriertes_feedback"; st.rerun()
                with b_col5:
                    if st.button("🔍 Offene Wissenslücken", use_container_width=True): st.session_state.bericht_filter = "offene_luecken"; st.rerun()
                with b_col6:
                    if st.button("📋 Gesamtübersicht", use_container_width=True): st.session_state.bericht_filter = "gesamtuebersicht"; st.rerun()
            else:
                STANDARD_DROPDOWNS = ["Ausstattung innen", "Ausstattung außen", "In der Nähe"]
                if df_wissen is not None and not df_wissen.empty:
                    bez_spalte = df_wissen.columns[0] if "Bezeichnung" not in df_wissen.columns else "Bezeichnung"
                    kat_spalte = df_wissen.columns[1] if "Wo?" not in df_wissen.columns else "Wo?"
                    
                    st.write("")
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
                        st.selectbox(label=f"hidden_{kat}", options=verfuegbare_bez, index=None, placeholder=f"🔎 {kat} wählen...", key=dp_key, on_change=reset_chat_flow, label_visibility="collapsed")
                            
                    for kat in STANDARD_DROPDOWNS:
                        val = st.session_state.get(f"dropdown_{kat}_{st.session_state.aktiver_use_case}")
                        if val is not None:
                            aktuelles_objekt = val
                            break

# ==============================================================================
# 7. CHAT FLOW & DETERMINISTISCHES ROUTING (Mit doppeltem Python-Sicherheitsnetz)
# ==============================================================================
if st.session_state.aktiver_use_case and "bericht" in st.session_state.aktiver_use_case.lower() and st.session_state.bericht_filter:
    with st.spinner("Analysiere Datenbasis..."):
        report_data_str = generate_raw_report_context(st.session_state.bericht_filter)
        if "Keine passenden Einträge" in report_data_str:
            report_output = f"Aktuell liegen keine Einträge für den Filter '{st.session_state.bericht_filter}' vor. ☀️"
        else:
            prompt = f"Du bist der administrative Analyst. Strukturiere diese Matrix-Daten professionell und chronologisch für den Host:\n\n{report_data_str}"
            report_output = call_gemini_api_raw(prompt, system_context="Liste Fakten auf, nutze Bulletpoints, bleibe sachlich.")
        
        st.session_state.messages.append({"role": "assistant", "content": report_output})
        st.session_state.bericht_filter = None
        st.rerun()

st.write("---")
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]): st.markdown(msg["content"])

if st.session_state.aktiver_use_case and "bericht" not in st.session_state.aktiver_use_case.lower():
    if user_input := st.chat_input("Wie kann ich dir helfen?"):
        st.session_state.messages.append({"role": "user", "content": user_input})
        st.rerun()

if st.session_state.messages and st.session_state.messages[-1]["role"] == "user":
    user_input = st.session_state.messages[-1]["content"]
    
    # OUTPUT-PFAD (Suchen, Synthetisieren und Verifizieren)
    if aktuelle_richtung == "OUTPUT":
        if aktuelles_objekt is None or aktuelles_objekt == "Nicht gefunden":
            with st.spinner("Transitional Routing aktiv..."):
                execute_transitional_routing(user_input, aktuelles_objekt)
                st.rerun()
        else:
            with st.spinner("Durchsuche Matrix..."):
                context_str = extract_context_for_object(aktuelles_objekt)
                
                # Hole die strukturierte Antwort von der KI
                structured_response = call_gemini_api_structured(user_input, context_str)
                
                # Lokale Extraktion für unbestechliche Variablensicherheit
                ist_luecke = structured_response.wissensluecke_erkannt
                ki_text = structured_response.antwort_text
                ki_text_lower = ki_text.lower()
                
                # Phrasen-Sicherheitsnetz gegen "versteckte" Absagen im Freitext
                luecken_phrasen = [
                    "keine information", 
                    "weiß ich nicht", 
                    "nicht hinterlegt", 
                    "leider nein", 
                    "nicht bekannt", 
                    "fehlen mir details",
                    "gern an die hosts weiter"
                ]
                
                # DETERMINISTISCHE ENTSCHEIDUNG IN PYTHON:
                # Pfad schlägt zu, wenn das Flag True ist ODER das Textfeld leer bleibt ODER eine Lücken-Phrase auftaucht
                if ist_luecke or ki_text == "" or any(phrase in ki_text_lower for phrase in luecken_phrasen):
                    with st.spinner("Sicherheitsnetz aktiv: Wissenslücke detektiert. Protokolliere..."):
                        execute_transitional_routing(user_input, aktuelles_objekt)
                else:
                    # Wissen existiert nachweislich und fehlerfrei, zeige Antwort an
                    st.session_state.messages.append({"role": "assistant", "content": ki_text})
                
                st.rerun()
    
    # INPUT-PFAD (Direktes Schreiben in die Matrix)
    elif aktuelle_richtung == "INPUT":
        with st.spinner("Protokolliere Eintrag in der Matrix..."):
            execute_matrix_input(st.session_state.aktiver_use_case, aktuelles_objekt, user_input)
            danke_satz = f"Vielen Dank! Ich habe deine Eingabe zum Thema '{st.session_state.aktiver_use_case}' für die Hosts eingetragen."
            st.session_state.messages.append({"role": "assistant", "content": danke_satz})
            st.cache_data.clear()
            st.rerun()
