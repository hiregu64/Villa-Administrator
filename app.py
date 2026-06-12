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
    return [c for c in df.columns if "status" not in c.lower() and c.lower().strip() not in geschuetzt]

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
    
    df_wissen.columns = [str(c).strip() for c in df_wissen.columns]
    df_lexikon.columns = [str(c).strip() for c in df_lexikon.columns]
    
    bez_col = df_wissen.columns[0] if "Bezeichnung" not in df_wissen.columns else "Bezeichnung"
    row_match = df_wissen[df_wissen[bez_col].astype(str).str.strip().str.lower() == objekt_name.lower().strip()]
    if row_match.empty: 
        st.session_state.last_extracted_context = f"Kein Treffer in 'Bezeichnung' für '{objekt_name}'."
        return ""
    
    aktuelle_rolle = str(st.session_state.get("aktive_rolle", "Gast")).strip().lower()
    context_parts = [f"Informationen zum Objekt: {objekt_name}"]

    if aktuelle_rolle == "host":
        for col in df_wissen.columns:
            if col != bez_col and "status" not in col.lower() and col.lower() != "wo?":
                val = row_match.iloc[0][col]
                if pd.notna(val) and str(val).strip() != "":
                    context_parts.append(f"- {col}: {str(val).strip()}")
        st.session_state.last_extracted_context = "\n".join(context_parts)
    else:
        lex_spalten_name = df_lexikon.columns[0]
        gast_freigabe_spalte = None
        for col in df_lexikon.columns:
            if "gast" in col.lower():
                gast_freigabe_spalte = col
                break
        
        if not gast_freigabe_spalte:
            gast_freigabe_spalte = df_lexikon.columns[3] if len(df_lexikon.columns) > 3 else df_lexikon.columns[-1]
        
        mask_ja = df_lexikon[gast_freigabe_spalte].astype(str).str.lower().str.strip() == "ja"
        freigegebene_tags = df_lexikon[mask_ja][lex_spalten_name].astype(str).str.strip().tolist()
        
        for col in df_wissen.columns:
            is_freigegeben = any(col.lower() == tag.lower() for tag in freigegebene_tags)
            if is_freigegeben and col in row_match.columns:
                val = row_match.iloc[0][col]
                if pd.notna(val) and str(val).strip() != "":
                    context_parts.append(f"- {col}: {str(val).strip()}")
                    
        if len(context_parts) <= 1:
            st.session_state.last_extracted_context = (
                f"⚠️ Objekt '{objekt_name}' gefunden, aber keine Spalte für den Gast freigegeben.\n"
                f"Ausgewertete Lexikon-Spalte: '{gast_freigabe_spalte}'\n"
                f"Gefundene Freigabe-Tags: {freigegebene_tags}"
            )
        else:
            st.session_state.last_extracted_context = "\n".join(context_parts)
                
    return "\n".join(context_parts)

# ==============================================================================
# 5. MATRIZEN-SCHREIBENGINE (openpyxl)
# ==============================================================================
def execute_matrix_input_direct(physische_zielspalte, objekt_name, freitext):
    if drive_service is None or df_wissen is None:
        st.session_state.last_write_status = "🛑 Schreibfehler: Drive-Service oder Datenbasis nicht geladen."
        return
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
        if not ziel_col_idx: 
            st.session_state.last_write_status = f"🛑 Spalte '{physische_zielspalte}' in Excel nicht gefunden!"
            return 
            
        zeitstempel = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
        nutzer = st.session_state.aktive_rolle if st.session_state.aktive_rolle else "System"
        alter_inhalt = ws.cell(row=ziel_row_idx, column=ziel_col_idx).value or ""
        
        neuer_eintrag = f"[{zeitstempel} | {nutzer}]: {freitext}"
        kompletter_text = f"{alter_inhalt}\n{neuer_eintrag}" if alter_inhalt else neuer_eintrag
            
        ziel_zelle = ws.cell(row=ziel_row_idx, column=ziel_col_idx)
        ziel_zelle.value = kompletter_text
        ziel_zelle.font = Font(color="1F4E78")
        ziel_zelle.alignment = Alignment(wrap_text=True)
        
        status_col_name = f"{physische_zielspalte} Status"
        status_col_idx = find_column_by_fuzzy_name(headers, status_col_name)
        
        if status_col_idx:
            status_wert = "aktiv" if "störung" in physische_zielspalte.lower() else "offen"
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
        
        st.session_state.last_write_status = f"✅ ERFOLG: Zeile {ziel_row_idx}, Spalte '{physische_zielspalte}' beschrieben um {zeitstempel}."
        st.toast("✅ Excel-Zentralmatrix aktualisiert!")
        
    except Exception as e:
        st.session_state.last_write_status = f"🛑 GDrive-API-Fehler beim Schreibvorgang: {e}"

def execute_matrix_input(use_case_name, objekt_name, freitext):
    if df_lexikon is None: return
    tag_col_name = df_lexikon.columns[4]
    lex_spalten_name = df_lexikon.columns[0]
    
    physische_zielspalte = None
    for _, row in df_lexikon.iterrows():
        tags_in_row = [t.strip().lower() for t in str(row[tag_col_name]).split(',')]
        if use_case_name.lower().strip() in tags_in_row:
            physische_zielspalte = str(row[lex_spalten_name]).strip()
            break
            
    if not physische_zielspalte and use_case_name == "Keine Information":
        for col in df_wissen.columns:
            if "information" in col.lower() and "status" not in col.lower():
                physische_zielspalte = col
                break
                
    if physische_zielspalte:
        execute_matrix_input_direct(physische_zielspalte, objekt_name, freitext)
    else:
        st.session_state.last_write_status = f"🛑 Admin-Fehler: Kein Tag im Spalten_Lexikon für '{use_case_name}'."

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
# 6. HMI PRESENTATION LAYER (Direkter, unkomplizierter Selektions-Reset)
# ==============================================================================
if "aktive_rolle" not in st.session_state: st.session_state.aktive_rolle = None
if "aktiver_use_case" not in st.session_state: st.session_state.aktiver_use_case = None
if "messages" not in st.session_state: st.session_state.messages = []
if "bericht_filter" not in st.session_state: st.session_state.bericht_filter = None
if "host_authentifiziert" not in st.session_state: st.session_state.host_authentifiziert = False
if "debug_modus_aktiv" not in st.session_state: st.session_state.debug_modus_aktiv = False

# Das Dropdown triggert den Reset sofort bei JEDER expliziten Interaktion des Nutzers
neue_rolle = st.selectbox("Rolle", options=["Gast", "Host"], index=None, placeholder="Wer bist du?", label_visibility="collapsed")

if neue_rolle is not None:
    # Sobald der Nutzer klickt, wird gnadenlos alles darunter gelöscht und auf Anfang gesetzt
    if neue_rolle != st.session_state.aktive_rolle:
        st.session_state.aktive_rolle = neue_rolle
        st.session_state.aktiver_use_case = None
        st.session_state.bericht_filter = None
        st.session_state.host_authentifiziert = False
        st.session_state.debug_modus_aktiv = False
        st.session_state.messages = []
        for key in list(st.session_state.keys()):
            if key.startswith("dropdown_") or key.startswith("target_col_"): del st.session_state[key]
        st.rerun()

# Wenn die Rolle Host gewählt wurde, aber noch nicht verifiziert ist, erzwingen wir die Passworteingabe
if st.session_state.aktive_rolle == "Host" and not st.session_state.host_authentifiziert:
    st.write("---")
    pwd_input = st.text_input("🔑 Bitte Passwort für Host-Sicht eingeben:", type="password", key="host_pwd_field")
    
    if pwd_input:
        if df_passwoerter is not None and not df_passwoerter.empty:
            p_rolle_col = df_passwoerter.columns[0]
            p_pwd_col = df_passwoerter.columns[1]
            p_func_col = df_passwoerter.columns[2] if len(df_passwoerter.columns) > 2 else None
            
            erster_rollen_eintrag = str(df_passwoerter.iloc[0][p_rolle_col]).strip()
            host_rows = df_passwoerter[df_passwoerter[p_rolle_col].astype(str).str.strip().str.lower() == erster_rollen_eintrag.lower()]
            
            treffer_gefunden = False
            for _, row in host_rows.iterrows():
                gespeichertes_pwd = str(row[p_pwd_col]).strip()
                if pwd_input.strip() == gespeichertes_pwd:
                    st.session_state.host_authentifiziert = True
                    # Wir überschreiben die aktive_rolle mit einem Marker, damit bei erneutem Klick auf das Dropdown der Reset wieder anspringt!
                    st.session_state.aktive_rolle = "Host_Verifiziert"
                    treffer_gefunden = True
                    
                    if p_func_col and pd.notna(row[p_func_col]):
                        debug_kriterium = "debug"
                        if len(host_rows) > 1:
                            debug_kriterium = str(host_rows.iloc[1][p_func_col]).strip().lower()
                        
                        funktion_wert = str(row[p_func_col]).strip().lower()
                        if funktion_wert == debug_kriterium:
                            st.session_state.debug_modus_aktiv = True
                    break
            
            if treffer_gefunden:
                st.success("Erfolgreich eingeloggt!")
                st.rerun()
            else:
                st.error("❌ Falsches Passwort. Zugriff verweigert.")
        else:
            st.error("🛑 Passwort-Matrix 'Passwort_Lexikon' nicht geladen oder leer.")
    st.stop()

# INITIALISIERUNG DER DATENREDUKTIONS-VARIABLEN
aktuelles_objekt = None
aktuelle_richtung = None
gewaehlte_direktspalte = None

chat_abfrage_text = "Wie kann ich dir helfen?"
danke_text_template = "Vielen Dank! Ich habe deine Eingabe zum Thema '{use_case}' für die Hosts eingetragen."

# Ab hier gilt die Host-Rolle für den restlichen UI-Aufbau als aktiv
if st.session_state.aktive_rolle in ["Host", "Host_Verifiziert"] or st.session_state.aktive_rolle == "Gast":
    if df_usecases is not None:
        st.write("---")
        
        uc_col = df_usecases.columns[0]
        dir_col = df_usecases.columns[1]
        hmi_col = df_usecases.columns[2]
        
        btn_col = df_usecases.columns[3] if len(df_usecases.columns) > 3 else None
        prompt_col = df_usecases.columns[4] if len(df_usecases.columns) > 4 else None
        danke_col = df_usecases.columns[5] if len(df_usecases.columns) > 5 else None
        
        mask_sichtbar = df_usecases[hmi_col].astype(str).str.lower().str.strip() == "ja"
        verfuegbare_uc = df_usecases[mask_sichtbar][uc_col].tolist()
        
        if st.session_state.aktive_rolle == "Gast":
            erlaubte_buttons = [uc for uc in verfuegbare_uc if any(x in uc.lower() for x in ["hilfe", "störung", "feedback"])]
        else:
            erlaubte_buttons = verfuegbare_uc

        cols = st.columns(len(erlaubte_buttons))
        neuer_use_case = st.session_state.aktiver_use_case
        
        for idx, uc_name in enumerate(erlaubte_buttons):
            with cols[idx]:
                is_active = (st.session_state.aktiver_use_case == uc_name)
                button_label = uc_name
                if btn_col:
                    btn_match = df_usecases[df_usecases[uc_col].astype(str).str.strip() == str(uc_name).strip()]
                    if not btn_match.empty and pd.notna(btn_match.iloc[0][btn_col]):
                        button_label = str(btn_match.iloc[0][btn_col]).strip()
                
                if st.button(button_label, use_container_width=True, type="primary" if is_active else "secondary", key=f"btn_{uc_name}"):
                    neuer_use_case = uc_name

        if neuer_use_case != st.session_state.aktiver_use_case:
            st.session_state.aktiver_use_case = neuer_use_case
            st.session_state.bericht_filter = None
            st.session_state.messages = []
            for key in list(st.session_state.keys()):
                if key.startswith("dropdown_") or key.startswith("target_col_"): del st.session_state[key]
            st.rerun()

        if st.session_state.aktiver_use_case:
            uc_row = df_usecases[df_usecases[uc_col].astype(str).str.lower().str.strip() == st.session_state.aktiver_use_case.lower().strip()]
            if not uc_row.empty:
                aktuelle_richtung = str(uc_row.iloc[0][dir_col]).strip().upper()
                
                if prompt_col and pd.notna(uc_row.iloc[0][prompt_col]):
                    chat_abfrage_text = str(uc_row.iloc[0][prompt_col]).strip()
                    
                if danke_col and pd.notna(uc_row.iloc[0][danke_col]):
                    danke_text_template = str(uc_row.iloc[0][danke_col]).strip()
                
                # KASKADE FÜR BERICHTSTELLUNG
                if "bericht" in st.session_state.aktiver_use_case.lower():
                    st.write("")
                    bericht_optionen = {
                        "⚠️ Offene Störungen": "offene_stoerungen",
                        "✅ Behobene Störungen": "behobene_stoerungen",
                        "💡 Offenes Feedback": "offenes_feedback",
                        "❌ Ignoriertes Feedback": "ignoriertes_feedback",
                        "🔍 Offene Wissenslücken": "offene_luecken",
                        "📋 Gesamtübersicht": "gesamtuebersicht"
                    }
                    
                    gewaehlter_bericht_label = st.selectbox(
                        label="Gewünschten System-Bericht auswählen:",
                        options=list(bericht_optionen.keys()),
                        index=None,
                        placeholder="📋 Berichtstyp wählen...",
                        label_visibility="collapsed"
                    )
                    if gewaehlter_bericht_label:
                        st.session_state.bericht_filter = bericht_optionen[gewaehlter_bericht_label]
                
                # KASKADE FÜR STANDARD USE CASES
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
                        
                        # Spaltenauswahl-Dropdown für "Neue Information"
                        if st.session_state.aktive_rolle == "Host_Verifiziert" and st.session_state.aktiver_use_case == "Neue Information" and aktuelles_objekt:
                            spalten_options = get_datenspalten_options(df_wissen)
                            col_key = f"target_col_{st.session_state.aktiver_use_case}"
                            gewaehlte_direktspalte = st.selectbox(
                                label="Zieldokumentation Spalte wählen:",
                                options=spalten_options,
                                index=None,
                                placeholder="📍 In welche Matrix-Spalte soll dokumentiert werden?...",
                                key="host_direct_col_select"
                            )
                            if gewaehlte_direktspalte:
                                st.session_state[col_key] = gewaehlte_direktspalte
                            else:
                                gewaehlte_direktspalte = st.session_state.get(col_key)

# ==============================================================================
# 6.5 SYSTEM-DIAGNOSE MONITOR
# ==============================================================================
if st.session_state.debug_modus_aktiv:
    st.write("")
    with st.expander("🔍 SYSTEM-DIAGNOSE MONITOR (Laufzeit-Metriken)", expanded=True):
        d_col1, d_col2 = st.columns(2)
        with d_col1:
            st.metric(label="1. Aktive Rolle", value=str(st.session_state.get("aktive_rolle", "None")))
            st.metric(label="3. Gewähltes Objekt", value=str(aktuelles_objekt))
            st.metric(label="4. Direkt-Zielspalte (Host)", value=str(gewaehlte_direktspalte))
        with d_col2:
            st.metric(label="2. Use Case | Richtung", value=f"{st.session_state.get('aktiver_use_case')} | {aktuelle_richtung}")
        
        target_display_name = "Details Nutzung"
        if df_lexikon is not None and not df_lexikon.empty:
            target_display_name = str(df_lexikon.iloc[0, 0]).strip()
            
        st.write(f"**5. Letzter Matrix-Lesestatus Spalte '{target_display_name}':**")
        if df_wissen is not None and not df_wissen.empty:
            st.success(f"✅ Daten erfolgreich geladen ({len(df_wissen)} Objekte in Matrix verifiziert)")
        else:
            st.error(f"🛑 Lesefehler: Spalte '{target_display_name}' nicht synchronisiert.")
        
        st.write("**6. Letzter Matrix-Schreibstatus:**")
        st.info(st.session_state.get("last_write_status", "Kein Status"))
        
        st.write("**7. Letzter Kontext-Extrakt (KI-Input):**")
        st.text_area(label="Matrix-Rohdaten", value=st.session_state.get("last_extracted_context", ""), height=100, disabled=True, label_visibility="collapsed")

# ==============================================================================
# 7. CHAT FLOW & DETERMINISTISCHES ROUTING
# ==============================================================================
if st.session_state.aktiver_use_case and "bericht" in st.session_state.aktiver_use_case.lower() and st.session_state.bericht_filter:
    with st.spinner("Analysiere Datenbasis und generiere Systembericht..."):
        report_data_str = generate_raw_report_context(st.session_state.bericht_filter)
        if "Keine passenden Einträge" in report_data_str:
            report_output = f"Aktuell liegen keine Einträge für den Filter '{st.session_state.bericht_filter}' vor. ☀️"
        else:
            prompt = f"Du bist der administrative Analyst. Strukturiere diese Matrix-Daten professionell und chronologisch für den Host:\n\n{report_data_str}"
            report_output = call_gemini_api_raw(prompt, system_context="Liste Fakten auf, nutze Bulletpoints, bleibe sachlich.")
        
        st.session_state.messages.append({"role": "assistant", "content": report_output})
        st.session_state.bericht_filter = None

st.write("---")
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]): st.markdown(msg["content"])

zeige_chat_input = False
if st.session_state.aktiver_use_case and "bericht" not in st.session_state.aktiver_use_case.lower():
    if aktuelle_richtung == "OUTPUT":
        zeige_chat_input = True
    elif aktuelle_richtung == "INPUT":
        if st.session_state.aktiver_use_case == "Neue Information":
            if aktuelles_objekt and gewaehlte_direktspalte:
                zeige_chat_input = True
        else:
            zeige_chat_input = True

if zeige_chat_input:
    if user_input := st.chat_input(chat_abfrage_text):
        st.session_state.messages.append({"role": "user", "content": user_input})
        st.rerun()

if st.session_state.messages and st.session_state.messages[-1]["role"] == "user":
    user_input = st.session_state.messages[-1]["content"]
    
    if aktuelle_richtung == "OUTPUT":
        if aktuelles_objekt is None or aktuelles_objekt == "Nicht gefunden":
            with st.spinner("Transitional Routing active..."):
                execute_transitional_routing(user_input, "Nicht gefunden")
                st.rerun()
        else:
            with st.spinner("Durchsuche Matrix..."):
                context_str = extract_context_for_object(aktuelles_objekt)
                structured_response = call_gemini_api_structured(user_input, context_str)
                
                ist_luecke = structured_response.wissensluecke_erkannt
                ki_text = structured_response.antwort_text
                ki_text_lower = ki_text.lower()
                
                luecken_phrasen = [
                    "keine information", "weiß ich nicht", "nicht hinterlegt", 
                    "leider nein", "nicht bekannt", "fehlen mir details",
                    "gern an die hosts weiter"
                ]
                
                if ist_luecke or ki_text == "" or any(phrase in ki_text_lower for phrase in luecken_phrasen):
                    with st.spinner("Sicherheitsnetz aktiv: Wissenslücke detektiert. Protokolliere..."):
                        execute_transitional_routing(user_input, aktuelles_objekt)
                else:
                    st.session_state.messages.append({"role": "assistant", "content": ki_text})
                st.rerun()
    
    elif aktuelle_richtung == "INPUT":
        with st.spinner("Protokolliere Eintrag in der Matrix..."):
            if st.session_state.aktiver_use_case == "Neue Information" and gewaehlte_direktspalte:
                execute_matrix_input_direct(gewaehlte_direktspalte, aktuelles_objekt, user_input)
            else:
                ziel_obj = aktuelles_objekt if aktuelles_objekt else "Nicht gefunden"
                execute_matrix_input(st.session_state.aktiver_use_case, ziel_obj, user_input)
                
            danke_satz = danke_text_template.replace("{use_case}", st.session_state.aktiver_use_case)
            st.session_state.messages.append({"role": "assistant", "content": danke_satz})
            st.cache_data.clear()
            st.rerun()
