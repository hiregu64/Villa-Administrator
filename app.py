import streamlit as st
import pandas as pd
import datetime
from google import genai
from google.genai import types
from google.oauth2 import service_account
from googleapiclient.discovery import build
import openpyxl
from io import BytesIO

# ==============================================================================
# 1. INITIALISIERUNG & CONFIG
# ==============================================================================
st.set_page_config(page_title="Villa Avatar", page_icon="☀️", layout="centered")

FALLBACK_SATZ = "Ich habe dazu leider keine Informationen, Ich gebe das aber gern an die Hosts weiter."

# Google Drive File ID deiner Excel-Zentralmatrix
SPREADSHEET_ID = "1hQrxRD4Jpeq_FMwAvfsOANT72ilCgRuxPEbMc1JcPPo"

# Session State initialisieren
if "messages" not in st.session_state:
    st.session_state.messages = []
if "aktive_rolle" not in st.session_state:
    st.session_state.aktive_rolle = None
if "aktiver_use_case" not in st.session_state:
    st.session_state.aktiver_use_case = None
if "last_write_status" not in st.session_state:
    st.session_state.last_write_status = "Kein Status"
if "last_extracted_context" not in st.session_state:
    st.session_state.last_extracted_context = ""

# ==============================================================================
# 2. API-VERBINDUNGEN & CACHING
# ==============================================================================
def get_gdrive_service():
    """Authentifiziert sich über die Streamlit Secrets mit dem Service Account."""
    creds_dict = st.secrets["GOOGLE_CREDENTIALS"]
    creds = service_account.Credentials.from_service_account_info(
        creds_dict, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)

def get_genai_client():
    """Initialisiert den offiziellen Google GenAI SDK Client."""
    return genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

@st.cache_data(ttl=30)
def load_excel_from_drive():
    """Lädt die Excel-Zentralmatrix im Binärstrom herunter (30s Cache)."""
    try:
        service = get_gdrive_service()
        request = service.files().get_media(fileId=SPREADSHEET_ID)
        file_stream = BytesIO(request.execute())
        return file_stream.getvalue()
    except Exception as e:
        st.error(f"Kritischer Fehler beim Laden der Zentralmatrix: {str(e)}")
        return None

def clean_string(val):
    """Bereinigt Strings für robusten Abgleich (String-Sanitizing)."""
    if pd.isna(val):
        return ""
    return str(val).strip().lower()

def sanitize_header(header_name):
    """
    Entfernt erklärende Zusätze in eckigen Klammern aus den Spaltenüberschriften,
    um Kaltstart-Fehler und administrative Notizen zu tolerieren.
    Beispiel: 'Button_Label [Text auf dem Button]' -> 'button_label'
    """
    s = str(header_name).split("[")[0]
    return s.strip().lower()

# ==============================================================================
# 3. DATEN-EXTRAKTION (DATAFRAMES)
# ==============================================================================
excel_bytes = load_excel_from_drive()

df_wissen = None
df_spalten = None
df_usecase = None

if excel_bytes:
    # 3.1 Wissensbasis laden
    df_wissen = pd.read_excel(BytesIO(excel_bytes), sheet_name="Wissensbasis")
    
    # 3.2 Spalten_Lexikon laden & Überschriften bereinigen
    df_spalten = pd.read_excel(BytesIO(excel_bytes), sheet_name="Spalten_Lexikon")
    df_spalten.columns = [sanitize_header(c) for c in df_spalten.columns]
    
    # 3.3 UseCase_Lexikon laden & Überschriften bereinigen
    df_usecase = pd.read_excel(BytesIO(excel_bytes), sheet_name="UseCase_Lexikon")
    df_usecase.columns = [sanitize_header(c) for c in df_usecase.columns]

# ==============================================================================
# 4. EXCEL MATRIX-SCHREIB-ENGINE (INPUT-PFAD)
# ==============================================================================
def execute_matrix_input(use_case_name, objekt_name, text_content):
    """Schreibt Benutzerdaten über openpyxl im Binärstrom zurück in die Matrix."""
    global excel_bytes
    try:
        # Aktuelle Daten holen
        live_bytes = load_excel_from_drive()
        wb = openpyxl.load_workbook(BytesIO(live_bytes))
        ws = wb["Wissensbasis"]
        
        # 1. Zielspalte über das Spalten_Lexikon ermitteln
        # Wir suchen die Spalte, die dem Use Case semantisch zugeordnet ist
        target_column_name = None
        status_column_name = None
        
        # Durchsuche das Spalten-Lexikon nach dem Tag
        for _, row in df_spalten.iterrows():
            zuordnung = str(row.get("zuordnung use case / richtung", "")).strip()
            if use_case_name.lower() in zuordnung.lower():
                if "status" in zuordnung.lower():
                    status_column_name = str(row.get("spaltenname in der wissensbasis", "")).strip()
                else:
                    target_column_name = str(row.get("spaltenname in der wissensbasis", "")).strip()
        
        if not target_column_name:
            st.session_state.last_write_status = f"Fehler: Keine Zielspalte für Use Case '{use_case_name}' definiert."
            return
        
        # Finden der physikalischen Spalten-Indizes in der Wissensbasis
        headers = [str(cell.value).strip() for cell in ws[1]]
        
        # Spalten-Matching mit Toleranz
        target_idx = None
        status_idx = None
        for i, h in enumerate(headers):
            if clean_string(h) == clean_string(target_column_name):
                target_idx = i + 1
            if status_column_name and clean_string(h) == clean_string(status_column_name):
                status_idx = i + 1
                
        # 2. Zielzeile ermitteln (Objekt-Matching oder Catch-All "Nicht gefunden")
        target_row_idx = None
        fallback_row_idx = None
        
        for r in range(2, ws.max_row + 1):
            obj_val = str(ws.cell(row=r, column=1).value).strip()
            if clean_string(obj_val) == clean_string(objekt_name) and objekt_name:
                target_row_idx = r
                break
            if clean_string(obj_val) == "nicht gefunden":
                fallback_row_idx = r
                
        final_row = target_row_idx if target_row_idx else fallback_row_idx
        
        if not final_row:
            st.session_state.last_write_status = "Fehler: Weder Objekt-Zeile noch 'Nicht gefunden'-Anker ermittelt."
            return
        
        # 3. Daten schreiben (Chronologischer Append & Formatierung)
        zeitstempel = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
        rolle = st.session_state.aktive_rolle if st.session_state.aktive_rolle else "Unbekannt"
        
        # Text-Zelle manipulieren
        cell = ws.cell(row=final_row, column=target_idx)
        alter_inhalt = str(cell.value) if cell.value is not None else ""
        neuer_eintrag = f"[{zeitstempel} - {rolle}]: {text_content}"
        cell.value = f"{alter_inhalt}\n{neuer_eintrag}".strip()
        
        # Blau färben (#1F4E78) & Textumbruch aktivieren
        cell.font = openpyxl.styles.Font(color="1F4E78", name="Arial")
        cell.alignment = openpyxl.styles.Alignment(wrap_text=True)
        
        # Status-Zelle manipulieren falls vorhanden
        if status_idx:
            status_cell = ws.cell(row=final_row, column=status_idx)
            status_cell.value = "aktiv" if use_case_name.lower() == "störung" else "offen"
            status_cell.font = openpyxl.styles.Font(color="1F4E78", name="Arial")
            status_cell.alignment = openpyxl.styles.Alignment(wrap_text=True)
            
        # 4. Datei zurück auf Google Drive hochladen
        out_stream = BytesIO()
        wb.save(out_stream)
        out_stream.seek(0)
        
        service = get_gdrive_service()
        from googleapiclient.http import MediaIoBaseUpload
        media = MediaIoBaseUpload(out_stream, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", resumable=True)
        service.files().update(fileId=SPREADSHEET_ID, media_body=media).execute()
        
        # Cache explizit leeren für Datenkonsistenz
        st.cache_data.clear()
        st.session_state.last_write_status = f"Erfolgreich geschrieben in Zeile {final_row}, Spalte {target_idx} am {zeitstempel}"
        
    except Exception as e:
        st.session_state.last_write_status = f"API-Schreibfehler: {str(e)}"

# ==============================================================================
# 5. CORE KI-LOGIK (STRUCTURED OUTPUTS)
# ==============================================================================
def ask_villa_avatar(nutzer_frage, extrahierter_kontext):
    """Sendet Kontext und Frage an Gemini unter Verwendung erzwungener Schemata."""
    client = get_genai_client()
    
    system_instruction = (
        "Du bist Villa Avatar, der präzise digitale Klon-Helfer für die Immobilie.\n"
        "Deine Tonalität ist kurz, freundlich und smartphone-optimiert.\n"
        "Halte dich strikt an den bereitgestellten Kontext. Erfinde niemals Daten (Halluzinationsverbot).\n"
        "Erwähne niemals interne Strukturen, Spalten- oder Dateinamen gegenüber dem Nutzer.\n"
        "Wenn der Kontext die Frage nicht beantworten kann, setze 'wissensluecke_erkannt' auf True."
    )
    
    prompt = f"Excel-Kontext:\n{extrahierter_kontext}\n\nFrage des Nutzers: {nutzer_frage}"
    
    # JSON-Schema erzwingen (Structured Outputs)
    class VillaResponse(types.BaseModel):
        wissensluecke_erkannt: bool
        antwort_text: str

    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(
                system_instruction=system_instruction,
                temperature=0.2,
                response_mime_type="application/json",
                response_schema=VillaResponse,
            ),
        )
        # Result parsen
        import json
        res_data = json.loads(response.text)
        return res_data.get("wissensluecke_erkannt", False), res_data.get("antwort_text", "")
    except Exception as e:
        return True, f"Fehler bei der KI-Kommunikation: {str(e)}"

# ==============================================================================
# 6. HMI PRESENTATION LAYER (BENUTZEROBERFLÄCHE)
# ==============================================================================
st.title("☀️ Villa Avatar")
st.subheader("Der digitale Begleiter für deinen Aufenthalt")

# 6.1 Rollenbasierte Weiche
rolle_auswahl = st.radio("Wähle deine Rolle:", ["Gast", "Host"], index=0, horizontal=True)
st.session_state.aktive_rolle = rolle_auswahl

# 6.2 Objekt-Dropdown (Single Source)
verfuegbare_objekte = [""]
if df_wissen is not None:
    # Filter für Relevanz Gast falls Rolle == Gast
    if st.session_state.aktive_rolle == "Gast":
        # Finde Spaltenname für Gast-Relevanz über Spalten-Lexikon
        g_spalte = "Relevanz Gast"
        for _, r in df_spalten.iterrows():
            if "gast" in str(r.get("sichtbar für gast", "")).lower() and "relevanz" in str(r.get("spaltenname in der wissensbasis", "")).lower():
                g_spalte = str(r.get("spaltenname in der wissensbasis", "")).strip()
        
        if g_spalte in df_wissen.columns:
            filtered_df = df_wissen[df_wissen[g_spalte].astype(str).str.lower().str.contains("x|ja", na=False)]
            verfuegbare_objekte += sorted(filtered_df.iloc[:, 0].dropna().astype(str).unique().tolist())
        else:
            verfuegbare_objekte += sorted(df_wissen.iloc[:, 0].dropna().astype(str).unique().tolist())
    else:
        verfuegbare_objekte += sorted(df_wissen.iloc[:, 0].dropna().astype(str).unique().tolist())

aktuelles_objekt = st.selectbox("Betroffenes Objekt / Ausstattung (Optional):", verfuegbare_objekte, index=0)

# ==============================================================================
# DYNAMISCHE BUTTONS AUS DEM USECASE_LEXIKON
# ==============================================================================
aktuelle_richtung = "OUTPUT" # Standard
placeholder_text = "Wie kann ich dir helfen?" # Standard-Fallback

if df_usecase is not None and not df_usecase.empty:
    st.write("**Wähle ein Anliegen:**")
    
    # Ermittle Spaltenname für die Sichtbarkeit anhand der Rolle
    sichtbarkeits_spalte = f"sichtbar für {st.session_state.aktive_rolle.lower()}"
    # Finde die echte Spalte im DataFrame via Sanitized Header
    real_vis_col = None
    for c in df_usecase.columns:
        if sichtbarkeits_spalte in c:
            real_vis_col = c
            break
            
    # Filtere Use Cases, die im HMI sichtbar sein sollen und für die Rolle freigegeben sind
    hmi_col = [c for c in df_usecase.columns if "im hmi" in c][0]
    label_col = [c for c in df_usecase.columns if "button_label" in c][0]
    prompt_col = [c for c in df_usecase.columns if "chat_prompt" in c][0]
    direction_col = [c for c in df_usecase.columns if "richtung" in c][0]
    uc_name_col = [c for c in df_usecase.columns if "use case" in c][0]
    
    if real_vis_col:
        valid_usecases = df_usecase[
            (df_usecase[real_vis_col].astype(str).str.lower().str.strip() == "ja") &
            (df_usecase[hmi_col].astype(str).str.lower().str.strip() == "ja")
        ]
        
        # Erzeuge Layout-Spalten für die Buttons nebeneinander
        cols = st.columns(len(valid_usecases))
        for idx, (_, uc_row) in enumerate(valid_usecases.iterrows()):
            btn_label = str(uc_row[label_col]).strip()
            system_uc_name = str(uc_row[uc_name_col]).strip()
            
            with cols[idx]:
                if st.button(btn_label, key=f"btn_{system_uc_name}", use_container_width=True):
                    st.session_state.aktiver_use_case = system_uc_name
                    st.rerun()

    # Logik für den aktiven Use Case laden (Richtung und Chat-Prompt ermitteln)
    if st.session_state.aktiver_use_case:
        match_row = df_usecase[df_usecase[uc_name_col].astype(str).str.lower().str.strip() == st.session_state.aktiver_use_case.lower()]
        if not match_row.empty:
            aktuelle_richtung = str(match_row.iloc[0][direction_col]).strip().upper()
            placeholder_text = str(match_row.iloc[0][prompt_col]).strip()

# Visuelle Bestätigung des aktiven Modus für den Nutzer
if st.session_state.aktiver_use_case:
    st.info(f"Aktivierter Modus: **{st.session_state.aktiver_use_case}**")

# ==============================================================================
# 7. CHAT LOGIK & ASYMMETRISCHES INTERFACE
# ==============================================================================
# Chat-Historie rendern
for msg in st.session_state.messages:
    if msg["role"] == "user":
        # Asymmetrisches Design: Nutzer rechtsbündig, grauer Hintergrund
        st.markdown(
            f'<div style="display: flex; justify-content: flex-end; margin-bottom: 10px;">'
            f'<div style="background-color: #f0f2f6; color: #31333F; padding: 10px 15px; '
            f'border-radius: 15px; max-width: 75%; text-align: left; box-shadow: 1px 1px 2px rgba(0,0,0,0.1);">'
            f'{msg["content"]}</div></div>', 
            unsafe_allow_html=True
        )
    else:
        # KI-Nachricht linksbündig (Klassischer Streamlit Chat-Stil)
        with st.chat_message("assistant", avatar="☀️"):
            st.write(msg["content"])

# Dynamischer Chat-Input mit dem geladenen Placeholder aus Excel
if user_input := st.chat_input(placeholder_text):
    # Nachricht sofort visuell hinzufügen
    st.session_state.messages.append({"role": "user", "content": user_input})
    
    # Kontext aus Excel-Matrix extrahieren
    context_str = ""
    if df_wissen is not None and aktuelles_objekt:
        # Finde Zeile des Objekts
        row_data = df_wissen[df_wissen.iloc[:, 0].astype(str).str.lower().str.strip() == aktuelles_objekt.lower()]
        if not row_data.empty:
            # Nur für die Rolle freigegebene Spalten extrahieren (Spalten_Lexikon)
            freigegebene_spalten = []
            role_col_lex = f"sichtbar für {st.session_state.aktive_rolle.lower()}"
            
            # Finde die echte Spalte im Spalten-Lexikon
            lex_vis_col = None
            for c in df_spalten.columns:
                if role_col_lex in c:
                    lex_vis_col = c
                    break
            
            wissen_col_name = [c for c in df_spalten.columns if "spaltenname in der wissensbasis" in c][0]
            
            if lex_vis_col:
                allowed_rows = df_spalten[df_spalten[lex_vis_col].astype(str).str.lower().str.strip() == "ja"]
                freigegebene_spalten = allowed_rows[wissen_col_name].dropna().astype(str).tolist()
            
            # Kontext formieren
            for col in df_wissen.columns:
                if any(clean_string(col) == clean_string(f_col) for f_col in freigegebene_spalten):
                    context_str += f"{col}: {row_data.iloc[0][col]}\n"
                    
    st.session_state.last_extracted_context = context_str
    
    # ROUTING-ENTSCHEIDUNG
    if aktuelle_richtung == "INPUT":
        # Direktes Schreiben in Matrix (z.B. bei Störungen/Feedback)
        execute_matrix_input(st.session_state.aktiver_use_case, aktuelles_objekt, user_input)
        ai_response = "Vielen Dank. Ich habe deine Nachricht erfolgreich in meiner Matrix registriert und an die Hosts weitergeleitet."
        st.session_state.messages.append({"role": "assistant", "content": ai_response})
    else:
        # Lese-Workflow (OUTPUT) mit KI-Auswertung
        wissensluecke, ai_response = ask_villa_avatar(user_input, context_str)
        
        if wissensluecke:
            # Kaskadierender Richtungswechsel (Transitional Routing)
            st.session_state.messages.append({"role": "assistant", "content": FALLBACK_SATZ})
            # Im Hintergrund in die Spalte für Wissenslücken ("Keine Information") schreiben
            execute_matrix_input("Keine Information", aktuelles_objekt, user_input)
        else:
            st.session_state.messages.append({"role": "assistant", "content": ai_response})
            
    st.rerun()

# ==============================================================================
# # --- DIAGNOSE-BLOCK ---
# ==============================================================================
st.write("")
with st.expander("🔍 SYSTEM-DIAGNOSE MONITOR (Laufzeit-Metriken)", expanded=True):
    d_col1, d_col2 = st.columns(2)
    with d_col1:
        st.metric(label="1. Aktive Rolle", value=str(st.session_state.get("aktive_rolle", "None")))
        st.metric(label="3. Gewähltes Objekt", value=str(aktuelles_objekt))
    with d_col2:
        st.metric(label="2. Use Case | Richtung", value=f"{st.session_state.get('aktiver_use_case')} | {aktuelle_richtung}")
        
    st.write("**4. Letzter Matrix-Lesestatus Spalte 'Details Nutzung':**")
    if df_wissen is not None and not df_wissen.empty:
        st.success(f"✅ Daten erfolgreich geladen ({len(df_wissen)} Objekte in Matrix verifiziert)")
    else:
        st.error("🛑 Lesefehler: Spalte 'Details Nutzung' nicht synchronisiert.")
        
    st.write("**5. Letzter Matrix-Schreibstatus:**")
    st.info(st.session_state.get("last_write_status", "Kein Status"))
    
    st.write("**6. Letzter Kontext-Extrakt (KI-Input):**")
    st.text_area(label="Matrix-Rohdaten", value=st.session_state.get("last_extracted_context", ""), height=100, disabled=True, label_visibility="collapsed")
