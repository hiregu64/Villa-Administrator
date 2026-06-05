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
# CONFIGURATION & DRIVE INTERFACE
# ==========================================
st.set_page_config(page_title="Villa Avatar", page_icon="☀️", layout="centered")
FILE_ID = '1FzhWZuO6aRZkdRuQBzaojhkq7bQDyprl'

# Exakter Wortlaut strikt gemäß deiner Master-PPT
FALLBACK_SATZ = "Ich habe dazu leider keine Informationen, Ich gebe das aber gern an die Hosts weiter."

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
        xl = pd.ExcelFile(fh)
        if "Wissensbasis" not in xl.sheet_names or "Spalten_Lexikon" not in xl.sheet_names:
            st.error("Fehler: Worksheets 'Wissensbasis' oder 'Spalten_Lexikon' fehlen.")
            return None, None, None
            
        df_wissen = pd.read_excel(fh, sheet_name="Wissensbasis")
        fh.seek(0)
        df_lexikon = pd.read_excel(fh, sheet_name="Spalten_Lexikon")
        
        if df_wissen is not None and not df_wissen.empty and "Wo?" in df_wissen.columns:
            df_wissen["Wo?"] = df_wissen["Wo?"].ffill()
            
        return df_wissen, df_lexikon, service
    except Exception as e:
        st.error(f"Fehler beim Laden der Excel-Matrix: {e}")
        return None, None, None

with st.spinner("Initialisiere dynamische Matrix..."):
    df_wissen, df_lexikon, drive_service = load_dynamic_data()

# ==========================================
# 1. CORE DATA ENGINE (Isolierte Schreib-Schnittstelle)
# ==========================================
def find_column_by_fuzzy_name(headers, target_name):
    cleaned_headers = [str(h).strip().lower().replace("\n", " ").replace("\r", " ") for h in headers]
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
            t_col = headers.index(zielspalte_name) + 1 if zielspalte_name in headers else (find_column_by_fuzzy_name(headers, "Details Nutzung") or 8)
        
        if t_col is None:
            return False

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

# ==========================================
# 2. KI CORE 
# ==========================================
VILLA_PROMPT = f"""Du bist „Villa Avatar“, digitaler Helfer für Gäste, Hosts und Admins. Antworte smartphone-optimiert. Nutze AUSSCHLIESSLICH die Tabellendaten. Erwähne niemals Tabellenstrukturen oder Spaltennamen."""

@st.cache_resource
def get_ki_client():
    if "GEMINI_API_KEY" in st.secrets: return genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
    return None
client = get_ki_client()

def ai_waehle_stammdaten_spalte(user_text, verfuegbare_spalten):
    if client is None or not verfuegbare_spalten: return "Details Nutzung [Output]"
    try:
        spalten_str = "\n".join([f"- {s}" for s in verfuegbare_spalten])
        prompt = f'Analysiere: "{user_text}"\nIn welche Spalte gehört das?\n{spalten_str}\nAntworte NUR mit dem Spaltennamen.'
        response = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        return response.text.strip().replace('"', '').replace("'", "")
    except: return "Details Nutzung [Output]"

def generate_ki_response(prompt_text):
    if client is None: return "🛑 KI nicht konfiguriert."
    try:
        return client.models.generate_content(model="gemini-2.5-flash", contents=prompt_text, config=types.GenerateContentConfig(system_instruction=VILLA_PROMPT)).text
    except:
        try: return client.models.generate_content(model="gemini-2.0-flash", contents=prompt_text, config=types.GenerateContentConfig(system_instruction=VILLA_PROMPT)).text
        except: return "🛑 KI temporär überlastet."

# ==========================================
# 3. HMI PRESENTATION LAYER
# ==========================================
st.markdown("""<style>
    div.stButton > button[kind="primary"] { background-color: #e3f2fd !important; color: #1565c0 !important; border: 1px solid #bbdefb !important; font-weight: bold !important; }
</style>""", unsafe_allow_html=True)

st.title("☀️ Villa Avatar")

STANDARD_DROPDOWNS = ["Ausstattung innen", "Ausstattung außen", "In der Nähe"]

HMI_MATRIX = {
    "Gast": {
        "Ich brauche Hilfe.": {"sys_action": "Hilfe", "reply": "Wobei kann ich dir helfen?", "dd": STANDARD_DROPDOWNS},
        "Ich möchte eine Störung melden.": {"sys_action": "Störung", "reply": "Was ist defekt? Ich kümmere mich darum.", "dd": STANDARD_DROPDOWNS},
        "Ich möchte Feedback geben.": {"sys_action": "Feedback", "reply": "Welches Feedback hast du für uns?", "dd": STANDARD_DROPDOWNS}
    },
    "Host": {
        "Ich brauche Hilfe.": {"sys_action": "Hilfe", "reply": "Wobei kann ich dir helfen?", "dd": STANDARD_DROPDOWNS},
        "Ich habe neue Informationen.": {"sys_action": "Information", "reply": "Gern ordne ich deine Notizen zu.", "dd": STANDARD_DROPDOWNS},
        "Ich benötige einen Bericht.": {"sys_action": "Bericht", "reply": "Nenne mir bitte den Zeitraum und das Thema.", "dd": []},
        "Ich möchte eine Störung melden.": {"sys_action": "Störung", "reply": "Was ist passiert? Ich halte es fest.", "dd": STANDARD_DROPDOWNS},
        "Ich möchte Feedback geben.": {"sys_action": "Feedback", "reply": "Welches Feedback dokumentieren wir?", "dd": STANDARD_DROPDOWNS}
    }
}

if "messages" not in st.session_state: st.session_state.messages = []
if "aktive_aktion" not in st.session_state: st.session_state.aktive_aktion = None
if "vorherige_rolle" not in st.session_state: st.session_state.vorherige_rolle = None

def handle_button_click(aktions_satz):
    for key in list(st.session_state.keys()):
        if key.startswith("sub_cat_wahl_"): del st.session_state[key]
    st.session_state.aktive_aktion = aktions_satz
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
    if nutzer_rolle == "Gast":
        col1, col2, col3 = st.columns(3)
        for col, satz in zip([col1, col2, col3], HMI_MATRIX["Gast"].keys()):
            with col:
                if st.button(satz, use_container_width=True, type="primary" if st.session_state.aktive_aktion == satz else "secondary"): handle_button_click(satz)
    elif nutzer_rolle == "Host":
        col1, col2, col3 = st.columns(3)
        col4, col5 = st.columns(2)
        for col, satz in zip([col1, col2, col3, col4, col5], HMI_MATRIX["Host"].keys()):
            with col:
                if st.button(satz, use_container_width=True, type="primary" if st.session_state.aktive_aktion == satz else "secondary"): handle_button_click(satz)

    if st.session_state.aktive_aktion and nutzer_rolle in HMI_MATRIX:
        cfg = HMI_MATRIX[nutzer_rolle].get(st.session_state.aktive_aktion)
        if cfg:
            with st.chat_message("assistant"): st.markdown(cfg['reply'])
            if df_wissen is not None and not df_wissen.empty:
                bez_spalte = "Bezeichnung" if "Bezeichnung" in df_wissen.columns else df_wissen.columns[0]
                kat_spalte = "Wo?" if "Wo?" in df_wissen.columns else df_wissen.columns[1]
                for kat in cfg["dd"]:
                    if "innen" in kat.lower(): mask = df_wissen[kat_spalte].astype(str).str.contains("innen", case=False, na=False)
                    elif "außen" in kat.lower() or "aussen" in kat.lower(): mask = df_wissen[kat_spalte].astype(str).str.contains("außen|aussen", case=False, na=False)
                    else: mask = df_wissen[kat_spalte].astype(str).str.contains("nähe|naehe|In der Nähe", case=False, na=False)
                    
                    if nutzer_rolle == "Gast" and "Relevanz Gast" in df_wissen.columns:
                        mask = mask & (df_wissen["Relevanz Gast"].astype(str).str.strip().str.lower() == "x")
                    
                    verfuegbare_bez = sorted([str(b).strip() for b in df_wissen[mask][bez_spalte].dropna().drop_duplicates().tolist()])
                    if "Nicht gefunden" in verfuegbare_bez: verfuegbare_bez.remove("Nicht gefunden")
                    verfuegbare_bez.append("Nicht gefunden")
                    st.selectbox(label=f"Hidden_{kat}", options=verfuegbare_bez, index=None, placeholder=f"📍 {kat} wählen...", key=f"sub_cat_wahl_{kat}_{st.session_state.aktive_aktion}", label_visibility="collapsed")

st.write("---")
for message in st.session_state.messages:
    with st.chat_message(message["role"]): st.markdown(message["content"])

# ==========================================
# 4. LOGICAL ARCHITECTURE CONTROLLER
# ==========================================
if prompt := st.chat_input("Bitte schreibe hier..."):
    if nutzer_rolle is None or not st.session_state.aktive_aktion:
        st.warning("Bitte wähle Rolle und Anliegen aus!")
    else:
        with st.chat_message("user"): st.markdown(prompt)
        st.session_state.messages.append({"role": "user", "content": prompt})
        
        aktuelle_sys_action = HMI_MATRIX[nutzer_rolle][st.session_state.aktive_aktion]["sys_action"]
        
        konkrete_auswahlen = {}
        for kat in HMI_MATRIX[nutzer_rolle][st.session_state.aktive_aktion]["dd"]:
            key = f"sub_cat_wahl_{kat}_{st.session_state.aktive_aktion}"
            if key in st.session_state and st.session_state[key] is not None:
                konkrete_auswahlen[kat] = st.session_state[key]
        
        gewaehlte_objekte_str = ", ".join([f"{k}: {v}" for k, v in konkrete_auswahlen.items()]) if konkrete_auswahlen else "Keines"
        gewaehltes_objekt = list(konkrete_auswahlen.values())[0] if konkrete_auswahlen else None
        
        # 1. Standard-Einträge protokollieren
        if drive_service is not None and aktuelle_sys_action in ["Störung", "Feedback", "Information"]:
            with st.spinner("Protokolliere Eintrag in Excel..."):
                dynamic_write_to_excel(drive_service, prompt, nutzer_rolle, gewaehltes_objekt, aktuelle_sys_action, df_lexikon)

        # 2. DETERMINISTISCHE PRÜFUNG AUF WISSENSLÜCKEN (Kein Umweg über KI-Text-Parsing!)
        wissensluecke_erkannt = False
        kontext = ""
        
        if df_wissen is not None and not df_wissen.empty:
            df_gefiltert = df_wissen.copy()
            bez_spalte = "Bezeichnung" if "Bezeichnung" in df_gefiltert.columns else df_gefiltert.columns[0]
            kat_spalte = "Wo?" if "Wo?" in df_gefiltert.columns else df_gefiltert.columns[1]
            
            if gewaehltes_objekt and gewaehltes_objekt != "Nicht gefunden":
                df_gefiltert = df_gefiltert[df_gefiltert[bez_spalte].astype(str).str.strip().str.lower() == str(gewaehltes_objekt).strip().lower()]
            elif konkrete_auswahlen:
                mask = pd.Series(False, index=df_gefiltert.index)
                for kat in konkrete_auswahlen.keys():
                    if "innen" in kat.lower(): mask = mask | df_gefiltert[kat_spalte].astype(str).str.contains("innen", case=False, na=False)
                    elif "außen" in kat.lower() or "aussen" in kat.lower(): mask = mask | df_gefiltert[kat_spalte].astype(str).str.contains("außen|aussen", case=False, na=False)
                    else: mask = mask | df_gefiltert[kat_spalte].astype(str).str.contains("nähe|naehe|In der Nähe", case=False, na=False)
                df_gefiltert = df_gefiltert[mask]
                
            if nutzer_rolle == "Gast":
                if "Relevanz Gast" in df_gefiltert.columns:
                    df_gefiltert = df_gefiltert[df_gefiltert["Relevanz Gast"].astype(str).str.strip().str.lower() == "x"]
                if df_lexikon is not None and "Sichtbar für Gast" in df_lexikon.columns:
                    erlaubt = df_lexikon[df_lexikon["Sichtbar für Gast"].astype(str).str.strip().str.lower() == "ja"]["Spaltenname"].tolist()
                    df_gefiltert = df_gefiltert[[c for c in df_gefiltert.columns if str(c).strip() in erlaubt]]

            # Wenn keine Zeilen existieren ODER alle relevanten Textfelder komplett leer (NaN) sind -> Sofortige Lücke!
            if df_gefiltert.empty or df_gefiltert.dropna(how='all').empty:
                wissensluecke_erkannt = True
            else:
                lexikon_text = ""
                if df_lexikon is not None and not df_lexikon.empty:
                    for _, row in df_lexikon.iterrows():
                        spaltenname = str(row.get("Spaltenname", ""))
                        if spaltenname in df_gefiltert.columns:
                            lexikon_text += f"- '{spaltenname}': {row.get('Bedeutung / Beschreibung', '')}\n"
                kontext = f"\n\n{lexikon_text}\nAktuelle Daten:\n{df_gefiltert.to_string(index=False)}"

        # 3. AUSGABE-STEUERUNG ANHAND DER ERKENNTNIS
        with st.chat_message("assistant"):
            if wissensluecke_erkannt:
                # Erkenntnis liegt direkt vor -> System-Aktion & PPT Fallback sofort abfeuern!
                st.markdown(FALLBACK_SATZ)
                st.session_state.messages.append({"role": "assistant", "content": FALLBACK_SATZ})
                
                if drive_service is not None:
                    with st.spinner("Wissenslücke wird direkt dokumentiert..."):
                        if dynamic_write_to_excel(drive_service, prompt, nutzer_rolle, gewaehltes_objekt, "Keine Information", df_lexikon):
                            st.cache_data.clear()
            else:
                # Daten vorhanden -> KI formatiert die Antwort smartphone-optimiert
                with st.spinner("Villa Avatar überlegt..."):
                    antwort_text = generate_ki_response(f"Rolle: {nutzer_rolle}\nSystem-Aktion: {aktuelle_sys_action}\nObjekt: {gewaehlte_objekte_str}\nAnfrage: {prompt} {kontext}")
                st.markdown(antwort_text)
                st.session_state.messages.append({"role": "assistant", "content": antwort_text})
