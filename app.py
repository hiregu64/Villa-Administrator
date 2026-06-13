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
    wissensluecke_erkannt: bool = Field(description="True bei unvollständigem Excel-Kontext.")
    antwort_text: str = Field(description="Antworttext für Gast. Leerstring bei Wissenslücke.")

# ==============================================================================
# 2. GLOBAL CONFIGURATION & STYLING
# ==============================================================================
st.set_page_config(page_title="Villa Avatar", page_icon="☀️", layout="centered")
FILE_ID = '1FzhWZuO6aRZkdRuQBzaojhkq7bQDyprl'
FALLBACK_SATZ = "Ich habe dazu leider keine Informationen, Ich gebe das aber gern an die Hosts weiter."

st.markdown("""
    <style>
    div.stButton > button[kind="primary"] { background-color: #e3f2fd !important; color: #1565c0 !important; border: 1px solid #bbdefb !important; font-weight: bold !important; }
    div[data-testid="stChatMessage"]:has(div[aria-label="Chat message from user"]) { flex-direction: row-reverse !important; background-color: #F0F2F6 !important; border-radius: 10px !important; padding: 10px !important; }
    div[data-testid="stChatMessage"]:has(div[aria-label="Chat message from user"]) div[data-testid="stChatMessageContent"] { text-align: right !important; width: 100% !important; }
    </style>
""", unsafe_allow_html=True)

# Session State flach initialisieren
for key, value in [
    ("aktive_rolle", None), ("aktiver_use_case", None), ("selected_object", None), 
    ("selected_field", None), ("messages", []), ("host_authentifiziert", False), 
    ("debug_modus_aktiv", False), ("last_write_status", "Kein Schreibvorgang."), ("last_extracted_context", "")
]:
    if key not in st.session_state: st.session_state[key] = value

# ==============================================================================
# 3. DATEN-LADE ENGINE & UTILS
# ==============================================================================
@st.cache_data(ttl=30)
def load_dynamic_data():
    try:
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
        if df_wissen is not None and "Wo?" in df_wissen.columns: df_wissen["Wo?"] = df_wissen["Wo?"].ffill()
        return df_wissen, df_lexikon, df_usecases, df_passwoerter, service
    except Exception as e:
        st.error(f"Fehler bei Synchronisation: {e}")
        return None, None, None, None, None

with st.spinner("Synchronisiere Matrix..."):
    df_wissen, df_lexikon, df_usecases, df_passwoerter, drive_service = load_dynamic_data()

for df in [df_usecases, df_passwoerter, df_wissen, df_lexikon]:
    if df is not None: df.columns = [str(c).strip() for c in df.columns]

def find_column_by_fuzzy_name(headers, target_name):
    cleaned = [str(h).strip().lower() for h in headers]
    search = str(target_name).strip().lower()
    return cleaned.index(search) + 1 if search in cleaned else next((i + 1 for i, h in enumerate(cleaned) if search in h), None)

# ==============================================================================
# 4. API & CONTEXT ROUTING ENGINE
# ==============================================================================
def call_gemini(prompt, context="", structured=True):
    client = genai.Client(api_key=st.secrets.get("GEMINI_API_KEY")) if "GEMINI_API_KEY" in st.secrets else None
    if not client: return "KI nicht konfiguriert." if not structured else KiAntwortSchema(wissensluecke_erkannt=True, antwort_text="")
    
    sys_instruction = "Du bist „Villa Avatar“. Antworte kurz, präzise, smartphone-optimiert. Nutze den Excel-Kontext intelligent. Erwähne niemals Tabellenstrukturen."
    try:
        if structured:
            res = client.models.generate_content(model="gemini-2.5-flash", contents=f"Kontext:\n{context}\n\nFrage: {prompt}", config=types.GenerateContentConfig(system_instruction=sys_instruction, temperature=0.2, response_mime_type="application/json", response_schema=KiAntwortSchema))
            data = json.loads(res.text)
            return KiAntwortSchema(wissensluecke_erkannt=bool(data.get("wissensluecke_erkannt", True)), antwort_text=str(data.get("antwort_text", "")))
        return client.models.generate_content(model="gemini-2.5-flash", contents=prompt, config=types.GenerateContentConfig(system_instruction="Verfasse einen sachlichen Bericht basierend auf den Daten.", temperature=0.2)).text
    except:
        return KiAntwortSchema(wissensluecke_erkannt=True, antwort_text="") if structured else "Fehler."

def extract_context_for_object(objekt_name):
    if df_wissen is None or df_lexikon is None or not objekt_name: return ""
    bez_col = "Bezeichnung" if "Bezeichnung" in df_wissen.columns else df_wissen.columns[0]
    row = df_wissen[df_wissen[bez_col].astype(str).str.strip().str.lower() == objekt_name.lower().strip()]
    if row.empty: return ""
    
    context_parts = [f"Objekt: {objekt_name}"]
    if str(st.session_state.aktive_rolle).lower() == "host":
        for col in df_wissen.columns:
            if col != bez_col and "status" not in col.lower() and col.lower() != "wo?":
                if pd.notna(row.iloc[0][col]): context_parts.append(f"- {col}: {row.iloc[0][col]}")
    else:
        gast_col = next((c for c in df_lexikon.columns if "gast" in c.lower()), df_lexikon.columns[-1])
        tags = df_lexikon[df_lexikon[gast_col].astype(str).str.lower().str.strip() == "ja"][df_lexikon.columns[0]].tolist()
        for col in df_wissen.columns:
            if any(col.lower() == t.lower() for t in tags) and col in row.columns and pd.notna(row.iloc[0][col]):
                context_parts.append(f"- {col}: {row.iloc[0][col]}")
                
    st.session_state.last_extracted_context = "\n".join(context_parts)
    return st.session_state.last_extracted_context

# ==============================================================================
# 5. MATRIZEN-SCHREIBENGINE
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
            old = ws.cell(row_idx, col_idx).value or ""
            ws.cell(row_idx, col_idx).value = f"{old}\n[{datetime.datetime.now().strftime('%d.%m.%Y %H:%M')} | {st.session_state.aktive_rolle}]: {text}".strip()
            ws.cell(row_idx, col_idx).alignment = Alignment(wrap_text=True)
            
            status_idx = find_column_by_fuzzy_name(headers, f"{physische_spalte} Status")
            if status_idx: ws.cell(row_idx, status_idx).value = "offen"
            
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            drive_service.files().update(fileId=FILE_ID, media_body=MediaIoBaseUpload(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')).execute()
            st.session_state.last_write_status = f"✅ Gespeichert in Zeile {row_idx}, Spalte {physische_spalte}"
            st.toast("✅ Matrix aktualisiert!")
    except Exception as e: st.error(f"Schreibfehler: {e}")

def execute_matrix_input(use_case, objekt, text):
    if df_lexikon is None: return
    spalte = next((str(r[df_lexikon.columns[0]]).strip() for _, r in df_lexikon.iterrows() if use_case.lower().strip() in [t.strip().lower() for t in str(r[df_lexikon.columns[4]]).split(',')]), None)
    if not spalte and use_case == "Keine Information": spalte = next((c for c in df_wissen.columns if "information" in c.lower() and "status" not in c.lower()), None)
    if spalte: execute_matrix_input_direct(spalte, objekt, text)

# ==============================================================================
# 6. HMI PRESENTATION LAYER
# ==============================================================================
st.title("☀️ Villa Avatar")

# Rollenauswahl
role = st.selectbox("Rolle", options=["Gast", "Host"], index=None, placeholder="Wer bist du?", label_visibility="collapsed")
if role and role != st.session_state.aktive_rolle:
    st.session_state.aktive_rolle = role
    st.session_state.aktiver_use_case, st.session_state.selected_object, st.session_state.selected_field, st.session_state.messages = None, None, None, []
    st.rerun()

if not st.session_state.aktive_rolle: st.stop()

# Passwort-Gate
if st.session_state.aktive_rolle == "Host" and not st.session_state.host_authentifiziert:
    pwd = st.text_input("🔑 Passwort eingeben:", type="password")
    if pwd and df_passwoerter is not None:
        if any(pwd.strip() == str(r[df_passwoerter.columns[1]]).strip() for _, r in df_passwoerter.iterrows()):
            st.session_state.host_authentifiziert = True
            st.session_state.debug_modus_aktiv = any(str(r[df_passwoerter.columns[2]]).strip().lower() == "debug" for _, r in df_passwoerter.iterrows() if pwd.strip() == str(r[df_passwoerter.columns[1]]).strip())
            st.rerun()
    st.stop()

# Use Cases Buttons
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
                st.rerun()

if not st.session_state.aktiver_use_case: st.stop()

# Use Case Metadaten extrahieren
uc_row = df_usecases[df_usecases[df_usecases.columns[0]].astype(str).str.lower().str.strip() == st.session_state.aktiver_use_case.lower().strip()]
richtung = str(uc_row.iloc[0][df_usecases.columns[1]]).strip().upper() if not uc_row.empty else "OUTPUT"
fragetext = str(uc_row.iloc[0][df_usecases.columns[4]]).strip() if not uc_row.empty and len(df_usecases.columns) > 4 and pd.notna(uc_row.iloc[0][df_usecases.columns[4]]) else "Wie kann ich helfen?"
danke_tmpl = str(uc_row.iloc[0][df_usecases.columns[5]]).strip() if not uc_row.empty and len(df_usecases.columns) > 5 and pd.notna(uc_row.iloc[0][df_usecases.columns[5]]) else "Danke!"

# ==============================================================================
# SPEZIAL-MODE: SYSTEM-BERICHTE (Keine Objektauswahl nötig)
# ==============================================================================
if "bericht" in st.session_state.aktiver_use_case.lower():
    typ = st.selectbox("Typ", ["Offene Störungen", "Behobene Störungen", "Offenes Feedback", "Offene Wissenslücken", "Gesamtübersicht"], index=None, placeholder="Berichtsart wählen...")
    if typ and st.button("📊 Bericht generieren", type="primary", use_container_width=True):
        lines = []
        for c in [col for col in df_wissen.columns if any(x in col.lower() for x in ["störung", "feedback", "information"]) and "status" not in col.lower()]:
            for _, r in df_wissen.iterrows():
                if pd.notna(r[c]) and str(r[c]).strip(): lines.append(f"Objekt: {r[df_wissen.columns[0]]} | Feld: {c}\nEintrag: {r[c]}\n---")
        st.markdown(call_gemini(f"Strukturiere das chronologisch:\n\n" + "\n".join(lines), structured=False) if lines else "Keine Einträge.")
    st.stop()

# ==============================================================================
# PROGRESSIVES HMI: SCHRITT 1 – DIE 3 ENTKOPPELTEN DROP-DOWNS
# ==============================================================================
bez_col, kat_col = df_wissen.columns[0], ("Wo?" if "Wo?" in df_wissen.columns else df_wissen.columns[1])

def get_liste(pattern):
    mask = df_wissen[kat_col].astype(str).str.contains(pattern, case=False, na=False)
    if st.session_state.aktive_rolle == "Gast" and "Relevanz Gast" in df_wissen.columns:
        mask = mask & (df_wissen["Relevanz Gast"].astype(str).str.strip().str.lower() == "x")
    return sorted(df_wissen[mask][bez_col].dropna().drop_duplicates().astype(str).str.strip().tolist())

# Callbacks regeln den exklusiven State-Reset im Hintergrund
def cb_innen():
    if st.session_state.sel_innen:
        st.session_state.selected_object = st.session_state.sel_innen
        st.session_state.sel_aussen = st.session_state.sel_naehe = None
        st.session_state.selected_field = None
        st.session_state.messages = []

def cb_aussen():
    if st.session_state.sel_aussen:
        st.session_state.selected_object = st.session_state.sel_aussen
        st.session_state.sel_innen = st.session_state.sel_naehe = None
        st.session_state.selected_field = None
        st.session_state.messages = []

def cb_naehe():
    if st.session_state.sel_naehe:
        st.session_state.selected_object = st.session_state.sel_naehe
        st.session_state.sel_innen = st.session_state.sel_aussen = None
        st.session_state.selected_field = None
        st.session_state.messages = []

# Widget-Keys initialisieren
if "sel_innen" not in st.session_state: st.session_state.sel_innen = None
if "sel_aussen" not in st.session_state: st.session_state.sel_aussen = None
if "sel_naehe" not in st.session_state: st.session_state.sel_naehe = None

# Rendern der 3 Felder laut Spezifikation
st.selectbox("🏠 Ausstattung innen:", options=[None] + get_liste("innen"), key="sel_innen", on_change=cb_innen)
st.selectbox("🌳 Ausstattung außen:", options=[None] + get_liste("außen|aussen"), key="sel_aussen", on_change=cb_aussen)
st.selectbox("📍 In der Nähe:", options=[None] + get_liste("nähe|naehe"), key="sel_naehe", on_change=cb_naehe)

if st.checkbox("Ich kann das gewünschte Objekt nicht finden.", key="not_found_check"):
    if st.session_state.selected_object != "Nicht gefunden":
        st.session_state.selected_object = "Nicht gefunden"
        st.session_state.sel_innen = st.session_state.sel_aussen = st.session_state.sel_naehe = None
        st.session_state.selected_field = None
        st.session_state.messages = []
        st.rerun()

# Harter Stop, falls noch kein Objekt bestimmt wurde
if not st.session_state.selected_object: st.stop()

# Status-Banner für visuelles Feedback
st.markdown(f"<div style='background-color:#e8f5e9; padding:10px; border-radius:5px; color:#2e7d32; font-weight:bold; margin-top:10px; margin-bottom:15px;'>Aktives Objekt: {st.session_state.selected_object}</div>", unsafe_allow_html=True)

# ==============================================================================
# PROGRESSIVES HMI: SCHRITT 2 – DIE DYNAMISCH EINGEBLENDETE DETAIL-EBENE
# ==============================================================================

# FALL A: "Neue Information" verlangt zwingend ein zweites Dropdown (Spaltenauswahl)
if st.session_state.aktiver_use_case == "Neue Information" and st.session_state.aktive_rolle == "Host":
    options_spalten = [c for c in df_wissen.columns if c.lower() not in ["bezeichnung", "wo?", "id", "kategorie", "relevanz gast"] and "status" not in c.lower()]
    st.selectbox("📄 Bitte wähle die Art der Information aus (Zielspalte):", options=[None] + options_spalten, key="selected_field")
    
    if not st.session_state.selected_field: st.stop()
    
    txt = st.text_area(f"Gib hier die neue Information für '{st.session_state.selected_object}' ein:")
    if st.button("💾 In Excel-Zentralmatrix speichern", type="primary") and txt.strip():
        execute_matrix_input_direct(st.session_state.selected_field, st.session_state.selected_object, txt.strip())
        st.success("Erfolgreich in Matrix dokumentiert!")
        st.cache_data.clear()
    st.stop()

# FALL B: Chat-Interface für Hilfe, Störung und Feedback
else:
    if st.session_state.debug_modus_aktiv:
        with st.expander("🔍 SYSTEM-DIAGNOSE", expanded=False):
            st.write(f"**UseCase:** {st.session_state.aktiver_use_case} | **Richtung:** {richtung}")
            st.write(f"**Schreibstatus:** {st.session_state.last_write_status}")
            st.text_area("Extrahierter Kontext", st.session_state.last_extracted_context, height=100)

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]): st.markdown(msg["content"])
        
    if user_input := st.chat_input(fragetext):
        st.session_state.messages.append({"role": "user", "content": user_input})
        st.rerun()
        
    if st.session_state.messages and st.session_state.messages[-1]["role"] == "user":
        u_text = st.session_state.messages[-1]["content"]
        
        if richtung == "OUTPUT":
            if st.session_state.selected_object == "Nicht gefunden":
                st.session_state.messages.append({"role": "assistant", "content": FALLBACK_SATZ})
                execute_matrix_input("Keine Information", "Nicht gefunden", u_text)
            else:
                with st.spinner("Prüfe Daten..."):
                    res = call_gemini(u_text, extract_context_for_object(st.session_state.selected_object))
                    if res.wissensluecke_erkannt or not res.antwort_text or any(p in res.antwort_text.lower() for p in ["keine information", "weiß ich nicht", "leider nein"]):
                        st.session_state.messages.append({"role": "assistant", "content": FALLBACK_SATZ})
                        execute_matrix_input("Keine Information", st.session_state.selected_object, u_text)
                    else:
                        st.session_state.messages.append({"role": "assistant", "content": res.antwort_text})
            st.cache_data.clear()
            st.rerun()
            
        elif richtung == "INPUT":
            with st.spinner("Protokolliere in Matrix..."):
                execute_matrix_input(st.session_state.aktiver_use_case, st.session_state.selected_object, u_text)
                st.session_state.messages.append({"role": "assistant", "content": danke_tmpl.replace("{use_case}", st.session_state.aktiver_use_case)})
                st.cache_data.clear()
                st.rerun()
