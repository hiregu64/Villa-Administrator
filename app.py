import streamlit as st
from google import genai
from google.genai import types
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import pandas as pd
import io
import datetime
import json

# Google File ID der Excel-Tabelle
FILE_ID = '1FzhWZuO6aRZkdRuQBzaojhkq7bQDyprl'

# ==========================================
# 1. LIVE-DATEN AUS GOOGLE DRIVE LESEN
# ==========================================
@st.cache_data(ttl=30)  
def load_data_from_drive():
    try:
        creds_dict = st.secrets["GOOGLE_CREDENTIALS"]
        creds = service_account.Credentials.from_service_account_info(creds_dict)
        service = build('drive', 'v3', credentials=creds)
        
        request = service.files().get_media(fileId=FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            
        fh.seek(0)
        df = pd.read_excel(fh)
        
        # Leere Zellen in Spalte A (Kategorie) automatisch auffüllen für KI-Kontext
        if df is not None and not df.empty:
            df.iloc[:, 0] = df.iloc[:, 0].ffill()
            
        return df, service
    except Exception as e:
        st.error(f"Fehler bei der Verbindung zur Google Drive Wissensbasis: {e}")
        return None, None

with st.spinner("Verbindung zur Google Drive Wissensbasis wird hergestellt..."):
    df_wissen, drive_service = load_data_from_drive()

# ==========================================
# 2. INTELLIGENTES WRITE-BACK IN DIE ZEILEN (XLS)
# ==========================================
def save_data_to_drive(df, service, aktions_typ, text_inhalt, gewaehltes_objekt=None, rolle="Gast"):
    try:
        df_aktualisiert = df.copy()
        kat_col = df_aktualisiert.columns[0]
        bez_col = df_aktualisiert.columns[1] if len(df_aktualisiert.columns) > 1 else df_aktualisiert.columns[0]
        
        # Spaltennamen bereinigen für exaktes Matching
        col_mapping = {str(c).strip(): c for c in df_aktualisiert.columns}
        
        # Zielspalten identifizieren
        col_stoerfall = col_mapping.get("Störfall", "Störfall")
        col_stoerung_status = col_mapping.get("Störung Status [aktiv, OK]", "Störung Status [aktiv, OK]")
        col_feedback = col_mapping.get("Feedback", "Feedback")
        col_feedback_status = col_mapping.get("Feedback Status [offen, Nein, OK]", "Feedback Status [offen, Nein, OK]")
        
        # Sicherstellen, dass die Spalten im DataFrame existieren
        for col in [col_stoerfall, col_stoerung_status, col_feedback, col_feedback_status]:
            if col not in df_aktualisiert.columns:
                df_aktualisiert[col] = None
        
        # Zeilen-Index des ausgewählten Objekts (Spalte B) finden
        zeilen_index = None
        if gewaehltes_objekt:
            matches = df_aktualisiert[df_aktualisiert[bez_col].astype(str).str.strip() == str(gewaehltes_objekt).strip()]
            if not matches.empty:
                zeilen_index = matches.index[0]
        
        timestamp = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
        
        if aktions_typ == "Störung":
            if zeilen_index is not None:
                df_aktualisiert.at[zeilen_index, col_stoerfall] = text_inhalt
                df_aktualisiert.at[zeilen_index, col_stoerung_status] = "aktiv"
            else:
                # Fallback: Neue Zeile anhängen, falls kein Objekt gewählt/gefunden wurde
                neue_zeile = {kat_col: "Allgemein", bez_col: gewaehltes_objekt or "Allgemein", col_stoerfall: text_inhalt, col_stoerung_status: "aktiv"}
                df_aktualisiert = pd.concat([df_aktualisiert, pd.DataFrame([neue_zeile])], ignore_index=True)
                
        elif aktions_typ == "Feedback":
            if zeilen_index is not None:
                df_aktualisiert.at[zeilen_index, col_feedback] = text_inhalt
                df_aktualisiert.at[zeilen_index, col_feedback_status] = "offen"
            else:
                # Fallback
                neue_zeile = {kat_col: "Allgemein", bez_col: gewaehltes_objekt or "Allgemein", col_feedback: text_inhalt, col_feedback_status: "offen"}
                df_aktualisiert = pd.concat([df_aktualisiert, pd.DataFrame([neue_zeile])], ignore_index=True)
                
        else:
            # Für Information, Bericht oder Änderung: Logeintrag anhängen
            col_update = col_mapping.get("Eintrag / Update", "Eintrag / Update")
            col_nutzer = col_mapping.get("Nutzer", "Nutzer")
            col_zeit = col_mapping.get("Zeitstempel", "Zeitstempel")
            
            neue_zeile = {
                col_zeit if col_zeit in col_mapping else "Zeitstempel": timestamp,
                col_nutzer if col_nutzer in col_mapping else "Nutzer": rolle,
                kat_col: "Allgemein",
                bez_col: gewaehltes_objekt or "Allgemein",
                col_update if col_update in col_mapping else "Eintrag / Update": text_inhalt
            }
            df_aktualisiert = pd.concat([df_aktualisiert, pd.DataFrame([neue_zeile])], ignore_index=True)

        # Zurück nach Google Drive schreiben
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_aktualisiert.to_excel(writer, index=False)
        output.seek(0)
        
        media = MediaIoBaseUpload(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        service.files().update(fileId=FILE_ID, media_body=media).execute()
        st.cache_data.clear()  
        return True
    except Exception as e:
        st.error(f"Fehler beim Schreiben in Google Drive: {e}")
        return False

# ==========================================
# 3. KI-GEHIRN (GEPRÜFTER SYSTEM-PROMPT)
# ==========================================
VILLA_PROMPT = """
Du bist „Villa Avatar“, der digitale Helfer für die Bewohner, Eigentümer und Admins der Villa. Deine Aufgabe ist es, den Betrieb und Erhalt des Hauses so einfach wie möglich zu halten.

DEINE FUNKTIONSWEISE:
1. DATEN-INTERPRETATION: 
   - Greife für alle Antworten exklusiv auf die bereitgestellte Datenbank zu.
   - SPALTENFARBE ALS ROLLE: Spaltenköpfe in "Orange" enthalten Host/ Admin-exklusive Informationen. Wenn der Nutzer nicht als Host oder Admin authentifiziert ist, dürfen diese Informationen NICHT ausgegeben werden.

2. PROZESS-LOGIK FÜR STÖRUNGEN und FEEDBACK:
   - Wenn ein Nutzer Feedback gibt: Schreibe dieses Feedback mit dem Status "offen" direkt in die entsprechende Zeile der Datenbank.
   - Bestätige dem Nutzer kurz und freundlich den Eingang („Ich habe dein Feedback notiert, danke!“).
   - Wenn ein Nutzer eine Störung meldet: Schreibe diese Störung mit dem Status "aktiv" direkt in die entsprechende Zeile der Datenbank.
   - Bestätige dem Nutzer kurz und freundlich den Eingang („Ich habe deine Störungsmeldung notiert, danke!“).
   - Vermeide bei der Bestätigung jeweils technische Details zur Datenbank.

3. KOMMUNIKATIONS-REGELN:
   - Antworte immer kurz, präzise und Smartphone-optimiert.
   - Nutze die vom HMI übergebene Rolle (Gast, Host oder Admin) als Arbeitsgrundlage.
   - WICHTIG: Erwähne gegenüber Gast und Host keine interne Dateinamen, Bildbezeichnungen (wie '.jfif' oder '.jpg'), URLs oder die tatsächliche Tabellenstruktur. Antworte so, als hättest du dieses Wissen einfach im Kopf.
"""

@st.cache_resource
def get_ki_client():
    if "GEMINI_API_KEY" in st.secrets:
        return genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
    return None

client = get_ki_client()

def generate_ki_response(prompt_text):
    if client is None:
        return "KI-Dienst nicht konfiguriert (API Key fehlt in den Secrets)."
    
    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt_text,
            config=types.GenerateContentConfig(system_instruction=VILLA_PROMPT)
        )
        return response.text
    except Exception as e:
        error_msg = str(e)
        if "429" in error_msg or "RESOURCE_EXHAUSTED" in error_msg or "503" in error_msg or "UNAVAILABLE" in error_msg:
            try:
                response = client.models.generate_content(
                    model="gemini-1.5-flash",
                    contents=prompt_text,
                    config=types.GenerateContentConfig(system_instruction=VILLA_PROMPT)
                )
                return response.text
            except Exception as e_fallback:
                if "429" in str(e_fallback) or "RESOURCE_EXHAUSTED" in str(e_fallback):
                    return "🛑 Das tägliche kostenlose Abfrage-Limit der Villa Avatar ist leider für heute aufgebraucht. Bitte versuche es morgen wieder!"
                return "⏳ Die KI-Server sind aktuell stark ausgelastet. Bitte warte einen kurzen Moment und sende deine Nachricht noch einmal."
        
        return f"Fehler bei der KI-Verarbeitung: {e}"

# ==========================================
# 4. BENUTZEROBERFLÄCHE (HMI) & STYLING
# ==========================================
st.set_page_config(page_title="Villa Avatar", page_icon="☀️", layout="centered")

st.markdown("""
    <style>
    /* Zartes Blau für aktive Buttons */
    div.stButton > button[kind="primary"] {
        background-color: #e3f2fd !important;
        color: #1565c0 !important;
        border: 1px solid #bbdefb !important;
        font-weight: bold !important;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #bbdefb !important;
        border: 1px solid #64b5f6 !important;
    }
    div[data-testid="stSelectbox"] div[data-baseweb="select"] { font-weight: bold; font-size: 15px; }
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
st.markdown("Hallo! Ich bin Villa Avatar, dein digitaler **'Helfer'**! Wähle deine Rolle aus, um zu beginnen.")

# HMI Einleitungs-Fragen Matrix laut Spezifikation (PPT)
HMI_TEXTE = {
    "Hilfe": "Wobei kann ich dir helfen?",
    "Störung": "Was ist passiert?",
    "Bericht": "Nenne mir bitte den Zeitraum und das Thema.",
    "Information": "Gern nehme ich deine Informationen auf und ordne sie in meiner Wissensbasis zu.",
    "Änderung": "Beschreibe deine Änderung so genau wie möglich.",
    "Feedback": "Welches Feedback hast du?"
}

if "messages" not in st.session_state:
    st.session_state.messages = []
if "aktive_aktion" not in st.session_state:
    st.session_state.aktive_aktion = None
if "vorherige_rolle" not in st.session_state:
    st.session_state.vorherige_rolle = None

def handle_button_click(aktions_name):
    st.session_state.aktive_aktion = aktions_name
    st.session_state.messages = []  
    st.rerun()

# Rollenauswahl (Gast, Host, Admin konform zur PPT)
nutzer_rolle = st.selectbox(
    label="Hidden_Rollen_Label",
    options=["Gast", "Host", "Admin"],
    index=None,
    placeholder="Wer bist du?",
    label_visibility="collapsed"
)

if nutzer_rolle != st.session_state.vorherige_rolle:
    st.session_state.vorherige_rolle = nutzer_rolle
    st.session_state.aktive_aktion = None
    st.session_state.messages = []  
    st.rerun()

if nutzer_rolle is not None:
    st.write("---")
    
    with st.container():
        st.markdown(
            "<div style='display: flex; justify-content: flex-end; align-items: center; gap: 8px; margin-bottom: 10px;'> "
            "<span style='font-weight: bold; font-size: 1.2rem; font-family: inherit;'>Mein Anliegen:</span>"
            "<div style='width: 32px; height: 32px; background-color: rgb(255, 75, 75); border-radius: 8px; display: flex; align-items: center; justify-content: center;'>"
            "<svg viewBox='0 0 24 24' width='20' height='20' stroke='white' stroke-width='2' fill='none' stroke-linecap='round' stroke-linejoin='round'>"
            "<circle cx='12' cy='12' r='10'></circle>"
            "<path d='M8 14s1.5 2 4 2 4-2 4-2'></path>"
            "<line x1='9' y1='9' x2='9.01' y2='9'></line>"
            "<line x1='15' y1='9' x2='15.01' y2='9'></line>"
            "</svg>"
            "</div>"
            "</div>", 
            unsafe_allow_html=True
        )
    
    # Dynamisches Rendering der Buttons passend zur gewählten Rolle
    if nutzer_rolle == "Gast":
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("Ich brauche Hilfe.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Hilfe" else "secondary"):
                handle_button_click("Hilfe")
        with col2:
            if st.button("Es gibt eine Störung.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Störung" else "secondary"):
                handle_button_click("Störung")
        with col3:
            if st.button("Ich möchte Feedback geben.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Feedback" else "secondary"):
                handle_button_click("Feedback")
                
    elif nutzer_rolle == "Host":
        col1, col2, col3 = st.columns(3)
        col4, col5 = st.columns(2)
        with col1:
            if st.button("Ich brauche Hilfe.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Hilfe" else "secondary"):
                handle_button_click("Hilfe")
        with col2:
            if st.button("Ich habe neue Informationen.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Information" else "secondary"):
                handle_button_click("Information")
        with col3:
            if st.button("Es gibt eine Störung.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Störung" else "secondary"):
                handle_button_click("Störung")
        with col4:
            if st.button("Ich benötige einen Bericht.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Bericht" else "secondary"):
                handle_button_click("Bericht")
        with col5:
            if st.button("Ich möchte Feedback geben.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Feedback" else "secondary"):
                handle_button_click("Feedback")
                
    elif nutzer_rolle == "Admin":
        col1, col2, col3 = st.columns(3)
        col4, col5, col6 = st.columns(3)
        with col1:
            if st.button("Ich brauche Hilfe.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Hilfe" else "secondary"):
                handle_button_click("Hilfe")
        with col2:
            if st.button("Ich habe neue Informationen.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Information" else "secondary"):
                handle_button_click("Information")
        with col3:
            if st.button("Es gibt eine Störung.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Störung" else "secondary"):
                handle_button_click("Störung")
        with col4:
            if st.button("Ich benötige einen Bericht.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Bericht" else "secondary"):
                handle_button_click("Bericht")
        with col5:
            if st.button("Ich möchte eine Änderung an der Wissensbasis vornehmen.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Änderung" else "secondary"):
                handle_button_click("Änderung")
        with col6:
            if st.button("Ich möchte Feedback geben.", use_container_width=True, type="primary" if st.session_state.aktive_aktion == "Feedback" else "secondary"):
                handle_button_click("Feedback")

    # Einmaliges, klares Such- und Auswahlelement für Bezeichnungen aus Spalte B
    gewaehltes_objekt = None
    if st.session_state.aktive_aktion:
        st.write("")
        with st.chat_message("assistant"):
            st.markdown(HMI_TEXTE.get(st.session_state.aktive_aktion, "Wie kann ich dir helfen?"))
        
        if df_wissen is not None and not df_wissen.empty:
            bez_spalte = df_wissen.columns[1] if len(df_wissen.columns) > 1 else df_wissen.columns[0]
            verfuegbare_bezeichnungen = df_wissen[bez_spalte].dropna().astype(str).str.strip().unique().tolist()
            verfuegbare_bezeichnungen = sorted([b for b in verfuegbare_bezeichnungen if b and b != "nan"])
            
            gewaehltes_objekt = st.selectbox(
                label="Hidden_Object_Label",
                options=verfuegbare_bezeichnungen,
                index=None,
                placeholder="Betroffenes Objekt / Ausstattung wählen (optional)",
                key=f"object_wahl_{st.session_state.aktive_aktion}",
                label_visibility="collapsed"
            )

st.write("---")
# Nachrichten-Historie zeichnen
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# Chat-Eingabe verarbeiten
if prompt := st.chat_input("Bitte schreibe hier oder sprich mit mir 🎙️"):
    if nutzer_rolle is None:
        st.warning("Bitte wähle oben zuerst aus, wer du bist!")
    elif not st.session_state.aktive_aktion:
        st.warning("Bitte wähle oben zuerst ein Anliegen aus!")
    else:
        with st.chat_message("user"):
            st.markdown(prompt)
        st.session_state.messages.append({"role": "user", "content": prompt})
        
        # Interaktives Write-Back ausführen für Störungen, Feedback oder Informationen
        if st.session_state.aktive_aktion in ["Störung", "Feedback", "Information"] and df_wissen is not None:
            with st.spinner("Eintrag wird in der Google Drive Datenbank verarbeitet..."):
                erfolg = save_data_to_drive(
                    df=df_wissen, 
                    service=drive_service, 
                    aktions_typ=st.session_state.aktive_aktion, 
                    text_inhalt=prompt, 
                    gewaehltes_objekt=gewaehltes_objekt,
                    rolle=nutzer_rolle
                )
                if erfolg:
                    # Daten sofort neu laden, damit die KI den aktuellsten Stand besitzt
                    df_wissen, _ = load_data_from_drive()

        # Kontext für die KI vorbereiten
        kontext_daten = f"\n\nAktuelle Echtzeit-Datenbank der Villa:\n{df_wissen.to_string(index=False)}" if df_wissen is not None else ""
        
        with st.chat_message("assistant"):
            with st.spinner("Villa Avatar überlegt..."):
                antwort_text = generate_ki_response(
                    f"Rolle des Benutzers: {nutzer_rolle}\n"
                    f"Aktuelle HMI-Aktion: {st.session_state.aktive_aktion}\n"
                    f"Gewähltes Datenbank-Objekt (Spalte B): {gewaehltes_objekt or 'Keines ausgewählt'}\n"
                    f"Direkte Benutzereingabe: {prompt} {kontext_daten}"
                )
            st.markdown(antwort_text)
            st.session_state.messages.append({"role": "assistant", "content": antwort_text})
