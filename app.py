import streamlit as st
import google.generativeai as genai
import pandas as pd
from google.oauth2.service_account import Credentials
import gspread

# 1. SEITEN-KONFIGURATION & SMARTPHONE-OPTIMIERUNG
st.set_page_config(
    page_title="Villa Avatar",
    page_icon="🏡",
    layout="centered", # 'centered' eignet sich besser für Mobile/Smartphones
    initial_sidebar_state="collapsed"
)

# CSS für Smartphone-Optimierung (Kacheln und Chat-Layout)
st.markdown("""
<style>
    .status-card {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 10px;
    }
    .stChatMessage {
        border-radius: 15px;
    }
</style>
""", unsafe_allow_html=True)

# 2. INITIALISIERUNG & API-KEYS
if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
else:
    st.error("Bitte GEMINI_API_KEY in den Streamlit Secrets hinterlegen.")
    st.stop()

# 3. DATEN-ZUGRIFF (WISSENSBASIS VIA SERVICE ACCOUNT)
@st.cache_data(ttl=60) # Cache für 60 Sekunden, um API-Limits zu schonen
def load_wissensbasis():
    try:
        # Erstelle Credentials aus den korrekten Streamlit Secrets (GOOGLE_CREDENTIALS)
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(st.secrets["GOOGLE_CREDENTIALS"], scopes=scope)
        client = gspread.authorize(creds)
        
        # Öffne das Spreadsheet (Nutze die ID aus deinem Link)
        sheet_id = "1FzhWZuO6aRZkdRuQBzaojhkq7bQDyprl"
        sheet = client.open_by_key(sheet_id).sheet1 # Liest das erste Tabellenblatt
        
        # In DataFrame laden (Köpfe in Zeile 1 werden automatisch genutzt)
        data = sheet.get_all_records()
        df = pd.DataFrame(data)
        return df
    except Exception as e:
        st.error(f"Fehler beim Laden der Wissensbasis: {e}")
        return pd.DataFrame()

df_knowledge = load_wissensbasis()

# 4. BENUTZEROBERFLÄCHE (NUTZER-EBENE)
st.title("🏡 Villa Avatar")
st.subheader("Dein digitaler Begleiter")

# Rollenauswahl (Gast, Host, Admin)
if "role" not in st.session_state:
    st.session_state.role = "Gast"

col1, col2, col3 = st.columns(3)
with col1:
    if st.button("🧳 Gast", use_container_width=True): st.session_state.role = "Gast"
with col2:
    if st.button("🔑 Host", use_container_width=True): st.session_state.role = "Host"
with col3:
    if st.button("🛠️ Admin", use_container_width=True): st.session_state.role = "Admin"

st.caption(f"Aktuelle Ansicht optimiert für: **{st.session_state.role}**")

# STATUS-KACHELN (Beispielhafte Darstellung basierend auf Live-Daten)
st.markdown("### 📊 System-Status")
kachel_cols = st.columns(2)
with kachel_cols[0]:
    st.markdown('<div class="status-card"><strong>WLAN</strong><br/>✅ Online</div>', unsafe_allow_html=True)
with kachel_cols[1]:
    # Dynamischer Wert aus der Tabelle gezogen (Beispiel)
    status_pool = "Bereit" if not df_knowledge.empty else "Keine Daten"
    st.markdown(f'<div class="status-card"><strong>Pool-Temperatur</strong><br/>🏊 {status_pool}</div>', unsafe_allow_html=True)

# 5. CHAT-LOGIK & PROMPT-INJEKTION
VILLA_PROMPT_TEMPLATE = """
Du bist „Villa Avatar“, der digitale Helfer für die Gäste (Gast), Gastgeber (Host) und Administratoren (Admin) der Villa. Deine Aufgabe ist es, den Aufenthalt, Betrieb und Erhalt des Hauses so einfach und angenehm wie möglich zu gestalten.

WICHTIGER KONTEXT & VERHALTEN:
- Antworte immer kurz, freundlich, präzise und smartphone-optimiert (nutze kurze Absätze oder Aufzählungspunkte).
- Passe deine Tonalität an die übergebene Rolle an: Zu Gästen (Gast) bist du einladend und herzlich, zu Hosts und Admins agierst du effizient und lösungsorientiert.
- Beziehe dich bei Antworten exakt auf die mitgegebenen Live-Daten aus der Wissensbasis. 
- WICHTIG (Wissenslücken): Wenn die mitgegebenen Daten keine Antwort auf die Frage des Nutzers enthalten, erfinde niemals Informationen! Antworte stattdessen: „Dazu liegen mir aktuell leider keine Informationen vor. Ich leite dein Anliegen aber gerne an das Team weiter.“
- ABSOLUTES VERBOT: Erwähne NIEMALS interne Dateinamen, Bildbezeichnungen (wie '.jfif' oder '.jpg') oder die Struktur der Excel-Tabelle (wie 'Spalte A', 'Spalte B', Überschriften oder Zeilen). Antworte so, als hättest du dieses Wissen einfach natürlich im Kopf.

AKTUELLE ROLLE DES NUTZERS: {role}

LIVE-DATEN AUS DER WISSENSBASIS (Nutze diese exklusiv für Sachfragen):
{knowledge_data}
"""

# Chat-Verlauf initialisieren
if "messages" not in st.session_state:
    st.session_state.messages = []

# Chat-Verlauf anzeigen
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# Chat-Eingabe (Tastatur / Spracheingabe des Smartphones wird nativ unterstützt)
if user_input := st.chat_input("Wie kann ich dir heute helfen?"):
    
    # Nutzer-Nachricht anzeigen und speichern
    st.session_state.messages.append({"role": "user", "content": user_input})
    with st.chat_message("user"):
        st.markdown(user_input)
        
    # Kontext für Gemini vorbereiten
    knowledge_str = df_knowledge.to_markdown(index=False) if not df_knowledge.empty else "Keine Daten verfügbar."
    
    system_instruction = VILLA_PROMPT_TEMPLATE.format(
        role=st.session_state.role,
        knowledge_data=knowledge_str
    )
    
    # Gemini Modell aufrufen (Nutzt system_instruction für strikte Verhaltensregeln)
    try:
        model = genai.GenerativeModel(
            model_name="gemini-1.5-pro",
            system_instruction=system_instruction
        )
        
        # Chat-Historie für das Modell übersetzen (ohne System-Prompts, da diese oben fixiert sind)
        chat = model.start_chat(history=[
            {"role": "user" if m["role"] == "user" else "model", "parts": [m["content"]]}
            for m in st.session_state.messages[:-1]
        ])
        
        # Antwort generieren
        with st.chat_message("assistant"):
            response_placeholder = st.empty()
            response = chat.send_message(user_input)
            
            # Antwort anzeigen und speichern
            response_placeholder.markdown(response.text)
            st.session_state.messages.append({"role": "assistant", "content": response.text})
            
    except Exception as e:
        st.error(f"Fehler bei der KI-Verarbeitung: {e}")
