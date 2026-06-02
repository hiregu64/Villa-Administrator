import streamlit as st
from google import genai
from google.genai import types
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import pandas as pd
import io
import datetime

# --- 1. KONFIGURATION & ICONS ---
FILE_ID = '1FzhWZuO6aRZkdRuQBzaojhkq7bQDyprl'

# Das finale, korrekte grüne Icon
GREEN_ICON_SVG = """
<div style='width: 32px; height: 32px; background-color: #2e7d32; border-radius: 8px; display: flex; align-items: center; justify-content: center;'>
    <svg viewBox='0 0 24 24' width='20' height='20' fill='none' stroke='white' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'>
        <circle cx='12' cy='12' r='10'></circle>
        <path d='M8 14s1.5 2 4 2 4-2 4-2'></path>
        <circle cx='9' cy='9.5' r='1' fill='white' stroke='none'></circle>
        <circle cx='15' cy='9.5' r='1' fill='white' stroke='none'></circle>
        <path d='M8 9c1-1.5 2.5-2 4-2s3 .5 4 2'></path>
    </svg>
</div>
"""

# --- 2. DATEN & KI-FUNKTIONEN ---
@st.cache_data(ttl=30)
def load_data_from_drive():
    try:
        creds = service_account.Credentials.from_service_account_info(st.secrets["GOOGLE_CREDENTIALS"])
        service = build('drive', 'v3', credentials=creds)
        request = service.files().get_media(fileId=FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done: _, done = downloader.next_chunk()
        fh.seek(0)
        df = pd.read_excel(fh)
        if not df.empty: df.iloc[:, 0] = df.iloc[:, 0].ffill()
        return df, service
    except: return None, None

def generate_ki_response(prompt_text):
    client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
    try:
        response = client.models.generate_content(model="gemini-2.5-flash", contents=prompt_text, config=types.GenerateContentConfig(system_instruction="Du bist „Villa Avatar“. Antworte kurz, präzise und hilfsbereit."))
        return response.text
    except Exception as e:
        return "⏳ Die Server sind ausgelastet."

# --- 3. UI & STYLING ---
st.set_page_config(page_title="Villa Avatar", page_icon="☀️", layout="centered")
st.markdown("""
    <style>
    button[data-testid="stBaseButton-primary"] { background-color: #e1f5fe !important; color: #0288d1 !important; border: 1px solid #b3e5fc !important; }
    </style>
""", unsafe_allow_html=True)

df_wissen, drive_service = load_data_from_drive()
st.title("☀️ Villa Avatar")

# Rollen & Matrix-Logik
nutzer_rolle = st.selectbox("Wer bist du?", ["Besucher", "Eigentümer", "Admin"], index=None, label_visibility="collapsed", placeholder="Wer bist du?")

if nutzer_rolle:
    st.write("---")
    c1, c2 = st.columns([0.8, 0.2])
    with c1: st.subheader("Mein Anliegen:")
    with c2: st.markdown(GREEN_ICON_SVG, unsafe_allow_html=True)
    
    # HIER IST DIE GANZE LOGIK WIEDER:
    if nutzer_rolle == "Besucher":
        col1, col2 = st.columns(2)
        if col1.button("Hilfe"): st.session_state.aktion = "Hilfe"
        if col2.button("Störung"): st.session_state.aktion = "Störung"
    elif nutzer_rolle == "Admin":
        col1, col2, col3 = st.columns(3)
        if col1.button("Hilfe"): st.session_state.aktion = "Hilfe"
        if col2.button("Info"): st.session_state.aktion = "Info"
        if col3.button("Störung"): st.session_state.aktion = "Störung"
    # ... (weitere Rollen ergänzen)

    # Chat-Input
    if prompt := st.chat_input("Was beschäftigt dich?"):
        with st.chat_message("user"): st.markdown(prompt)
        with st.chat_message("assistant"):
            st.markdown(f"{GREEN_ICON_SVG} Villa Avatar überlegt...")
            st.markdown(generate_ki_response(prompt))
