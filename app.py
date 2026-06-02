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
        return pd.read_excel(fh), service
    except Exception as e:
        st.error(f"Fehler bei der Verbindung zur Google Drive Wissensbasis: {e}")
        return None, None

# Visueller Lade-Spinner beim Abruf der Daten
with st.spinner("Verbindung zur Google Drive Wissensbasis wird hergestellt..."):
    df_wissen, drive_service = load_data_from_drive()

# ==========================================
# 2. LIVE-UPDATE IN GOOGLE DRIVE SCHREIBEN
# ==========================================
def append_info_to_drive(df, neuer_text, nutzername, kategorie="Nicht definiert"):
    try:
        neue_zeile = {
            "Zeitstempel": datetime.datetime.now().strftime("%d.%m.%Y %H:%M"),
            "Nutzer": nutzername,
            "Kategorie": kategorie,
            "Eintrag / Update": neuer_text
        }
        
        df_aktualisiert = pd.concat([df, pd.DataFrame([neue_zeile])], ignore_index=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_aktualisiert.to_excel(writer, index=False)
        output.seek(0)
        
        media = MediaIoBaseUpload(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        drive_service.files().update(fileId=FILE_ID, media_body=media).execute()
        st.cache_data.clear()  
        return True
    except Exception as e:
        st.error(f"Fehler beim Schreiben in Google Drive: {e}")
        return False

# ==========================================
# 3. KI-GEHIRN (OFFIZIELLES GOOGLE-GENAI SDK)
# ==========================================
VILLA_PROMPT = """
Du bist „Villa Avatar“, der digitale Helfer für die Bewohner und Helfer der Villa. Deine Aufgabe ist es, den Betrieb und Erhalt des Hauses so einfach wie möglich zu halten.
Beziehe dich bei allgemeinen Abläufen auf 'Villa Wissen_72.jfif' und bei der Wasserversorgung auf 'PXL_20260516_202437801_72.jpg'.

WICHTIGER KONTEXT & VERHALTEN:
- Antworte immer kurz, präzise und smartphone-optimiert.
- Nutze die vom HMI übergebene Rolle und die gewählte Kategorie/Bezeichnung zwingend als Arbeitsgrundlage.
- Wenn das HMI dir eine konkrete Bezeichnung (z. B. "Beregnungssystem") übergibt, beziehe dich exakt darauf.
- Wenn das HMI KEINE konkrete Bezeichnung übergibt (weil der Nutzer das Drop-down ignoriert hat), musst du anhand der Frage des Nutzers logisch nachdenken und den Kontext selbstständig der passenden Kategorie (Systeme, Geräte innen, Geräte außen) zuordnen.
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
            config=types.GenerateContentConfig(
                system_instruction=VILLA_PROMPT
            )
        )
        return response.text
    except Exception as e:
        try:
            response = client.models.generate_content(
                model="gemini-1.5-flash",
                contents=prompt_text,
                config=types.GenerateContentConfig(
                    system_instruction=VILLA_PROMPT
                )
            )
            return response.text
        except Exception:
            return f"Fehler bei der KI-Verarbeitung: {e}"

# ==========================================
# 4. BENUTZEROBERFLÄCHE (HMI)
# ==========================================
# Issue 17: Neuer Titel im Browsertab
st.set_page_config(page_title="Villa Avatar", page_icon="☀️", layout="centered")

# FIX für Issue 15/18: Korrekter Parametername "unsafe_allow_html"
st.markdown("""
    <style>
    div[data-testid="stSelectbox"] label { font-weight: bold; font-size: 15px; }
    </style>
""", unsafe_allow_html=True)

# Titel auf der Oberfläche
st.title("☀️ Villa Avatar")

# Issue 18: Neuer Begrüßungstext direkt unter dem Titel
st.markdown("Hallo! Ich bin Villa Avatar, dein digitaler **'Helfer'**! Wähle unten die Rolle aus, um zu beginnen.")

# 4.1 Rollen-Auswahl
if "vorherige_rolle" not in st.session_state:
    st.session_state.vorherige_rolle = "Bitte auswählen..."

nutzer_rolle = st.selectbox("Wer bist du?", ["Bitte auswählen...", "Besucher", "Eigentümer", "Administrator", "Handwerker/Helfer"])

# Issue 16: Automatischer UI-Reset bei Rollenwechsel
if nutzer_rolle != st.session_state.vorherige_rolle:
    st.session_state.vorherige_rolle = nutzer_rolle
    # Alle Dropdown-Auswahlen im Session State löschen, um sie zurückzusetzen
    for key in list(st.session_state.keys()):
        if key.startswith("sub_cat_"):
            del st.session_state[key]

# Chat-Verlauf im Session-State initialisieren
if "messages" not in st.session_state:
    st.session_state.messages = []

# Variablen für Buttons und Aktionen initialisieren
button_prompt = None
gewaehlte_aktion = "Allgemein"

# Wenn eine Rolle gewählt wurde, zeige "Mein Anliegen" und die Buttons
if nutzer_rolle != "Bitte auswählen...":
    st.write("---")
    
    # Issue 18: Überschrift "Mein Anliegen"
    st.subheader("Mein Anliegen:")
    
    if nutzer_rolle == "Besucher":
        col1, col2 = st.columns(2)
        with col1:
            if st.button("ℹ️ Ich brauche Hilfe.", use_container_width=True):
                button_prompt = "Was möchtest du wissen?"
                gewaehlte_aktion = "Hilfe"
        with col2:
            if st.button("⚠️ Es gibt eine Störung.", use_container_width=True):
                button_prompt = "Was ist passiert?"
                gewaehlte_aktion = "Störung"
    else:
        col1, col2, col3 = st.columns(3)
        col4, col5 = st.columns(2)
        with col1:
            if st.button("ℹ️ Ich brauche Hilfe.", use_container_width=True):
                button_prompt = "Was möchtest du wissen?"
                gewaehlte_aktion = "Hilfe"
        with col2:
            if st.button("⚠️ Es gibt eine Störung.", use_container_width=True):
                button_prompt = "Was ist passiert?"
                gewaehlte_aktion = "Störung"
        with col3:
            if st.button("📊 Ich benötige einen Bericht.", use_container_width=True):
                button_prompt = "Nenne mir bitte den Zeitraum und das Thema."
                gewaehlte_aktion = "Bericht"
        with col4:
            if st.button("📝 Ich habe neue Informationen.", use_container_width=True):
                button_prompt = "Gern nehme ich deine Informationen auf und ordne sie in meiner Wissensbasis zu."
                gewaehlte_aktion = "Information"
        with col5:
            if st.button("🛠️ Ich möchte eine Änderung am XLS vornehmen.", use_container_width=True):
                button_prompt = "Beschreibe deine Änderung so genau wie möglich."
                gewaehlte_aktion = "Änderung"

    # Issue 16: Wenn ein Button gedrückt wurde, werden die Dropdowns ebenfalls zurückgesetzt
    if button_prompt:
        for key in list(st.session_state.keys()):
            if key.startswith("sub_cat_"):
                del st.session_state[key]

    # ==========================================
    # 5. DYNAMISCHE EINZEILER DROP-DOWNS (ISSUE 15)
    # ==========================================
    st.write("")
    
    # Bestimme, welche Kategorien für welche Rolle sichtbar sind
    kategorien_fuer_rolle = []
    if nutzer_rolle == "Besucher":
        kategorien_fuer_rolle = ["Geräte / Ausst. innen", "Geräte / Ausst. außen"]
    elif nutzer_rolle in ["Eigentümer", "Administrator"]:
        kategorien_fuer_rolle = ["Systeme", "Geräte / Ausst. innen", "Geräte / Ausst. außen"]
    else: # Handwerker/Helfer
        kategorien_fuer_rolle = ["Systeme", "Geräte / Ausst. außen"]

    # Generiere dynamisch die Drop-downs als Einzeiler (Standardwert = Name der Kategorie)
    konkrete_auswahlen = {}
    
    if df_wissen is not None and not df_wissen.empty:
        # Spalten bestimmen (Spalte 0 = Kategorie, Spalte 1 = Bezeichnung)
        spalten = df_wissen.columns.tolist()
        kat_spalte = spalten[0]
        bez_spalte = spalten[1] if len(spalten) > 1 else spalten[0]

        for kat in kategorien_fuer_rolle:
            # Filtere alle Bezeichnungen aus der Excel, die zu dieser Oberkategorie gehören
            suchbegriff = kat.replace(" ", "").lower()
            mask = df_wissen[kat_spalte].astype(str).str.replace(" ", "").str.lower().str.contains(suchbegriff)
            verfuegbare_bezeichnungen = df_wissen[mask][bez_spalte].dropna().drop_duplicates().tolist()
            
            # Sortieren für bessere Übersicht
            verfuegbare_bezeichnungen = sorted([str(b) for b in verfuegbare_bezeichnungen])
            
            # Packe den Kategorienamen als ersten Eintrag in die Liste (Standard)
            dropdown_optionen = [kat] + verfuegbare_bezeichnungen
            
            # Rendere das Dropdown-Feld (Label zeigt dynamisch den Kategorienamen an)
            wahl = st.selectbox(
                f"{kat}:", 
                dropdown_optionen,
                key=f"sub_cat_{kat}"
            )
            
            # Wenn etwas anderes als der Standard gewählt wird, ist es aktiv
            if wahl != kat:
                konkrete_auswahlen[kat] = wahl

    # Verarbeitung der Button-Klicks
    if button_prompt:
        st.session_state.messages.append({"role": "user", "content": f"Aktion gewählt: {gewaehlte_aktion}"})
        
        if gewaehlte_aktion in ["Information", "Änderung"] and df_wissen is not None:
            kat_text = ", ".join(konkrete_auswahlen.keys()) if konkrete_auswahlen else "Allgemein"
            append_info_to_drive(df_wissen, f"Button-Aktion: {gewaehlte_aktion}", nutzer_rolle, kat_text)
        
        kontext = f"\n\nAktuelle Daten aus der Wissensbasis:\n{df_wissen.to_string(index=False)}" if df_wissen is not None else ""
        hmi_hinweis = f"Gewählte Konkretisierungen im HMI: {json.dumps(konkrete_auswahlen)}" if konkrete_auswahlen else "Nutzer hat kein Drop-down genutzt. Bitte analysiere seine Absicht selbstständig."
        
        with st.spinner("Villa Avatar formuliert Antwort..."):
            antwort_text = generate_ki_response(
                f"SYSTEM-BEFEHL: Der Nutzer hat den Button für '{gewaehlte_aktion}' gedrückt. "
                f"Antworte ihm exakt mit der Spezifikations-Gegenfrage: '{button_prompt}'. "
                f"Gib keine weiteren Erklärungen ab. {hmi_hinweis} {kontext}"
            )
        st.session_state.messages.append({"role": "assistant", "content": antwort_text})
        st.rerun()

# ==========================================
# 6. CHAT-ANZEIGE UND MANUELLER INPUT
# ==========================================
st.write("---")
# Zeige den bisherigen Chat-Verlauf an
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# Manueller Chat-Input ganz unten
if prompt := st.chat_input("Wie kann ich helfen? (z.B. 'Frage: Wo ist der Hauptwasserhahn?')"):
    if nutzer_rolle == "Bitte auswählen...":
        st.warning("Bitte wähle oben zuerst aus, wer du bist!")
    else:
        with st.chat_message("user"):
            st.markdown(prompt)
        st.session_state.messages.append({"role": "user", "content": prompt})
        
        # Präfix-Erkennung für Direkt-Speicherung
        if prompt.strip().lower().startswith("information:") and df_wissen is not None:
            reiner_text = prompt.split(":", 1)[1].strip()
            kat_text = ", ".join(konkrete_auswahlen.keys()) if 'konkrete_auswahlen' in locals() and konkrete_auswahlen else "Allgemein"
            with st.spinner("Speichere Update direkt in Google Drive..."):
                erfolg = append_info_to_drive(df_wissen, reiner_text, nutzer_rolle, kat_text)
            if erfolg:
                st.success("Eintrag erfolgreich in Google Drive gespeichert!")
                st.cache_data.clear()
                with st.spinner("Lade aktualisierte Wissensbasis..."):
                    df_wissen, _ = load_data_from_drive()

        # KI-Antwort generieren unter Berücksichtigung der HMI-Auswahlen
        kontext = f"\n\nAktuelle Daten aus der Wissensbasis:\n{df_wissen.to_string(index=False)}" if df_wissen is not None else ""
        hmi_hinweis = f"Ausgewählte HMI-Spezifikation: {json.dumps(konkrete_auswahlen)}" if 'konkrete_auswahlen' in locals() and konkrete_auswahlen else "Keine HMI-Konkretisierung gewählt. Nutze Kontextanalyse für den Nutzertext."
        
        with st.chat_message("assistant"):
            with st.spinner("Villa Avatar überlegt..."):
                antwort_text = generate_ki_response(
                    f"SYSTEM-KONTEXT: Der Nutzer tippt in der Rolle '{nutzer_rolle}'.\n"
                    f"{hmi_hinweis}\n"
                    f"Anfrage: {prompt} {kontext}"
                )
            st.markdown(antwort_text)
            st.session_state.messages.append({"role": "assistant", "content": antwort_text})
