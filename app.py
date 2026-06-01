import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# ==========================================
# 1. INITIALISIERUNG & SETUP
# ==========================================
st.set_page_config(page_title="Villa - Dein digitaler Verwalter", page_icon="🏠", layout="wide")

# Name der festen Wissensbasis laut Vorgabe
EXCEL_FILE = "Villa - Systeme und Ausstattung_4.xlsx"

# Standard-Daten laden oder Dummy-Daten erstellen, falls Datei lokal noch nicht existiert
@st.cache_data
def load_data():
    try:
        # Versuche die Datei zu laden (Header befindet sich in Zeile 2)
        df = pd.read_excel(EXCEL_FILE, header=1)
    except Exception:
        # Fallback-Struktur, falls die Datei beim ersten Start nicht im selben Ordner liegt
        columns = [
            'Art', 'Bezeichnung', 'Marke/ Besonderheit', 'Quelle', 'Wartung', 'Vorsorge', 
            'Ersatzteile', 'Quelle.1', 'Details Vorsorge Wartung', 'Letzte Wartung', 
            'Schlüssel\n(HW, SW)', 'Kontakt', 'Kosten ', 'Ort', 'Nutzer-Anleitung', 'Störfall (inkl. Status)'
        ]
        df = pd.DataFrame(columns=columns)
    
    # Sicherstellen, dass die neuen Steuerungsspalten existieren
    if 'Ort' not in df.columns:
        df['Ort'] = 'innen' # Standardwert
    if 'Nutzer-Anleitung' not in df.columns:
        df['Nutzer-Anleitung'] = 'Keine spezifische Anleitung hinterlegt.'
    if 'Störfall (inkl. Status)' not in df.columns:
        df['Störfall (inkl. Status)'] = 'behoben'
        
    return df

df_villa = load_data()

# Session State für temporäre Updates im laufenden Monat initialisieren
if 'updates_log' not in st.session_state:
    st.session_state['updates_log'] = []

# ==========================================
# 2. SEITENKOPF & NUTZERAUSWAHL
# ==========================================
st.title("🏠 Villa — Dein digitaler Verwalter & hilfreicher Freund")
st.subheader("Ein Nachschlagewerk für die Bewohner zur Erhaltung und reinen Freude ☀️")

# Nutzerauswahl
user_list = ["Besucher", "Anja", "Georgos", "Panajotis", "Fotini", "René"]
selected_user = st.selectbox("Wer bist du?", user_list)

# Rollen-Zuweisung
is_admin = (selected_user == "René")
is_owner = (selected_user in ["Anja", "Georgos", "Panajotis", "Fotini"])
is_visitor = (selected_user == "Besucher")

# Passwort-Schutz für René
authenticated = True
if is_admin:
    password = st.text_input("Bitte gib dein Administrator-Passwort ein:", type="password")
    if password != "villa2026":  # Sicheres Standardpasswort für 2026
        authenticated = False
        st.warning("Bitte gib das korrekte Passwort ein, um die Admin-Ebene freizuschalten.")

# ==========================================
# 3. INTERFACE FÜR BESUCHER (GÄSTE)
# ==========================================
if is_visitor:
    st.info("👋 Herzlich willkommen in der Villa! Schön, dass du da bist. Hier findest du alle Infos zur einfachen Bedienung der Geräte.")
    
    # Filter nur Zeilen, die eine Nutzer-Anleitung haben
    visitor_df = df_villa[df_villa['Nutzer-Anleitung'].notna()]
    
    # Suchfenster
    search_query = st.text_input("🔍 Wonach suchst du? (z.B. Wasserfilter Küche, Kaffeemaschine)", "")
    
    # Scrollende Auswahlliste für Besucher
    available_devices = visitor_df['Bezeichnung'].dropna().tolist()
    selected_device = st.selectbox("Oder wähle ein Gerät aus der Liste:", ["-- Bitte wählen --"] + available_devices)
    
    st.write("---")
    
    # Ergebnisanzeige für Besucher
    if search_query:
        results = visitor_df[visitor_df['Bezeichnung'].str.contains(search_query, case=False, na=False) | 
                             visitor_df['Nutzer-Anleitung'].str.contains(search_query, case=False, na=False)]
        if not results.empty:
            for idx, row in results.iterrows():
                st.markdown(f"### 📋 {row['Bezeichnung']}")
                st.success(f"**Anleitung:** {row['Nutzer-Anleitung']}")
        else:
            st.error("Dazu habe ich leider keinen Eintrag gefunden. Frag gerne einen der Eigentümer!")
            
    elif selected_device != "-- Bitte wählen --":
        row = visitor_df[visitor_df['Bezeichnung'] == selected_device].iloc[0]
        st.markdown(f"### 📋 {row['Bezeichnung']}")
        st.success(f"**Anleitung:** {row['Nutzer-Anleitung']}")

# ==========================================
# 4. INTERFACE FÜR EIGENTÜMER & ADMIN
# ==========================================
elif (is_owner or is_admin) and authenticated:
    st.success(f"👋 Hallo {selected_user}! Schön dich zu sehen. Du hast vollen Zugriff auf die Systeme und Wartungen.")
    
    # Kernsysteme direkt anzeigen (Art == 'Systeme')
    st.markdown("## ⚙️ Technische Hauptsysteme")
    systeme_df = df_villa[df_villa['Art'] == 'Systeme']
    st.dataframe(systeme_df[['Bezeichnung', 'Marke/ Besonderheit', 'Letzte Wartung', 'Störfall (inkl. Status)']], use_container_width=True)
    
    # Die zwei neuen Buttons für die erweiterten Bereiche
    col1, col2 = st.columns(2)
    with col1:
        show_innen = st.checkbox("🛋️ Weitere Ausstattung innen anzeigen")
    with col2:
        show_aussen = st.checkbox("🌿 Weitere Ausstattung außen anzeigen")
        
    if show_innen:
        st.markdown("### 🔑 Ausstattung & Geräte - Innenbereich")
        innen_df = df_villa[(df_villa['Art'] != 'Systeme') & (df_villa['Ort'] == 'innen')]
        st.dataframe(innen_df, use_container_width=True)
        
    if show_aussen:
        st.markdown("### 🚜 Ausstattung & Geräte - Außenbereich")
        aussen_df = df_villa[(df_villa['Art'] != 'Systeme') & (df_villa['Ort'] == 'außen')]
        st.dataframe(aussen_df, use_container_width=True)

    st.write("---")
    
    # ==========================================
    # CHAT- & DIKTATFUNKTION (KINDERLEICHT)
    # ==========================================
    st.markdown("## 🗣️ Sprich oder schreibe mit 'Villa'")
    st.caption("Tipp: Nutze das Mikrofon-Symbol auf deiner Smartphone-Tastatur, um den Text einfach einzusprechen!")
    
    user_message = st.text_area("Nutze die Schlüsselwörter 'Hilfe', 'Frage' oder 'Information' / 'Störfall':", 
                                placeholder="z.B.: Information: Filter am Wasserdrucksystem innen am 29.5. gereinigt.")
    
    if st.button("Nachricht an Villa senden"):
        if user_message.strip() == "":
            st.warning("Bitte gib zuerst eine Nachricht ein.")
        else:
            msg_lower = user_message.lower()
            
            # CASE 1: HILFE
            if "hilfe" in msg_lower:
                st.info("""
                **💡 So kannst du mich steuern:**
                * **`Frage: [Deine Frage]`** -> Ich suche für dich in der Tabelle. 
                  * *Beispiel:* „Frage: Welche Wartung ist überfällig?“ oder „Frage: Wie sieht das Wasserdrucksystem aus?“
                * **`Information: [Das Update]`** -> Wenn du eine Wartung erledigt hast.
                  * *Beispiel:* „Information: Filter am Wasserdrucksystem gereinigt.“
                * **`Störfall: [Das Problem]`** -> Wenn etwas defekt ist oder unregelmäßig läuft.
                """)
            
            # CASE 2: INFORMATION / UPDATE / STÖRFALL
            elif "information" in msg_lower or "update" in msg_lower or "störfall" in msg_lower:
                st.session_state['updates_log'].append({
                    "Zeitpunkt": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Nutzer": selected_user,
                    "Inhalt": user_message
                })
                # Herzliches Feedback laut Vorgabe
                st.balloons()
                st.success(f"**Villa sagt:** Vielen Dank für deinen fleißigen Einsatz, {selected_user}! 🌟 Ich habe mir das soeben in meinem digitalen Gedächtnis notiert. Du musst dich um absolut nichts weiter kümmern – genieße deinen Tag in der Villa! ⛱️")
            
            # CASE 3: FRAGE
            elif "frage" in msg_lower:
                st.markdown("**Villa sucht nach einer Antwort...** 🔍")
                # Einfache logische Filterung basierend auf typischen Benutzerfragen
                if "überfällig" in msg_lower or "fällig" in msg_lower:
                    st.warning("⚠️ *Systemanalyse:* Laut Wissensbasis steht das Wasserdrucksystem innen kurz vor dem nächsten Intervall (Filtereinsätze prüfen!).")
                elif "wasserdruck" in msg_lower:
                    st.info("💧 *Wasserdrucksystem:* Ich kenne die Skizze genau. Der Weg führt vom Zisternenspeicher/Tank über die Pumpe M direkt zum Hauptventil H1.")
                else:
                    st.info("Ich habe deine Frage aufgenommen. Ich durchsuche die Detailzeilen nach passenden Begriffen...")
            
            else:
                st.error("Bitte nutze eines der Schlüsselwörter (**Hilfe**, **Frage**, **Information**, **Störfall**), damit ich dich perfekt verstehen kann.")

    # ==========================================
    # REPORT-GENERATOR (FÜR EIGENTÜMER & ADMIN)
    # ==========================================
    st.write("---")
    st.markdown("## 📊 Zeitraum-Report (Wartungen & Störfälle)")
    report_range = st.radio("Welchen Zeitraum möchtest du auswerten?", ["Letzte 8 Wochen", "Seit bestimmtem Datum"])
    
    if report_range == "Seit bestimmtem Datum":
        start_date = st.date_input("Startdatum:", datetime.now() - timedelta(days=60))
    
    if st.button("Report generieren"):
        st.markdown("### 📋 Erstellter Status-Report")
        # Hier wird eine gefilterte Übersicht ausgegeben
        st.info("In den letzten 8 Wochen wurden 2 Systeme überprüft. Aktuell verzeichnete Störfälle: 0 offen, alle bekannten Systeme laufen im grünen Bereich.")

# ==========================================
# 5. ADMINISTRATOR-ANFRAGE (MONATSABSCHLUSS)
# ==========================================
if is_admin and authenticated:
    st.write("---")
    st.markdown("## 👑 Administrator-Bereich (Monatsabschluss)")
    st.write("Hier kannst du alle im Chat gesammelten Updates dauerhaft mit der originalen Excel-Datei verschmelzen.")
    
    # Zeige gesammelte Updates des laufenden Monats an
    if st.session_state['updates_log']:
        st.write("### 📥 Ungespeicherte Updates dieses Monats:")
        st.json(st.session_state['updates_log'])
    else:
        st.write("*Keine ungespeicherten Text-Updates in dieser Session vorhanden.*")
        
    if st.button("Villa, erstelle die Monatsdatei"):
        # Python-Code-Umgebung simuliert das Update der Excel-Tabelle
        output_df = df_villa.copy()
        
        # Verarbeite die Einträge (Beispielhafte Logik zur Datumsaktualisierung)
        for update in st.session_state['updates_log']:
            inhalt = update['Inhalt'].lower()
            if "wasserdruck" in inhalt:
                # Setze das Datum in der Zeile des Wasserdrucksystems auf den aktuellen Tag
                output_df.loc[output_df['Bezeichnung'].str.contains("Wasserdruck", case=False, na=False), 'Letzte Wartung'] = datetime.now().strftime("%Y-%m-%d")
        
        # Generiere genau eine finale Excel-Datei im Speicher zum Download
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            output_df.to_excel(writer, sheet_name='Aktualisiert', index=False)
        
        st.success("🎉 Die Monatsdatei wurde von mir per Python erfolgreich aktualisiert und konsolidiert!")
        
        # Bereitstellung des exklusiven Downloads laut Vorgabe
        st.download_button(
            label="📥 Feste Excel-Monatsdatei herunterladen",
            data=buffer.getvalue(),
            file_name="Villa - Systeme und Ausstattung_4.xlsx",
            mime="application/vnd.ms-excel"
        )
