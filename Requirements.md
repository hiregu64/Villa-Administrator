# Requirements & Systemarchitektur — Villa Asset Management App

Dieses Dokument definiert die funktionalen Anforderungen, Datenstrukturen und Sicherheitsmechanismen für die KI-gestützte Asset-Management-Anwendung der Villa. Das System folgt strikt dem Prinzip **„Mensch zu Maschine“**: Die Excel-Datei im Google Drive ist die alleinige administrative Wahrheit; die Anwendung passt sich Änderungen dieser Datei vollständig dynamisch an. 

**HINWEIS ZUR MASTER-LOGIK (Single Source of Truth):** Die logische Matrix der Präsentation (GEM) definiert die unumstößliche Struktur der Benutzeroberfläche und des Datenflusses. Bei Diskrepanzen zwischen textlichen Beschreibungen und dem GEM gilt das GEM als primäre Wahrheit.

---

## 1. Architektur-Prinzip (Die Zwei-Blatt-Struktur)

Die Anwendung liest live eine Excel-Arbeitsmappe (`Villa - Systeme und Ausstattung.xlsx`), die aus zwei funktionalen Tabellenblättern besteht:

1. **`Wissensbasis` (Tabelle 1):** Enthält die konkreten Daten, Beschreibungen, Zustände und Historien aller Geräte, Systeme und Ressourcen der Villa (Zeilen = Objekte).
2. **`Spalten_Lexikon` (Tabelle 2):** Definiert die Metadaten, sprich die Bedeutung jeder Spalte, das erwartete Datenformat sowie die Zugriffsberechtigungen für die KI (Zeilen = Spaltennamen von Tabelle 1).

---

## 2. Der Zwei-Achsen-Sicherheitsfilter (Matrix-Filterung)

Um kritische interne Informationen (wie Kosten, Kontakte oder interne Steuerungsdetails) rigoros vor Gästen zu schützen, wendet die App vor jeder KI-Verarbeitung einen automatischen, zweidimensionalen Filter an:

### Achse 1: Zeilen-Filter (Welche Objekte existieren für den Gast?)
* **Regel:** Die App prüft in Tabelle 1 die Spalte `Relevanz Gast`.
* **Logik:** Nur Zeilen, die in dieser Spalte ein `x` enthalten, werden für die Rolle *Gast* geladen. Zeilen ohne `x` werden für den Gast komplett unsichtbar gemacht.

### Achse 2: Spalten-Filter (Welche Details darf der Gast sehen?)
* **Regel:** Die App prüft in Tabelle 2 die Spalte `Sichtbar für Gast`.
* **Logik:** Bei den freigegebenen Objekten werden für die Rolle *Gast* ausschließlich die Spalten an die KI übergeben, bei denen im Spalten-Lexikon ein `ja` hinterlegt ist. Spalten mit `nein` werden im Speicher gelöscht, bevor die KI den Kontext erhält.

---

## 3. Dynamisches Datenmodell (Schema-agnostisch)

Das System besitzt **keine** fest einprogrammierten Inhalts-Spalten. Das Datenmodell bestimmt sich bei jedem App-Start vollständig dynamisch aus der Excel-Datei nach folgendem Algorithmus:

1. **Schema-Erkennung:** Der Code liest die Header-Zeile von `Wissensbasis` und die Zeilen von `Spalten_Lexikon` ein.
2. **Rechte-Mapping:** Der Code erstellt zur Laufzeit eine Liste aller Spalten, bei denen im Lexikon `Sichtbar für Gast == "ja"` steht.
3. **Echtzeit-Filterung:** * Wählt der Nutzer die Rolle **Host**, bekommt die KI alle in Excel existierenden Spalten als Kontext übergeben.
   * Wählt der Nutzer die Rolle **Gast**, wirft der Code im RAM alle Spalten heraus, die im Lexikon nicht explizit als `ja` deklariert sind.
4. **Offline-Erweiterbarkeit:** Fügt der Admin offline eine neue Spalte in Tabelle 1 hinzu und pflegt sie in Tabelle 2 ein, wird sie von der App sofort ohne Code-Änderung berücksichtigt.

*Ausnahme (System-Anker):* Die Spalten `Bezeichnung`, `Wo?`, `Relevanz Gast`, sowie die Input/Status-Felder für Störungen, Feedbacks und „Keine Information“ (bzw. deren feldnahe Bezeichner wie `[Input]`) sind im Code namentlich hinterlegt, um die HMI-Eingabemasken und Kernfilter fehlertolerant zu steuern.

---

## 4. HMI-Schnittstelle & Use Cases (Abgestimmt auf GEM)

Die Benutzeroberfläche (Streamlit HMI) bietet ausschließlich voll ausformulierte Sätze als Interaktions-Buttons an. Es gibt keine separaten administrativen Schaltflächen oder Freitext-Buttons für System-Logs.

### Rolle: Gast
* **Use Case 1: „Ich brauche Hilfe.“** -> KI-gestützte Beantwortung von Fragen zur Villa. Kontext ist streng zweidimensional gefiltert (nur Zeilen mit `x`, nur Spalten mit `ja`).
* **Use Case 2: „Ich möchte eine Störung melden.“** -> Öffnet das HMI-Eingabefeld. Der Eintrag wird per Append-Logik in die Spalte `Störung` geschrieben. Der `Störung Status` wird im Hintergrund automatisch auf `aktiv` gesetzt.
* **Use Case 3: „Ich möchte Feedback geben.“** -> Öffnet das HMI-Eingabefeld. Eintrag wird per Append-Logik in `Feedback` geschrieben; `Feedback Status` wird auf `offen` gesetzt.
* **Hidden Use Case: Automatische Protokollierung bei Wissenslücken** -> Es existiert *kein* sichtbarer Button für den Gast. Wenn der Gast eine Frage über den Hilfe-Button stellt und die KI mangels Daten mit dem festen Fallback-Satz antwortet, routet das System den Prompt automatisch als Input in die Excel-Spalte `Keine Information` und setzt den zugehörigen Status auf `offen`.

### Rolle: Host
* **Use Case 1: „Ich brauche Hilfe.“** -> KI-Zugriff auf die gesamte Matrix (alle Zeilen, alle Spalten) für administrative Fragen.
* **Use Case 2: „Ich habe neue Informationen.“** -> Ermöglicht das Anhängen von operativen Notizen an bestehende Objekte (Append-Logik). Ein KI-Classifier ordnet die Information automatisch der passenden Stammdaten-Zielspalte zu.
* **Use Case 3: „Ich möchte eine Störung melden.“** -> Dokumentation und Erfassung von Systemausfällen.
* **Use Case 4: „Ich möchte Feedback geben.“** -> Direktes Abspeichern von Feedback-Einträgen im System.
* **Use Case 5: „Ich benötige einen Bericht.“** -> Die KI analysiert periodische Daten (z. B. alle aktiven Störungen, offene Feedbacks oder ungelöste Fragen) und erstellt eine strukturierte Zusammenfassung für den Host direkt im Chatfenster.

---

## 5. Technische Implementierungs-Vorgaben

* **Dynamische Spalten-Findung & Fuzzy-Matching:** Im Code dürfen keine statischen Spaltenindizes verwendet werden. Da Spaltennamen in der Praxis variieren können (z. B. `Störung` vs. `Störung [Input]`), muss die Anwendung ein automatisiertes, fehlertolerantes String-Matching (Fuzzy-Matching) anwenden, um die korrekten Zielspalten anhand der Header-Zeile zu ermitteln. Schlägt dies fehl, muss eine klare Fehlermeldung im UI ausgegeben werden.
* **Append-Logik bei Eingaben:** Bestehende Zellinhalte in den Input-Spalten (`Störung`, `Feedback`, `Keine Information`) dürfen *niemals* überschrieben werden. Neue Einträge werden chronologisch mit Zeitstempel und Rolle angehängt (z. B. `\n- [05.06.2026 | Gast]: Text`).
* **KI-Sicherheitsnetz (Fallback):** Anfragen werden primär an `gemini-2.5-flash` gestellt. Bei Überlastung oder Quoten-Limits fängt der Code den Fehler ab und schaltet automatisch auf `gemini-2.0-flash` um.
* **Token-Spargang:** Um API-Quoten zu schonen, wird der KI nicht immer die gesamte Tabelle übergeben. Über Dropdowns (Bereichsauswahl „Wo?“) wird der Tabellen-Kontext für die KI vorab gefiltert.
