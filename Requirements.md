# Requirements & Systemarchitektur — Villa Asset Management App

Dieses Dokument definiert die funktionalen Anforderungen, Datenstrukturen und Sicherheitsmechanismen für die KI-gestützte Asset-Management-Anwendung der Villa. Das System folgt strikt dem Prinzip **„Mensch zu Maschine“**: Die Excel-Datei im Google Drive ist die alleinige administrative Wahrheit; die Anwendung passt sich Änderungen dieser Datei vollständig dynamisch an.

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

*Ausnahme (System-Anker):* Die Spalten `Bezeichnung`, `Wo?`, `Relevanz Gast`, sowie die Input/Status-Felder für Störungen, Feedbacks und „Keine Information“ sind im Code namentlich hinterlegt, um die HMI-Eingabemasken und Kernfilter zu steuern.

---

## 4. HMI-Schnittstelle & Use Cases

Die Benutzeroberfläche (Streamlit HMI) verzichtet auf administrative Schaltflächen und unterscheidet strikt zwei Rollen:

### Rolle: Gast
* **Use Case 1: Ich brauche Hilfe / Frage stellen** -> KI-gestützte Beantwortung. Kontext ist streng zweidimensional gefiltert (nur Zeilen mit `x`, nur Spalten mit `ja`).
* **Use Case 2: Eine Störung melden** -> Freitextfeld. Der Eintrag wird per Append-Logik in die Spalte `Störung [Input]` geschrieben. Der `Störung Status` wird automatisch auf `aktiv` gesetzt.
* **Use Case 3: Feedback geben** -> Freitextfeld. Eintrag wird per Append-Logik in `Feedback [Input]` geschrieben; `Feedback Status` wird auf `offen` gesetzt.
* **Use Case 4: Unbeantwortete Fragen / Keine Information** -> Eintrag wird per Append-Logik in `Keine Information` geschrieben und der Status auf `offen` gesetzt.

### Rolle: Host
* **Use Case 1: Ich brauche Hilfe / Frage stellen** -> KI-Zugriff auf die gesamte Matrix (alle Zeilen, alle Spalten).
* **Use Case 2: Neue Informationen eintragen** -> Ermöglicht das Anhängen von operativen Notizen an bestehende Objekte (Append-Logik).
* **Use Case 3: Störung melden / verwalten** -> Kann Störungen melden oder den Status einsehen.
* **Use Case 4: Feedback erfassen** -> Dokumentation von Feedback im System.
* **Use Case 5: Unbeantwortete Fragen protokollieren** -> Festhalten von Wissenslücken aus Gastgesprächen.
* **Use Case 6: Bericht generieren** -> Die KI analysiert periodische Daten (z. B. alle aktiven Störungen, offene Feedbacks oder ungelöste Fragen) und erstellt eine strukturierte Zusammenfassung für den Host.

---

## 5. Technische Implementierungs-Vorgaben

* **Dynamische Spalten-Findung:** Im Code dürfen keine statischen Spaltenindizes verwendet werden (z. B. nicht `col=11`). Die App sucht die Spaltennummern beim Start dynamisch über die Namen aus der Header-Zeile.
* **Append-Logik bei Eingaben:** Bestehende Zellinhalte in den Input-Spalten (`Störung`, `Feedback`, `Keine Information`) dürfen *niemals* überschrieben werden. Neue Einträge werden chronologisch mit Zeitstempel und Rolle angehängt (z. B. `\n[05.06.2026 - Gast]: Text`).
* **KI-Sicherheitsnetz (Fallback):** Anfragen werden primär an `gemini-2.5-flash` gestellt. Bei Überlastung oder Quoten-Limits fängt der Code den Fehler ab und schaltet automatisch auf `gemini-2.0-flash` um.
* **Token-Spargang:** Um API-Quoten zu schonen, wird der KI nicht immer die gesamte Tabelle übergeben. Über Dropdowns (z. B. Bereichsauswahl „Wo?“) wird der Tabellen-Kontext für die KI vorab gefiltert.
