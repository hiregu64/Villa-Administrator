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
* **Logik:** Nur Zeilen, die in dieser Spalte ein `x` enthalten, werden für die Rolle *Gast* geladen. Zeilen ohne `x` (z. B. rein interne Infrastruktur) werden für den Gast komplett unsichtbar gemacht.

### Achse 2: Spalten-Filter (Welche Details darf der Gast sehen?)
* **Regel:** Die App prüft in Tabelle 2 die Spalte `Sichtbar für Gast`.
* **Logik:** Bei den freigegebenen Objekten werden für die Rolle *Gast* ausschließlich die Spalten an die KI übergeben, bei denen im Spalten-Lexikon ein `ja` hinterlegt ist. Spalten mit `nein` (z. B. *Kosten*, *Kontakt*, *Details Steuerung*) werden im Speicher gelöscht, bevor die KI den Kontext erhält.

---

## 3. Datenmodell & Spalten-Synchronisation

Die Spaltenköpfe in Tabelle 1 müssen zeichengenau mit den Einträgen der Spalte `Spaltenname` in Tabelle 2 übereinstimmen. 

### Definiertes Spalten-Set und Gast-Relevanz:
1. **Bezeichnung** (Sichtbar für Gast: `ja`)
2. **Wo?** (Sichtbar für Gast: `ja`)
3. **Relevanz Gast** (Sichtbar für Gast: `ja`)
4. **System** (Sichtbar für Gast: `nein`)
5. **Marke/ Typ** (Sichtbar für Gast: `nein`)
6. **Besonderheit** (Sichtbar für Gast: `nein`)
7. **Quelle Handwerker/ Verkäufer** (Sichtbar für Gast: `nein`)
8. **Details Nutzung [Output]** (Sichtbar für Gast: `ja`)
9. **Details Steuerung** (Sichtbar für Gast: `nein`)
10. **Störung [Input]** (Sichtbar für Gast: `ja`)
11. **Störung Status** (Sichtbar für Gast: `ja` | Erwartet: `[aktiv, OK]`)
12. **Wartung** (Sichtbar für Gast: `nein`)
13. **Vorsorge** (Sichtbar für Gast: `nein`)
14. **Ersatzteile** (Sichtbar für Gast: `nein`)
15. **Ersatzteil Quelle** (Sichtbar für Gast: `nein`)
16. **Ersatzteil Lagerort** (Sichtbar für Gast: `nein`)
17. **Details zur Vorsorge** (Sichtbar für Gast: `nein`)
18. **Details zur Wartung** (Sichtbar für Gast: `nein`)
19. **Wartung erfolgt** (Sichtbar für Gast: `nein`)
20. **Schlüssel (HW, SW)** (Sichtbar für Gast: `nein`)
21. **Schlüssel Gast [Output]** (Sichtbar für Gast: `ja`)
22. **Dokumente/ Link zur Anleitung [Output]** (Sichtbar für Gast: `ja`)
23. **Kontakt [Output]** (Sichtbar für Gast: `nein`)
24. **Kosten [Output]** (Sichtbar für Gast: `nein`)
25. **Feedback [Input]** (Sichtbar für Gast: `nein`)
26. **Feedback Status** (Sichtbar für Gast: `nein` | Erwartet: `[offen, Nein, OK]`)
27. **Keine Information** (Sichtbar für Gast: `nein` — Erfassungskanal für Wissenslücken)
28. **Keine Information Status** (Sichtbar für Gast: `nein` | Erwartet: `[offen, Nein, OK]`)

---

## 4. HMI-Schnittstelle & Use Cases

Die Benutzeroberfläche (Streamlit HMI) verzichtet auf administrative Schaltflächen (da der Admin direkt in Excel arbeitet) und unterscheidet strikt zwei Rollen:

### Rolle: Gast
* **Use Case 1: Ich brauche Hilfe / Frage stellen** -> KI-gestützte Beantwortung. Kontext ist streng zweidimensional gefiltert (nur Zeilen mit `x`, nur Spalten mit `ja`).
* **Use Case 2: Eine Störung melden** -> Freitextfeld. Der Eintrag wird per Append-Logik in die Spalte `Störung [Input]` der gewählten Zeile geschrieben. Der `Störung Status` wird automatisch auf `aktiv` gesetzt.
* **Use Case 3: Feedback geben** -> Freitextfeld. Eintrag wird per Append-Logik in `Feedback [Input]` geschrieben; `Feedback Status` wird auf `offen` gesetzt.
* **Use Case 4: Unbeantwortete Fragen / Keine Information** -> Wenn der Gast eine Information nicht finden konnte, kann er diese hier eintragen. Die App schreibt den Text in die Spalte `Keine Information` und setzt den `Keine Information Status` auf `offen`.

### Rolle: Host
* **Use Case 1: Ich brauche Hilfe / Frage stellen** -> KI-Zugriff auf die *gesamte* Matrix (alle Zeilen, alle Spalten).
* **Use Case 2: Neue Informationen eintragen** -> Ermöglicht das Anhängen von operativen Notizen an bestehende Objekte (Append-Logik).
* **Use Case 3: Störung melden / verwalten** -> Kann Störungen melden oder den Status einsehen.
* **Use Case 4: Feedback erfassen** -> Dokumentation von Feedback im System.
* **Use Case 5: Unbeantwortete Fragen protokollieren** -> Der Host kann Wissenslücken, die im Gespräch mit Gästen auffallen, direkt für den Admin einsteuern.
* **Use Case 6: Bericht generieren** -> Die KI analysiert periodische Daten (z. B. alle aktiven Störungen, offene Feedbacks oder ungelöste Fragen) und erstellt eine strukturierte Zusammenfassung für den Host.

---

## 5. Technische Implementierungs-Vorgaben

* **Dynamische Spalten-Findung:** Im Code dürfen keine statischen Spaltenindizes verwendet werden (z. B. nicht `col=11`). Die App sucht die Spaltennummern beim Start dynamisch über die Namen aus der Header-Zeile.
* **Append-Logik bei Eingaben:** Bestehende Zellinhalte in den Input-Spalten (`Störung`, `Feedback`, `Keine Information`) dürfen *niemals* überschrieben werden. Neue Einträge werden chronologisch mit Zeitstempel und Rolle angehängt (z. B. `\n[05.06.2026 - Gast]: Text`).
* **KI-Sicherheitsnetz (Fallback):** Anfragen werden primär an `gemini-2.5-flash` gestellt. Bei Überlastung oder Quoten-Limits fängt der Code den Fehler ab und schaltet automatisch auf `gemini-2.0-flash` um.
* **Token-Spargang:** Um API-Quoten zu schonen, wird der KI nicht immer die gesamte Tabelle übergeben. Über Dropdowns (z. B. Bereichsauswahl „Wo?“) wird der Tabellen-Kontext für die KI vorab gefiltert.
