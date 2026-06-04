# ☀️ System- & Architekturspezifikation: Villa Avatar

Dieses Dokument dient als unumstößliche Basis („Single Source of Truth“) für die funktionale und strukturelle Entwicklung der Applikation „Villa Avatar“. Alle weiteren Code-Generierungen und Feature-Erweiterungen müssen sich an diesen Definitionen orientieren.

---

## 1. Grundlegende Systemarchitektur
Das System besteht aus einer Streamlit-Benutzeroberfläche (HMI), einem KI-Core (Google Gemini) und einer zeilenorientierten Wissensbasis (Google Drive / Excel).

* **Rollenbasiert:** Das System unterscheidet strikt zwischen drei Nutzertypen: `Gast`, `Host` und `Admin`.
* **Wissensbasis-Filterung:** Datenflüsse werden nutzerspezifisch gefiltert. Gäste haben einen eingeschränkten Lesezugriff (User-Zugriffsfilter über die Spalte „Relevanz Gast“), während Hosts und Admins uneingeschränkten Zugriff auf alle Zeilen haben.
* **Zustandsorientiert (State Machine):** Die HMI folgt einer festen Sequenz von der Rollenwahl über die Intent-Ermittlung bis zum Daten-Datenbank-Workflow.

---

## 2. Die funktionale HMI- & Dialogmatrix
Basierend auf der offiziellen Systemarchitektur gliedern sich die sechs Use Cases wie folgt:

| Use Case | Berechtigte User | UI-Button Text | KI-Systemfrage (Prompt) | Drop-down Filter zur Eingabeverbesserung | Datenbank-Nutzung |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Hilfe** | Gast, Host, Admin | `Ich brauche Hilfe.` | „Wobei kann ich dir helfen?“ | Ausstattung innen, Ausstattung außen, In der Nähe | **Output** (Reines Lesen) |
| **Störung** | Gast, Host, Admin | `Es gibt eine Störung.` | „Was ist passiert?“ | Ausstattung innen, Ausstattung außen, In der Nähe | **Input** (Schreiben in Zeile) |
| **Feedback** | Gast, Host, Admin | `Ich möchte Feedback geben.` | „Welches Feedback hast du?“ | Ausstattung innen, Ausstattung außen, In der Nähe | **Input** (Schreiben in Zeile) |
| **Bericht** | Host, Admin | `Ich benötige einen Bericht.` | „Nenne mir bitte den Zeitraum und das Thema.“ | *Keine Drop-downs* | **Output** (Generieren/Lesen) |
| **Information** | Host, Admin | `Ich habe neue Informationen.` | „Gern nehme ich deine Informationen auf und ordne sie in meiner Wissensbasis zu.“ | Ausstattung innen, Ausstattung außen, In der Nähe | **Input** (Dynamische Stammdaten) |
| **Anpassung** | Admin | `Ich möchte eine Änderung an der Wissensbasis vornehmen.` | „Beschreibe deine Änderung so genau wie möglich.“ | Ausstattung innen, Ausstattung außen, In der Nähe | **Änderung** (Strukturell / Neue Zeile) |

---

## 3. Der standardisierte Dialog- & Darstellungsfluss
Jeder Use Case durchläuft zwingend die exakte, vertikale Reihenfolge der Darstellung im User Interface:

1. **Nutzer-Auswahl:** Gast, Host oder Admin via Drop-down wählen.
2. **Intent-Buttons:** Dynamische Anzeige der erlaubten Buttons (laut Matrix).
3. **System-Frage:** Villa Avatar stellt die exakte, use-case-spezifische Frage.
4. **Kontext-Filter:** Anzeige der Drop-downs (Innen, Außen, Nähe) zur Eingabeverbesserung.
5. **User-Eingabe:** Textfeld / Spracheingabe durch den Anwender.
6. **KI-Verarbeitung:** Verknüpfung von User-Eingabe + gefiltertem Tabellenkontext.
7. **DB-Nutzung:** Ausführung des Workflows: Daten ausgeben (Output), anhängen (Input) oder Struktur ändern.

---

## 4. Daten-Workflows & Integritätsregeln („Mini-Workflows“)

### A. Die Lese-Workflows (Output)
* **Hilfe:** Filtert die Excel-Tabelle. Für *Gäste* werden nur Zeilen geladen, bei denen Spalte C (`Relevanz Gast`) ein `X` enthält. Die KI liest diese Daten und antwortet smartphone-optimiert.
* **Bericht:** Erstellt für Hosts/Admins Zusammenfassungen über bestehende Einträge (z.B. historische Wartungen oder aufgelaufene Störungen) innerhalb eines definierten Zeitraums.

### B. Die Schreib-Workflows (Input)
* **Störung & Feedback:** Suchen in Spalte A nach dem im Drop-down gewählten Objekt. Der neue Text wird in die Spalte `Störung` (Spalte J) bzw. `Feedback` (Spalte X) mit einem Zeilenumbruch (`\n`) angehängt. Der Status (Spalte K bzw. Y) wird automatisch mit einem Zeitstempel, der Rolle und dem Initialstatus (`aktiv` / `offen`) versehen.
* **Information:** Nutzt die KI-Spaltenerkennung. Der Text wird analysiert, die logisch richtige Zielspalte (z.B. *Marke/Typ*, *Ersatzteile* oder *Kontakt*) aus dem vordefinierten Spalten-Mapping ermittelt und der Wert in der Zeile des ausgewählten Objekts ergänzt.

### C. Der Struktur-Workflow (Änderung)
* **Anpassung (Nur Admin):** Erlaubt das Hinzufügen komplett neuer Zeilen (Objekte) in die Excel-Tabelle. Das Objekt wird in Spalte A eingetragen, die gewählte Kategorie in Spalte B, und Spalte C erhält standardmäßig ein `X`, damit das neue Objekt sofort für Gäste sichtbar ist.

### D. Visuelle Revisions-Queue (Farbliche Kennzeichnung)
* **Qualitätssicherung:** Um neu eingegangene Daten für den Administrator sofort sichtbar zu machen, werden alle automatischen Schreibprozesse (neuer Text in Störung/Feedback, neue Stammdaten-Infos sowie komplett neu angelegte Struktur-Zeilen) in **blauer Schriftfarbe (Hex: 0000FF)** in die Excel-Datei geschrieben.
* **Review-Prozess:** Der Administrator nutzt diese farbliche Kennzeichnung als „Posteingang“. Nach redaktioneller Prüfung, Korrektur oder Freigabe der Zelle im Excel-Sheet setzt der Administrator die Schriftfarbe manuell wieder auf die Standardfarbe (Schwarz) zurück, um die Übersichtlichkeit zu wahren.
