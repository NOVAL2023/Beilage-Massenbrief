# Beilage-Massenbrief

## Project Overview

Dieses Projekt erstellt automatisch **Excel-Beilagen** (Übersichten je Kreditor) basierend auf einer Eingabe-Excel-Datei.  
Für **jeden Kreditor** wird in der Ausgabedatei ein eigenes Tabellenblatt generiert.  
Die Beilagen enthalten Forderungspositionen, Summen sowie – falls vorhanden – Begründungstexte für NA14-Codes.

#### How It Works

1. Eine Eingabe-Excel (`mock.xlsx`) enthält alle Kreditoren- und Rechnungsdaten.  

2. Das Skript gruppiert diese Daten nach:  
   - Kreditor-Code (`ithSupplierCode`)  
   - Kreditor-Name (`ithSupplierName`)  

3. Eine **Excel-Vorlage** (`Beilage Verfuegung.xlsx`) liefert das Layout (Seitenränder, Kopf-/Fußzeilen, Spaltenüberschriften).  

4. Für jeden Kreditor wird ein neues Blatt erstellt mit:  
   - Rechnungsdetails (Externe Nr., ER, Betrag, Verfügung, Code, Begründung)  
   - Summenzeile  
   - Optional: NA14-Begründungsblock, wenn entsprechende Einträge existieren  

5. Alle Blätter werden in einer Datei `Beilage_Verfuegung_per_Kreditor.xlsx` gespeichert.  

---

## Programming Language and Libraries

- **Language**: Python  
- **Main Packages**:  
  - `pandas` – Einlesen und Vorverarbeitung der Eingabedaten  
  - `openpyxl` – Erstellen, Formatieren und Schreiben der Excel-Beilagen  
  - `pathlib`, `re`, `sys` – Dateipfade, reguläre Ausdrücke, Fehlerbehandlung  

---

## Data Structure

**Excel File**: `mock.xlsx`  
Wichtige Spalten:  
- `ithSupplierCode` – Kreditor-Nr. (Gruppierungsschlüssel)  
- `ithSupplierName` – Kreditor-Name  
- `ithSupplierCity` – Kreditor-Ort (optional)  
- `ithSupplierExternalNbr1` – externe Rechnungsnummer  
- `ER` – Eingangsrechnungs-Nr.  
- `itlTotalAmount` – Betrag (wird summiert)  
- `itlCostCentreCode1` – Kostenstelle (mapped auf Status: „Anerkannt“, „Bestritten“, etc.)  
- `Code` – Verfügungscode (z. B. „NA14“)  
- `Begründung` – Textfeld für NA14 oder andere Codes  

---

## Code Architecture

1. **Setup and Initialization**  
   - Prüft Existenz von Eingabedatei und Vorlage  
   - Lädt Daten aus `mock.xlsx` mit `pandas`  
   - Normalisiert Strings und konvertiert Beträge zu Zahlen  

2. **Data Processing**  
   - Bildet Liste aller Kreditoren  
   - Filtert Datensätze pro Kreditor  
   - Sortiert nach „C-Nummer“ in `ithSupplierExternalNbr1`  

3. **Excel Rendering**  
   - Kopiert Vorlageblatt aus `Beilage Verfuegung.xlsx`  
   - Füllt Kopfbereich (B4: Kreditor-Nr., B5: Kreditor-Name/Ort)  
   - Schreibt Rechnungszeilen ab Zeile 10  
   - Erstellt Summenzeile (Total)  
   - Fügt NA14-Block (Begründung) ein, falls vorhanden  

4. **Output**  
   - Entfernt das ursprüngliche Vorlagenblatt  
   - Speichert Datei unter `Beilage_Verfuegung_per_Kreditor.xlsx`  

---

## Work Procedure

Die Entwicklung erfolgte iterativ:  
- Aufbau einer Excel-Vorlage mit korrekten Spaltenüberschriften und Formatierungen  
- Ergänzung robuster Fehlerbehandlung (fehlende Dateien, Spaltenprüfungen)  
- Implementierung spezieller Logik für NA14-Begründungen  
- Tests mit Beispiel-Excel-Dateien (`mock.xlsx`, `Beilage Verfuegung.xlsx`)  

---

## Prerequisites

Installiere benötigte Pakete mit pip:  

```bash
pip install pandas openpyxl
```

Klonen des Repositories:  

```bash
git clone https://github.com/NOVAL2023/Beilage-Massenbrief
cd Beilage-Massenbrief
```

---

## Usage

Lege die folgenden Dateien in den Projektordner:  
- `mock.xlsx` (Eingabedaten)  
- `Beilage Verfuegung.xlsx` (Excel-Vorlage)  

Führe das Skript aus:  

```bash
python PythonApplication3.py
```

Das Ergebnis (`Beilage_Verfuegung_per_Kreditor.xlsx`) liegt im Projektordner.  

---

## Authors

- Cristian Noya (Head Finance bei SDAG)
- Noé Peterhans (M&A Analyst bei SDAG)  
