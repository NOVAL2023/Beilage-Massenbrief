# -*- coding: utf-8 -*-
"""
Serienblätter je Kreditor aus mock.xlsx
nutzt Vorlage 'Beilage Verfuegung.xlsx' und erzeugt 'Beilage_Verfuegung_per_Kreditor.xlsx'
"""

import re
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break

# === Basispfade ===
BASE_DIR = Path(r"C:\Users\peno\Beilagebrief\Beilage-Massenbrief")  ### HIER MUSS EIGENER PFAD GEWÄHLT WERDEN ###
INPUT_XLSX    = BASE_DIR / "mock.xlsx"
TEMPLATE_XLSX = BASE_DIR / "Beilage Verfuegung.xlsx"
OUTPUT_XLSX   = BASE_DIR / "Beilage_Verfuegung_per_Kreditor.xlsx"

# === Spalten in mock.xlsx ===
COL_SUP_CODE = "ithSupplierCode"
COL_SUP_NAME = "ithSupplierName"
COL_SUP_CITY = "ithSupplierCity"          # optional
COL_SUP_EXT  = "ithSupplierExternalNbr1"
COL_ER       = "ER"
COL_AMOUNT   = "itlTotalAmount"
COL_CC       = "itlCostCentreCode1"
COL_CODE     = "Code"
COL_REASON   = "Begründung"

# === Welches Register (Sheet) soll gelesen werden === 
SHEET_NAME = "Kontierung"

# === Vorlage-Zellen ===
CELL_SUP_CODE = "B4"   # Kreditor Nr. -> Code
CELL_SUP_NAME = "B5"   # Kreditor -> Name (ggf. Stadt)

# KORREKTE Zeilen basierend auf Vorlage-Analyse
TABLE_START_ROW = 10   # Daten starten in Zeile 10
HEADER_ROW = 8         # Header-Titel sind in Zeile 8
TEMPLATE_ROW = 9       # Diese Zeile muss gelöscht werden
TEMPLATE_NA14_ROWS = [23, 24, 25]  # Standard NA14-Bereich der Vorlage löschen

# Spaltenzuordnung basierend auf der tatsächlichen Vorlage
COLS_TEMPLATE_ORDER = [
    (COL_SUP_EXT, "A"),  # Forderungseingabe
    (COL_ER,      "B"),  # RE-Nr.
    (COL_AMOUNT,  "C"),  # Betrag in CHF
    (COL_CC,      "D"),  # Klasse (Kostenstelle gemappt)
    (COL_CODE,    "E"),  # Verfügung (Code)
    (COL_REASON,  "F"),  # Begründung
    # G = Text (bleibt leer)
]

# Kostenstellen-Legende -> Bezeichnung
COST_CENTER_MAP = {
    "9099100": "Anerkannt",
    "9099200": "Bedingt anerkannt",
    "9099300": "Bestritten",
    "9099400": "Massa",
}

def safe_sheet_name(name: str) -> str:
    if not name or str(name).strip() == "":
        name = "Sheet"
    try:
        name = str(name).encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')
    except:
        name = "Sheet"
    name = re.sub(r'[\\/*?\[\]]+', "_", name)
    name = name.strip()
    return name[:31] or "Sheet"

def setup_page_formatting(ws):
    """Setzt die Seitenformatierung für A4 Querformat mit Fusszeile"""
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    
    ws.page_margins = PageMargins(
        left=0.7, right=0.7, top=0.75, bottom=1.0,
        header=0.3, footer=0.5
    )
    
    ws.oddFooter.center.text = "Seite &P von &N"
    ws.oddFooter.center.size = 10

def set_column_widths(ws):
    """Setzt optimale Spaltenbreiten für A4 Querformat"""
    column_widths = {
        'A': 18, 'B': 12, 'C': 15, 'D': 18, 'E': 12, 'F': 25, 'G': 15
    }
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

def clean_template_rows(ws):
    """Löscht alle störenden Zeilen aus der Vorlage"""
    # Zeile 9: Technische Spaltennamen
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        ws[f"{col}{TEMPLATE_ROW}"].value = None
        # Auch eventuelle Rahmen/Formatierung entfernen
        ws[f"{col}{TEMPLATE_ROW}"].border = Border()
        ws[f"{col}{TEMPLATE_ROW}"].fill = PatternFill()
    
    # Zeilen 23-25: Standard NA14-Bereich aus Vorlage löschen
    for row in TEMPLATE_NA14_ROWS:
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
            ws[f"{col}{row}"].value = None
            ws[f"{col}{row}"].border = Border()
            ws[f"{col}{row}"].fill = PatternFill()

def set_and_format_headers(ws, header_row):
    """Setzt die Header-Titel und formatiert sie - STRIKT nur A-F"""
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    header_font = Font(bold=True)
    bottom_border = Border(bottom=Side(style='thin'))
    
    # Nur Spalten A-F formatieren (G hat keinen Titel mehr, H nie)
    for col in ['A', 'B', 'C', 'D', 'E', 'F']:
        cell = ws[f"{col}{header_row}"]
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = bottom_border
    
    # Sicherstellen, dass G und H KEINE Formatierung haben
    for col in ['G', 'H']:
        cell = ws[f"{col}{header_row}"]
        cell.border = Border()  # Explizit leeren Rahmen setzen
        cell.fill = PatternFill()  # Explizit leere Füllung setzen

def apply_cell_formatting(ws, row, col_letter, value, is_total_row=False):
    """Wendet einheitliche Formatierung auf Zellen an"""
    cell = ws[f"{col_letter}{row}"]
    cell.value = value
    
    if col_letter == "A":
        cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    elif col_letter in ["B", "C", "D"]:
        cell.alignment = Alignment(horizontal='right', vertical='center', indent=1)
        if col_letter == "C" and isinstance(value, (int, float)) and value != 0:
            cell.number_format = "#,##0"
    elif col_letter == "E":
        cell.alignment = Alignment(horizontal='center', vertical='center')
    elif col_letter == "F":
        cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True, indent=1)
    else:
        cell.alignment = Alignment(horizontal='left', vertical='center')
    
    if is_total_row:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

def map_cost_center(value: str) -> str:
    """Gibt Bezeichnung aus Legende zurück; wenn unbekannt, Originalcode."""
    code = (value or "").strip()
    code_digits = "".join(ch for ch in code if ch.isdigit())
    return COST_CENTER_MAP.get(code_digits or code, COST_CENTER_MAP.get(code, code))

def calculate_optimal_na14_position(total_row_idx):
    """Berechnet die optimale Position für NA14 mit garantiertem 3-Zeilen-Abstand"""
    # IMMER 3 Zeilen Abstand nach Total-Zeile
    na14_start_row = total_row_idx + 4
    
    # Prüfen ob Seitenumbruch sinnvoll ist (ab Zeile 30)
    if na14_start_row > 30:
        print(f"NA14 wird in Zeile {na14_start_row} platziert (möglicherweise auf Seite 2)")
    
    return na14_start_row

def sort_by_c_number(df, col_name):
    """Sortiert DataFrame nach C-Nummern"""
    def extract_c_number(value):
        try:
            str_val = str(value).strip()
            if '_' in str_val:
                return int(str_val.split('_')[-1])
            elif str_val.startswith('C') and len(str_val) > 1:
                numbers = re.findall(r'\d+', str_val)
                if numbers:
                    return int(numbers[-1])
            return float('inf')
        except:
            return float('inf')
    
    df_sorted = df.copy()
    df_sorted['_sort_key'] = df_sorted[col_name].apply(extract_c_number)
    df_sorted = df_sorted.sort_values('_sort_key').drop('_sort_key', axis=1)
    return df_sorted

def main():
    if not INPUT_XLSX.exists():
        raise FileNotFoundError(f"Eingabedatei fehlt: {INPUT_XLSX}")
    if not TEMPLATE_XLSX.exists():
        raise FileNotFoundError(f"Vorlage fehlt: {TEMPLATE_XLSX}")

    try:
        df = pd.read_excel(INPUT_XLSX, engine="openpyxl", sheet_name=SHEET_NAME)
    except UnicodeDecodeError:
        df = pd.read_excel(INPUT_XLSX, engine="openpyxl", encoding="latin1")

    required = [COL_SUP_CODE, COL_SUP_NAME, COL_SUP_EXT, COL_ER, COL_AMOUNT, COL_CC, COL_CODE, COL_REASON]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Pflichtspalten fehlen in {INPUT_XLSX.name}: {missing}")

    # Normalisieren
    for c in [COL_SUP_CODE, COL_SUP_NAME, COL_SUP_CITY, COL_SUP_EXT, COL_ER, COL_CC, COL_CODE, COL_REASON]:
        if c in df.columns:
            df[c] = df[c].astype(str).fillna("").apply(
                lambda x: x.encode('utf-8', errors='ignore').decode('utf-8', errors='ignore').strip()
            )
    df[COL_AMOUNT] = pd.to_numeric(df[COL_AMOUNT], errors="coerce").fillna(0.0)

    wb = load_workbook(TEMPLATE_XLSX)
    base_ws = wb.active
    base_title = base_ws.title

    sup_cols = [COL_SUP_CODE, COL_SUP_NAME] + ([COL_SUP_CITY] if COL_SUP_CITY in df.columns else [])
    suppliers = (
        df[sup_cols]
        .drop_duplicates(subset=[COL_SUP_CODE])
        .sort_values(by=[COL_SUP_NAME, COL_SUP_CODE])
        .to_dict(orient="records")
    )

    for sup in suppliers:
        code = sup.get(COL_SUP_CODE, "")
        name = sup.get(COL_SUP_NAME, "")
        city = sup.get(COL_SUP_CITY, "") if COL_SUP_CITY in sup else ""

        part = df[df[COL_SUP_CODE] == code].copy()
        part = sort_by_c_number(part, COL_SUP_EXT)

        ws = wb.copy_worksheet(base_ws)
        ws.title = safe_sheet_name(name or code or "Kreditor")

        setup_page_formatting(ws)
        set_column_widths(ws)

        ws[CELL_SUP_CODE] = code
        ws[CELL_SUP_NAME] = f"{name}{(', ' + city) if city else ''}"

        # WICHTIG: Alle störenden Vorlage-Zeilen löschen
        clean_template_rows(ws)
        
        # Header formatieren (strikt nur A-G)
        set_and_format_headers(ws, HEADER_ROW)

        # Datenzeilen
        start_row = TABLE_START_ROW
        total_amount_sheet = float(part[COL_AMOUNT].sum())
        
        for i, (_, row) in enumerate(part.iterrows(), start=0):
            r = start_row + i
            for col_name, col_letter in COLS_TEMPLATE_ORDER:
                val = row.get(col_name, "")
                if col_name == COL_CC:
                    val = map_cost_center(val)
                elif col_name == COL_AMOUNT:
                    val = float(row.get(col_name, 0))
                apply_cell_formatting(ws, r, col_letter, val, is_total_row=False)

        # Total-Zeile (ohne Spalte G zu formatieren)
        total_row_idx = start_row + len(part)
        for col_letter, val in [("A", "Total"), ("B", ""), ("C", total_amount_sheet), 
                               ("D", ""), ("E", ""), ("F", "")]:
            apply_cell_formatting(ws, total_row_idx, col_letter, val, is_total_row=True)
        
        # Spalte G in Total-Zeile explizit NICHT formatieren
        ws[f"G{total_row_idx}"].fill = PatternFill()  # Keine Füllung
        ws[f"G{total_row_idx}"].border = Border()     # Kein Rahmen

        
        # NA15-Register erzeugen

        try:
            # Alle NA15-Zeilen dataset-weit
            df_na15 = df[df[COL_CODE].str.upper() == "NA15"].copy()

            if not df_na15.empty:
                # Optional sortieren: nach Kreditor-Name, dann ER
                sort_cols = []
                if COL_SUP_NAME in df_na15.columns:
                    sort_cols.append(COL_SUP_NAME)
                if COL_ER in df_na15.columns:
                    sort_cols.append(COL_ER)
                if sort_cols:
                    df_na15 = df_na15.sort_values(sort_cols)

                ws_na = wb.create_sheet(title="NA15_Begründungen")

                # Seitenlayout & Spaltenbreiten (du kannst deine Helfer wiederverwenden/variieren)
                setup_page_formatting(ws_na)
                # Eigene schmale Breiten: Name breiter, Begründung am breitesten
                widths = {'A': 35, 'B': 14, 'C': 90}
                for col, w in widths.items():
                    ws_na.column_dimensions[col].width = w

                # Header
                header_row = 1
                headers = [("A", "Kreditor"), ("B", "ER Nr."), ("C", "Begründung")]
                header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                header_font = Font(bold=True)
                bottom_border = Border(bottom=Side(style='thin'))
                for col, title in headers:
                    cell = ws_na[f"{col}{header_row}"]
                    cell.value = title
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = bottom_border

                # Daten schreiben
                row_idx = header_row + 1
                for _, rec in df_na15.iterrows():
                    # Kreditor-Name (Fallback auf Code)
                    kred_name = str(rec.get(COL_SUP_NAME, "")).strip()
                    if not kred_name:
                        kred_name = str(rec.get(COL_SUP_CODE, "")).strip()

                    er_val = str(rec.get(COL_ER, "")).strip()
                    begr = str(rec.get(COL_REASON, "")).strip()

                    ws_na[f"A{row_idx}"] = kred_name
                    ws_na[f"A{row_idx}"].alignment = Alignment(horizontal='left', vertical='top')

                    ws_na[f"B{row_idx}"] = er_val
                    ws_na[f"B{row_idx}"].alignment = Alignment(horizontal='center', vertical='top')

                    ws_na[f"C{row_idx}"] = begr
                    ws_na[f"C{row_idx}"].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

                    # Optionale dynamische Zeilenhöhe für lange Begründungen
                    if begr:
                        est_lines = max(1, len(begr) // 80 + begr.count("\n") + 1)
                        ws_na.row_dimensions[row_idx].height = min(est_lines * 15, 180)

                    row_idx += 1

                print(f"✓ NA15-Reiter mit {len(df_na15)} Einträgen erzeugt")

            else:
                print("Info: Keine NA15-Einträge gefunden – kein NA15-Register angelegt.")

        except Exception as e:
            print(f"Warnung: NA15-Register konnte nicht erstellt werden: {e}")


    wb.remove(wb[base_title])
    wb.save(OUTPUT_XLSX)
    print(f"Fertig. Datei erstellt:\n{OUTPUT_XLSX}")

if __name__ == "__main__":

    main()
