# ============================================================
# db_reader.py
# Liest die Master-Excel ("Datenbank") fuer die Web-Anwendung - 1.0
#
# Nur LESEND und OHNE COM (openpyxl reicht dafuer aus und ist deutlich
# schneller/robuster fuer reines Anzeigen/Filtern als eine Excel-COM-
# Instanz zu starten). Das eigentliche Schreiben neuer Zeilen bleibt
# ausschliesslich write_to_master_excel() in sysbew_common.py
# vorbehalten (dort ist der COM-Weg wegen Sensitivity-Label/externer
# Verknuepfung/Kommentare zwingend - siehe Kommentar dort).
# ============================================================

import datetime
import openpyxl

from sysbew_common import (
    MASTER_EXCEL_PFAD,
    MASTER_SHEET_NAME,
    MASTER_TABELLE_NAME,
    MASTER_SPALTEN_MAPPING,
)

def _zellwert_zu_text(wert):
    """Normalisiert einen rohen openpyxl-Zellwert auf str/None - alle
    Weiterverarbeitung (Web-Editor, template_filler, masterform_import)
    erwartet Strings bzw. None, Excel liefert bei rein numerisch
    formatierten Spalten (z.B. MLCSID, SAP-Nummer, Raum) aber int/
    float zurueck. Ohne diese Normalisierung bricht z.B. python-docx
    beim Befuellen einer Tabellenzelle mit einem int hart ab (siehe
    template_filler.set_cell_text -> Run.text erwartet str)."""
    if wert is None or isinstance(wert, str):
        return wert
    if isinstance(wert, float) and wert == int(wert):
        return str(int(wert))
    if isinstance(wert, (datetime.date, datetime.datetime)):
        return wert.strftime("%d.%m.%Y")
    return str(wert)

# Spalten, die in der Startseiten-Uebersicht/Filterung angezeigt bzw.
# durchsucht werden. Bewusst eine Teilmenge (nicht alle 140+ Spalten) -
# der volle Datensatz wird erst beim Oeffnen/Uebernehmen einer Zeile
# geladen.
UEBERSICHT_SPALTEN = [
    "MLCSID", "AS/BDIS-Name", "Dok. -Nr.", "Version",
    "Betrieb", "Gebaeude", "Anlage", "Erkannte_Version",
    "Ersteller", "SI/PL",
]

def _reverse_mapping():
    return {master: skript for skript, master in MASTER_SPALTEN_MAPPING.items()}

def read_master_rows():
    """Liest alle Datenzeilen der Excel-Tabelle `Tabelle1` (nur die
    tatsaechliche Tabellenausdehnung, nicht das ganze Sheet). Gibt
    (header, rows) zurueck - `header` als Liste der Skript-Feldnamen
    (Master-Spalten-Mapping bereits umgekehrt angewandt), `rows` als
    Liste von Dicts {Feldname: Wert}. Eine interne Zeilennummer wird
    unter dem Schluessel "_zeile" mitgefuehrt (1-basiert, Excel-Zeile -
    fuer Rueckverfolgung/Debugging, wird der Anwendung selbst aber
    nicht angezeigt).
    """
    # read_only=True liefert eine ReadOnlyWorksheet ohne .tables-Attribut,
    # daher hier ohne read_only (fuer die ueberschaubare Groesse der
    # Master-Excel unproblematisch).
    wb = openpyxl.load_workbook(MASTER_EXCEL_PFAD, data_only=True)
    try:
        ws = wb[MASTER_SHEET_NAME]
        tabelle = ws.tables.get(MASTER_TABELLE_NAME)
        if tabelle is None:
            raise ValueError(
                f"Excel-Tabelle '{MASTER_TABELLE_NAME}' nicht auf Sheet "
                f"'{MASTER_SHEET_NAME}' gefunden."
            )
        zellen = ws[tabelle.ref]
        header_zeile = zellen[0]
        rueck_mapping = _reverse_mapping()
        header = [rueck_mapping.get(c.value, c.value) for c in header_zeile]

        rows = []
        for excel_row_idx, zeile in enumerate(zellen[1:], start=2):
            eintrag = {}
            leer = True
            for name, zelle in zip(header, zeile):
                if not name:
                    continue
                eintrag[name] = _zellwert_zu_text(zelle.value)
                if zelle.value not in (None, ""):
                    leer = False
            if leer:
                continue
            eintrag["_zeile"] = excel_row_idx
            rows.append(eintrag)
        return header, rows
    finally:
        wb.close()

def filter_rows(rows, suchtext=""):
    """Einfache Volltextsuche ueber die UEBERSICHT_SPALTEN (case-
    insensitiv, Teilstring). Leerer Suchtext -> alle Zeilen."""
    suchtext = (suchtext or "").strip().lower()
    if not suchtext:
        return rows
    treffer = []
    for row in rows:
        for spalte in UEBERSICHT_SPALTEN:
            wert = row.get(spalte)
            if wert and suchtext in str(wert).lower():
                treffer.append(row)
                break
    return treffer

def get_row_by_index(rows, zeile):
    """Findet die Zeile mit interner Excel-Zeilennummer `zeile`
    (siehe "_zeile" in read_master_rows())."""
    for row in rows:
        if row.get("_zeile") == zeile:
            return row
    return None
