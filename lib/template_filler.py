# ============================================================
# template_filler.py
# Erzeugt eine neue Systembewertung (V11) aus einem Daten-Dict - 1.0
#
# Gegenstueck zu den extract_*-Funktionen in sysbew_common.py: dort
# werden Werte AUS einem ausgefuellten Dokument GELESEN, hier werden
# Werte IN das Leer-Template (assets/templates_docx/
# Systembewertung_V11_leer.docx) GESCHRIEBEN.
#
# Wichtiger Unterschied zur Extraktion: die Extraktion muss mit vielen
# unterschiedlichen, bereits ausgefuellten Realdokumenten (V7/V8/V10/
# V11, verschiedene Autoren) zurechtkommen und sucht Felder daher per
# Text-Heuristik. Der Filler bearbeitet dagegen IMMER dieselbe, fest
# bekannte Leer-Vorlage - deshalb sind hier feste Tabellen-/Zeilen-/
# Zellen-Indizes verwendet (robuster und einfacher als Text-Suche),
# ermittelt durch direkte Inspektion von
# assets/templates_docx/Systembewertung_V11_leer.docx.
#
# BEKANNTE EINSCHRAENKUNG (siehe auch README.md): folgende Bereiche
# des Templates werden NICHT automatisch befuellt, weil sich aus den
# in der Master-Excel gespeicherten Endergebnissen nicht eindeutig
# der zugrunde liegende Entscheidungsweg rekonstruieren laesst (mehrere
# Antwortpfade koennen zum selben Endergebnis fuehren - ein Raten
# waere in einem GxP-Dokument nicht vertretbar):
#   - Kapitel 3 (Detailfestlegung Klasse 1a/1b bei Globales CS)
#   - Kapitel 5-9 (Entscheidungsbaum Geraetekategorie/CS-Typ/ERES-Typ/
#     KI, inkl. der zugehoerigen Ja/Nein-Antworten)
#   - Kapitel 8 (Testtiefe-Matrix)
# Diese Kapitel muessen nach der automatischen Erstellung manuell in
# Word ergaenzt werden. Alles, was in der Zusammenfassungstabelle
# (Kapitel 2) und auf dem Deckblatt steht, wird dagegen vollstaendig
# befuellt.
# ============================================================

import os
import re
from datetime import date

from docx import Document

_NS = {
    "w":   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
}

DEFAULT_TEMPLATE_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "assets", "templates_docx", "Systembewertung_V11_leer.docx",
)

# ============================================================
# Low-Level: Checkbox-Zustand schreiben (inkl. sichtbarem Glyph)
# ============================================================
def _w14(tag):
    return f"{{{_NS['w14']}}}{tag}"

def _w(tag):
    return f"{{{_NS['w']}}}{tag}"

def get_checkbox_sdts_from_cell(cell):
    """Alle echten Checkbox-SDTs einer Zelle, in Dokumentreihenfolge
    (dieselbe Reihenfolge, die sysbew_common.get_checkboxes_from_cell
    beim Lesen verwendet - wichtig, damit cb_idx aus CHECKBOX_MAPPING_V11
    zur richtigen Checkbox passt)."""
    result = []
    for sdt in cell._element.findall(".//w:sdt", _NS):
        checkbox = sdt.find(".//w14:checkbox", _NS)
        if checkbox is not None:
            result.append(sdt)
    return result

def set_checkbox_state_in_sdt(sdt, checked):
    """Setzt sowohl den logischen Zustand (w14:checked/@w14:val) als
    auch das sichtbare Glyphen-Zeichen (☐/☒) im sdtContent - beides
    muss uebereinstimmen, sonst zeigt Word beim Oeffnen den falschen
    Haken an, bis man die Checkbox einmal anklickt."""
    checkbox = sdt.find(".//w14:checkbox", _NS)
    if checkbox is None:
        return False
    checked_el = checkbox.find("w14:checked", _NS)
    if checked_el is None:
        checked_el = checkbox.makeelement(_w14("checked"), {})
        checkbox.insert(0, checked_el)
    checked_el.set(_w14("val"), "1" if checked else "0")

    state_tag = "checkedState" if checked else "uncheckedState"
    state_el = checkbox.find(f"w14:{state_tag}", _NS)
    glyph = None
    if state_el is not None:
        code_hex = state_el.get(_w14("val"))
        if code_hex:
            glyph = chr(int(code_hex, 16))

    if glyph:
        sdt_content = sdt.find("w:sdtContent", _NS)
        if sdt_content is not None:
            t = sdt_content.find(".//w:t", _NS)
            if t is not None:
                t.text = glyph
    return True

def set_checkboxes_in_cell(cell, states):
    """`states`: dict {checkbox_index: bool}. Checkbox-Indizes, die
    nicht im dict stehen, bleiben unveraendert (Default des leeren
    Templates = nicht angekreuzt)."""
    sdts = get_checkbox_sdts_from_cell(cell)
    for idx, checked in states.items():
        if idx < len(sdts):
            set_checkbox_state_in_sdt(sdts[idx], bool(checked))

def set_cell_text(cell, text):
    """Ersetzt den gesamten Zellinhalt durch `text` (mehrzeilig via
    "\\n" wird von python-docx automatisch als Zeilenumbruch
    dargestellt). Leerer/None-Text wird als leere Zelle geschrieben
    (Platzhalter wird also auch dann entfernt, wenn keine Daten da
    sind - bewusst, damit kein "<<...>>"-Platzhaltertext im Enddokument
    uebrig bleibt)."""
    cell.text = text or ""

# ============================================================
# Checkbox-Gruppen: generisch ueber CHECKBOX_MAPPING_V11 (identische
# Struktur zu sysbew_common.extract_checkboxes_formularfelder, nur in
# Schreibrichtung)
# ============================================================
def fill_checkboxes_formularfelder(doc, checkbox_mapping, data):
    for table_idx, table_mapping in checkbox_mapping.items():
        table = doc.tables[table_idx]
        rows = list(table.rows)
        for row_idx, row_mapping in table_mapping.items():
            cells = rows[row_idx].cells
            for cell_idx, cell_mapping in row_mapping.items():
                cell = cells[cell_idx]
                states = {
                    cb_idx: data.get(col) == "r"
                    for cb_idx, col in cell_mapping.items()
                }
                set_checkboxes_in_cell(cell, states)

# ============================================================
# Deckblatt: Rollen/Namen
# ============================================================
def fill_deckblatt_rollen(doc, data):
    """Tabelle 0 (Unterschriftentabelle): Zeile 0/2/4/6/8/10, Spalte 1
    (siehe ROLLEN_SPALTEN-Reihenfolge in sysbew_common.py)."""
    table = doc.tables[0]
    zeilen_je_rolle = {
        "Ersteller": 0, "SI/PL": 2, "TSO": 4,
        "BSO": 6, "BQR": 8, "CSQ": 10,
    }
    for spalte, row_idx in zeilen_je_rolle.items():
        wert = data.get(spalte, "")
        if wert:
            set_cell_text(table.rows[row_idx].cells[1], wert)
    # SME: nur befuellen, wenn nicht ohnehin schon mit SI/PL identisch
    # (kombinierte Rolle "Projektleiter/SME" - Name steht dann bereits
    # in Zeile 2 bei SI/PL). Ansonsten SME wie eigene Rolle behandeln;
    # das Template hat dafuer keine eigene Zeile, SME wird deshalb bei
    # abweichendem Namen an die SI/PL-Zeile angehaengt.
    sme = data.get("SME", "")
    si_pl = data.get("SI/PL", "")
    if sme and sme != si_pl:
        bestehend = table.rows[2].cells[1].text.strip()
        neu = f"{bestehend}\n{sme}" if bestehend else sme
        set_cell_text(table.rows[2].cells[1], neu)

# ============================================================
# Klassifizierung (Tabelle 3, Zeile 1)
# ============================================================
def fill_klassifizierung(doc, data):
    row = doc.tables[3].rows[1]
    set_checkboxes_in_cell(row.cells[0], {0: data.get("KLASS_Lokal") == "r"})
    set_checkboxes_in_cell(row.cells[1], {
        0: data.get("KLASS_Multisite") == "r",
        1: data.get("KLASS_Multisite_NurLokal") == "r",
        2: data.get("KLASS_Multisite_LokalGlobal") == "r",
    })
    set_checkboxes_in_cell(row.cells[2], {
        0: data.get("KLASS_Global") == "r",
        1: data.get("KLASS_Global_1a") == "r",
        2: data.get("KLASS_Global_1b") == "r",
        3: data.get("KLASS_Global_2") == "r",
        4: data.get("KLASS_Global_3") == "r",
    })
    set_checkboxes_in_cell(row.cells[3], {0: data.get("KLASS_OhneCS") == "r"})

# ============================================================
# Kapitel 1 (Tabelle 2): Neuerstellung, Equipment-Name,
# Kurzbeschreibung, Systemtyp-Zugang, Einsatzbereich, Hersteller
# ============================================================
def fill_kapitel1(doc, data):
    table = doc.tables[2]

    set_checkboxes_in_cell(table.rows[1].cells[0], {
        0: data.get("Neuerstellung") == "r",
        1: data.get("Revisioniert") == "r",
    })

    if data.get("AS/BDIS-Name"):
        set_cell_text(table.rows[3].cells[1], data["AS/BDIS-Name"])

    if data.get("Kurzbeschreibung"):
        set_cell_text(table.rows[4].cells[1], data["Kurzbeschreibung"])

    set_checkboxes_in_cell(table.rows[5].cells[1], {
        0: data.get("Offen") == "r",
        1: data.get("Geschlossen") == "r",
        2: data.get("NA") == "r",
    })

    einsatzbereich = data.get("Betrieb", "")
    if data.get("Gebaeude"):
        einsatzbereich = f"{einsatzbereich}/{data['Gebaeude']}" if einsatzbereich else data["Gebaeude"]
    if einsatzbereich:
        set_cell_text(table.rows[6].cells[1], einsatzbereich)

    hersteller = data.get("Hersteller", "")
    if data.get("SW-Hersteller") and data["SW-Hersteller"] != hersteller:
        hersteller = f"{hersteller} / {data['SW-Hersteller']}" if hersteller else data["SW-Hersteller"]
    if hersteller:
        set_cell_text(table.rows[7].cells[1], hersteller)

# ============================================================
# Business Kritikalitaet (Tabelle 4) + GxP-Relevanz (Tabelle 5)
# - Detailfragen, deren Ergebnis in der Zusammenfassungstabelle
# (Kapitel 2) noch einmal auftaucht (dort ueber
# fill_checkboxes_formularfelder/CHECKBOX_MAPPING_V11 abgedeckt).
# ============================================================
def fill_business_kritikalitaet(doc, data):
    table = doc.tables[4]
    set_checkboxes_in_cell(table.rows[1].cells[0], {0: data.get("BCkritisch") == "r"})
    set_checkboxes_in_cell(table.rows[2].cells[0], {0: data.get("BCunkritisch") == "r"})

def fill_gxp_relevanz(doc, data):
    table = doc.tables[5]
    set_checkboxes_in_cell(table.rows[1].cells[0], {0: data.get("GxP_Relevan_JA") == "r"})
    set_checkboxes_in_cell(table.rows[2].cells[0], {0: data.get("GxP_Relevan_NEIN") == "r"})

# ============================================================
# Zusammenfassungstabelle (Tabelle 6): MLCS-ID, Anlage,
# Schnittstelle/UeberlagerteMLCS, Besonderheiten
# (die Checkbox-Spalten r6/r8 laufen ueber CHECKBOX_MAPPING_V11)
# ============================================================
def fill_zusammenfassung_text(doc, data):
    table = doc.tables[6]
    if data.get("MLCSID"):
        set_cell_text(table.rows[1].cells[3], data["MLCSID"])
    if data.get("Anlage"):
        set_cell_text(table.rows[2].cells[3], data["Anlage"])

    schnittstelle = data.get("Schnittstelle") or data.get("UeberlagerteMLCS") or ""
    if data.get("Schnittstelle") and data.get("UeberlagerteMLCS") and \
            data["Schnittstelle"] != data["UeberlagerteMLCS"]:
        schnittstelle = f"{data['Schnittstelle']} / {data['UeberlagerteMLCS']}"
    if schnittstelle:
        set_cell_text(table.rows[3].cells[3], schnittstelle)

    if data.get("Besonderheiten"):
        set_cell_text(table.rows[9].cells[1], data["Besonderheiten"])

# ============================================================
# GxP-Risikoklassifizierung im Detail (Tabelle 8) - Duplikat der
# GxP-Kritikalitaet aus der Zusammenfassungstabelle plus Begruendung
# ============================================================
def fill_gxp_risikoklassifizierung(doc, data):
    table = doc.tables[8]
    set_checkboxes_in_cell(table.rows[1].cells[0], {
        0: data.get("GxP-C") == "r",
        1: data.get("GxP-M") == "r",
        2: data.get("GxP-m2") == "r",
        3: data.get("GxP-NA") == "r",
    })

    teile = []
    for label, feld in [
        ("Produktqualität", "GxP_Produktqualitaet"),
        ("Patientensicherheit", "GxP_Patientensicherheit"),
        ("Datenintegrität", "GxP_Datenintegritaet"),
    ]:
        wert = data.get(feld, "")
        if wert:
            teile.append(f"{label}: {wert}" if not wert.lower().startswith(label.lower()) else wert)
    begruendung = " ".join(teile).strip()
    if begruendung:
        set_cell_text(
            table.rows[2].cells[0],
            f"Begründung der GxP Risikoklassifizierung: {begruendung}",
        )

# ============================================================
# Historie (Tabelle 1, Zeile 1) - fuer eine NEUE Systembewertung
# immer Version 1.0 mit heutigem Datum
# ============================================================
def fill_historie(doc, data):
    table = doc.tables[1]
    version = data.get("Version_Historie") or "1.0"
    datum = data.get("Datum") or date.today().strftime("%d.%m.%Y")
    set_cell_text(table.rows[1].cells[0], f"{version} / {datum}")

    grund = data.get("Historie") or (
        "Neuerstellung" if data.get("Neuerstellung") == "r" else ""
    )
    if grund:
        set_cell_text(table.rows[1].cells[1], grund)

# ============================================================
# Beschreibungstabelle (Tabelle 16)
# ============================================================
_BESCHREIBUNG_ZEILEN = {
    0: "Prozessbeschreibung",
    1: "Daten",
    2: "Parameter",
    3: "Alarme (GxP-relevant)",
    4: "Chargenprotokoll",
    5: "Audit Trail (AT)",
    6: "Benutzer-verwaltung?",
    7: "Schnittstellen mit PLS",
    9: "Angeschlossenes Equipment",
    10: "Sonstiges",
    11: "KI Bewertung",
}

def fill_beschreibungstabelle(doc, data):
    table = doc.tables[16]
    for row_idx, excel_col in _BESCHREIBUNG_ZEILEN.items():
        wert = data.get(excel_col)
        if wert:
            set_cell_text(table.rows[row_idx].cells[1], wert)

# ============================================================
# HAUPTFUNKTION
# ============================================================
def fill_template(data, template_path=None, output_path=None):
    """Erzeugt aus `data` (dieselbe Feldstruktur wie EXCEL_COLUMNS/
    Master-Excel) eine neue Systembewertung auf Basis des V11-Leer-
    Templates. Gibt den docx.Document zurueck; speichert zusaetzlich
    unter `output_path`, falls angegeben.

    Importiert CHECKBOX_MAPPING_V11 hier (statt am Modulanfang), um
    einen Zirkelimport mit word_parser_v11.py (das umgekehrt aus
    sysbew_common importiert, nicht aus diesem Modul) zu vermeiden -
    ist unproblematisch, da fill_template() ohnehin erst zur Laufzeit
    aufgerufen wird.
    """
    from word_parser_v11 import CHECKBOX_MAPPING_V11

    template_path = template_path or DEFAULT_TEMPLATE_PATH
    doc = Document(template_path)

    fill_deckblatt_rollen(doc, data)
    fill_historie(doc, data)
    fill_kapitel1(doc, data)
    fill_klassifizierung(doc, data)
    fill_business_kritikalitaet(doc, data)
    fill_gxp_relevanz(doc, data)
    fill_zusammenfassung_text(doc, data)
    fill_checkboxes_formularfelder(doc, CHECKBOX_MAPPING_V11, data)
    fill_gxp_risikoklassifizierung(doc, data)
    fill_beschreibungstabelle(doc, data)

    if output_path:
        doc.save(output_path)
    return doc
