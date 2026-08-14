# ============================================================
# sysbew_common.py
# Gemeinsame Basis der Systembewertungs-Extraktoren - 1.0
#
# Enthält ausschließlich Code, der in den drei Vorgänger-Skripten
# (word_parser_v8/v10/v11_formularfelder) Zeile für Zeile identisch
# war: Excel-Spalten, Master-Excel-Konfiguration, alle Hilfs- und
# Extraktionsfunktionen, die nicht vom Template-Aufbau der jeweiligen
# Version abhängen, sowie das Schreiben in die Master-Excel per COM.
#
# Version-spezifischer Code (CHECKBOX_MAPPING, VALIDATION_KATEGORIEN,
# die jeweilige Haupt-Extraktionsfunktion parse_systembewertung_vXX
# sowie alles, was NUR in einer Version existiert wie extract_z_felder
# in V11 oder extract_template_basis_version in V8) bleibt bewusst in
# den jeweiligen Erweiterungs-Modulen (word_parser_v8.py,
# word_parser_v10.py, word_parser_v11.py) - nicht hier.
#
# Wird nicht direkt per Drag & Drop verwendet, sondern von
# word_parser_main.py und den drei Erweiterungs-Modulen importiert.
# ============================================================

import re
import os
import time

# ============================================================
# EXCEL-SPALTEN (identisch in allen drei Vorgänger-Skripten)
# ============================================================
EXCEL_COLUMNS = [
    "MLCSID", "API", "Betrieb", "Gebaeude", "Version", "Dok. -Nr.",
    "AS/BDIS-Name", "Anlage",
    "GxP_Relevan_JA", "GxP_Relevan_NEIN",
    "GxP-C", "GxP-M", "GxP-m2", "GxP-NA",
    "Kurzbeschreibung", "SW-Version / Typ:", "SW-Name:", "SW-Hersteller",
    "Bemerkung1", "Bemerkung2", "Bemerkung3", "Bemerkung4",
    "GxP_Produktqualitaet", "GxP_Patientensicherheit", "GxP_Datenintegritaet",
    "Systemtyp_CIS", "Systemtyp_CE", "Subtyp_PCS", "Subtyp_LCE", "Subtyp_EE",
    "VNAP_S0", "VNAP_S1", "VNAP_S2", "Subtyp_NA",
    "Offen", "Geschlossen", "NA",
    "KAT1", "KAT3", "KAT4", "KAT5", "KATNA",
    "Ersteller", "SME", "SI/PL", "TSO", "BSO", "BQR", "CSQ",
    "Datum", "Version_Historie", "Historie", "Bearbeiter",
    "Hersteller", "Phenix", "SAP",
    "ERESTYP1", "ERESTYP2", "ERESTYP3", "ERESTYP4", "ERESTYPNA",
    "TTIEFEHOCH", "TTIEFEMITTEL", "TTIEFENIEDRIG",
    "Z1S1", "Z2S1", "Z3S1", "Z1S2", "Z1S3", "Z2S2", "Z2S3", "Z3S2", "Z3S3",
    "Lieferantennummer", "UeberlagerteMLCS", "Bedien-SOP", "SOP-Titel",
    "PLSTA", "DokNummerVorQualiPSO", "BE", "Raum", "Schnittstelle",
    "PNK", "BCkritisch", "BCunkritisch",
    "Neuerstellung", "Revisioniert",
    "GKATA", "GKATB1", "GKATB2", "GKATB3", "GKATB",
    "GKATC1", "GKATC2", "GKATC3", "GKATC", "GKATNA",
    "VQ", "NVQ", "QUAL", "VAL",
    "KI1", "KI2", "KI3", "KI4", "KI5", "KI6", "KINA",
    "Steuerung erfolgt über?",
    "Prozessbeschreibung", "Daten", "Parameter",
    "Alarme (GxP-relevant)", "Chargenprotokoll", "Audit Trail (AT)",
    "Benutzer-verwaltung?", "Schnittstellen mit PLS",
    "Angeschlossenes Equipment",
    "Hyperlink",
    "Sonstiges", "KI Bewertung", "Besonderheiten",
    "Erkannte_Version",
]

# ============================================================
# MAPPING: Textfelder → Excel-Spalten (identisch in allen drei
# Vorgänger-Skripten - TEXT_FIELD_MAPPING_V8/V10/V11 waren
# wortgleich)
# ============================================================
TEXT_FIELD_MAPPING = {
    "Bezeichnung des Equipments / Systemname": "AS/BDIS-Name",
    "Bezeichnung des Equipment / Systemname":  "AS/BDIS-Name",
    "Bezeichnung des Equipment/ Systemname":   "AS/BDIS-Name",
    "Bezeichnung des Equipments/ Systemname":  "AS/BDIS-Name",
    "Kurzbeschreibung / Verwendungszweck":     "Kurzbeschreibung",
    "Kurzbeschreibung/Verwendungszweck":       "Kurzbeschreibung",
    "Hersteller / SW-Ersteller / Lieferant":   "_hersteller_raw",
    "Hersteller/ SW-Ersteller/Lieferant":      "_hersteller_raw",
    "Einsatzbereich(e)/-ort(e)":               "_betrieb_raw",
    "Prozessbeschreibung":                     "Prozessbeschreibung",
    "Daten":                                   "Daten",
    "Parameter":                                "Parameter",
    "Alarme (GxP critical)":                   "Alarme (GxP-relevant)",
    "Alarme (GxP-relevant)":                   "Alarme (GxP-relevant)",
    "Chargenprotokoll":                        "Chargenprotokoll",
    "Audit Trail / Audit Trail Review":        "Audit Trail (AT)",
    "Audit Trail (AT)":                        "Audit Trail (AT)",
    "Audit-Trail / Audit-Trail-Review":        "Audit Trail (AT)",
    "Benutzerverwaltung":                      "Benutzer-verwaltung?",
    "Schnittstellen":                          "Schnittstellen mit PLS",
    "Angeschlossene Equipments":               "Angeschlossenes Equipment",
    "Angeschlossenes Equipment":               "Angeschlossenes Equipment",
    "Steuerung erfolgt über":                  "Steuerung erfolgt über?",
    "Bedien-SOP":                              "Bedien-SOP",
    "SOP-Titel":                               "SOP-Titel",
    "Sonstiges":                               "Sonstiges",
    "KI Bewertung":                            "KI Bewertung",
    "KI-Bewertung":                            "KI Bewertung",
}

# ============================================================
# MASTER-EXCEL: Konfiguration (identisch in allen drei
# Vorgänger-Skripten)
# ============================================================
MASTER_EXCEL_PFAD   = r"C:\Users\de020409\Sanofi\FBC Betriebsübergreifende Dokumente - General\!Systembewertungen_CS\00_Serienbrief\Systembewertungen_GESAMT.xlsx"
MASTER_SHEET_NAME    = "SysBew"
MASTER_TABELLE_NAME  = "Tabelle1"

# Spalten, deren Name im Skript anders lautet als in der Master-Excel.
# Schlüssel = Feldname im Skript (data-Dict), Wert = Spaltenname in der Master-Excel.
MASTER_SPALTEN_MAPPING = {
    "Erkannte_Version": "Erkannte Version2",
}

# ============================================================
# HILFSFUNKTIONEN
# ============================================================
def get_checkbox_state_from_sdt(sdt):
    ns = {"w14": "http://schemas.microsoft.com/office/word/2010/wordml"}
    checkbox = sdt.find(".//w14:checkbox", ns)
    if checkbox is None:
        return None
    checked = checkbox.find(".//w14:checked", ns)
    if checked is None:
        return "c"
    val = checked.get("{http://schemas.microsoft.com/office/word/2010/wordml}val")
    return "r" if val == "1" else "c"

def parse_text_checkboxes(text):
    """
    Parst Text-Checkboxen im Serienbrief-Format, z.B.:
    'c Critical* c Major r minor c N/A' -> geordnete Liste von Zuständen.
    r = angekreuzt, c = leer. Reihenfolge entspricht der Reihenfolge im Text,
    genau wie bei echten SDT-Checkboxen.
    """
    if not text:
        return []
    marker_positions = list(re.finditer(r"(?:^|\s)([rc])(?=\s)", text))
    if not marker_positions:
        return []
    result = []
    for i, m in enumerate(marker_positions):
        marker = m.group(1)
        start = m.end()
        end = marker_positions[i + 1].start() if i + 1 < len(marker_positions) else len(text)
        label = text[start:end].strip()
        result.append({"state": marker, "label": label})
    return result

def get_checkboxes_from_cell(cell):
    result = []
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for sdt in cell._element.findall(".//w:sdt", ns):
        state = get_checkbox_state_from_sdt(sdt)
        if state is not None:
            result.append({"state": state})
    if result:
        return result
    # Fallback: keine echten Formularfelder gefunden -> Text-Checkboxen (Serienbrief) parsen
    text = get_cell_text(cell)
    return parse_text_checkboxes(text)

def get_cell_text(cell):
    parts = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
    return " ".join(parts)

def text_matches(search_term, cell_text):
    s = search_term.lower().strip()
    c = cell_text.lower().strip()
    if s in c or c in s:
        return True
    s_clean = re.sub(r"[^a-z0-9äöüß]", "", s)
    c_clean = re.sub(r"[^a-z0-9äöüß]", "", c)
    return bool(s_clean and c_clean and
                (s_clean in c_clean or c_clean in s_clean))

def _is_boilerplate(text):
    patterns = [
        r"^weiter mit", r"^ende der systembewertung",
        r"^qualifizierung nach", r"^validierung nach",
        r"^qu-sop-", r"^qu-ope-", r"^qu-mt-",
        r"^fra-qu-", r"^keine mlcs", r"^\(info:",
        r"^begründung der gxp",
        r"^1\.systemname, mlcs",
        r"^anlagen-ids/equipment-nr\./qc-id$",
        r"^system-identifier",
    ]
    t = text.lower().strip()
    return any(re.match(p, t) for p in patterns)

# ============================================================
# ABLEITUNGS-FUNKTIONEN
# ============================================================
def parse_betrieb(raw):
    result = {}
    if not raw:
        return result
    m = re.search(r'\b([A-Z]\d+)\b', raw)
    if m:
        result["Gebaeude"] = m.group(1)
    if "/" in raw:
        betrieb = raw.split("/")[0].strip()
        result["Betrieb"] = betrieb
    elif "," in raw:
        parts = [p.strip() for p in raw.split(",")]
        nicht_gebaeude = [p for p in parts
                          if not re.match(r'^[A-Z]\d+$', p)]
        result["Betrieb"] = " ".join(nicht_gebaeude).strip()
    else:
        result["Betrieb"] = raw.strip()
    betrieb = result.get("Betrieb", "")
    if betrieb:
        result["API"] = betrieb.split()[0].strip().rstrip(",")
    return result

def parse_dok_nr(filename):
    m = re.search(r'(QU-[A-Z]+-\d+)', filename, re.IGNORECASE)
    return m.group(1) if m else None

def parse_sw_version(text):
    result = {}
    if not text:
        return result
    m = re.search(r'[Vv]ersion\s+([A-Z0-9][A-Z0-9._\-]+)', text)
    if m:
        result["SW-Version / Typ:"] = m.group(1)
    for pattern in [r'(800xA)', r'(Freelance)', r'(WinCC)',
                    r'(PCS\s*7)', r'(SCADA\s*\S+)', r'(TIA[-\s]*Portal)']:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            result["SW-Name:"] = m.group(1)
            break
    return result

def parse_gkat_subtypen(text):
    result = {}
    subtypen = {
        "GKATB1": r'B1', "GKATB2": r'B2', "GKATB3": r'B3',
        "GKATC1": r'C1', "GKATC2": r'C2', "GKATC3": r'C3',
    }
    for col, pattern in subtypen.items():
        result[col] = "r" if re.search(pattern, text or "") else "c"
    return result

def parse_gxp_einfluss(text):
    result = {
        "GxP_Produktqualitaet":    "",
        "GxP_Patientensicherheit": "",
        "GxP_Datenintegritaet":    "",
    }
    if not text or not text.strip():
        return result

    t = text.strip()
    hat_produkt = "produktqualität" in t.lower()
    hat_patient = "patientensicherheit" in t.lower()
    hat_daten   = ("datenintegrität" in t.lower() or
                   "datenintegrit" in t.lower())

    if hat_produkt or hat_patient or hat_daten:
        # Nur strukturiert splitten wenn Doppelpunkt vorhanden
        hat_doppelpunkt = bool(re.search(
            r'(?i)(produktqualität\s*:|patientensicherheit\s*:'
            r'|datenintegrit[äa]t\s*:)',
            t
        ))
        if not hat_doppelpunkt:
            result["GxP_Produktqualitaet"] = t
            return result

        teile = re.split(
            r'(?i)(produktqualität\s*[:\-]?|patientensicherheit\s*[:\-]?'
            r'|datenintegrit[äa]t\s*[:\-]?)',
            t
        )
        aktueller_key = None
        for teil in teile:
            tl = teil.lower().strip().rstrip(":").rstrip("-").strip()
            if tl == "produktqualität":
                aktueller_key = "GxP_Produktqualitaet"
            elif tl == "patientensicherheit":
                aktueller_key = "GxP_Patientensicherheit"
            elif tl in ("datenintegrität", "datenintegritaet"):
                aktueller_key = "GxP_Datenintegritaet"
            elif aktueller_key and teil.strip():
                # Komma-Artefakte am Anfang entfernen
                cleaned = teil.strip().lstrip(",").lstrip(".").strip()
                if cleaned:
                    result[aktueller_key] = (
                        result[aktueller_key] + " " + cleaned
                    ).strip()
    else:
        result["GxP_Produktqualitaet"] = t

    return result

def berechne_systemtyp_ce(data):
    if any(data.get(f) == "r"
           for f in ["Subtyp_PCS", "Subtyp_LCE", "Subtyp_EE"]):
        return "r"
    return "c"

# ============================================================
# TEMPLATE-VERSIONSERKENNUNG (identisch in allen drei
# Vorgänger-Skripten)
# ============================================================
def detect_template_version(doc):
    """
    Erkennt die Template-Version anhand struktureller Merkmale:
    V8:  Kein KI-Kapitel
    V10: KI-Kapitel + N/A bei Vereinfachte Qualifizierung
         ODER KI-Kapitel + Testtiefe-Eintrag in Kap. 2 (Checkbox oder Text)
    V11: GAMP 5 (2nd Edition) im Text
         ODER KI-Kapitel ohne obige V10-Merkmale
    """
    full_text = ""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                full_text += get_cell_text(cell).lower() + " "

    # Merkmal 1: GAMP 5 2nd Edition → eindeutig V11
    if "2nd edition" in full_text:
        return 11

    # Merkmal 2: Hat das Dokument ein KI-Kapitel?
    has_ki = any(k in full_text for k in [
        "künstliche intelligenz",
        "ki zum einsatz",
        "reifegrad",
        "autonomie",
        "steuerungsdesign",
    ])

    if not has_ki:
        return 8  # Kein KI-Kapitel → V8

    # Merkmal 3: Hat "Vereinfachte Qualifizierung" eine N/A-Option?
    has_vq_na = False
    for table in doc.tables:
        rows = list(table.rows)
        for row_idx, row in enumerate(rows):
            for cell in row.cells:
                text = get_cell_text(cell).lower()
                if "vereinfachte qualifizierung" in text:
                    for next_row in rows[row_idx+1:row_idx+6]:
                        for nc in next_row.cells:
                            nc_text = get_cell_text(nc).strip().lower()
                            cbs = get_checkboxes_from_cell(nc)
                            if nc_text in ["n/a", " n/a", "na"] and cbs:
                                has_vq_na = True

    if has_vq_na:
        return 10  # KI vorhanden + N/A → V10

    # Merkmal 4: Testtiefe in Kap. 2 vorhanden?
    # V10: Testtiefe steht in Zusammenfassungstabelle
    #      als echte Checkbox ODER als Text "r Gering/Mittel/Hoch/N/A"
    # V11: Testtiefe NUR als Matrix in Kap. 8
    has_testtiefe_in_kap2 = False
    for table in doc.tables:
        table_text = ""
        for row in table.rows:
            for cell in row.cells:
                table_text += get_cell_text(cell).lower() + " "

        # Suche Zusammenfassungstabelle (hat mlcs + testtiefe + gxp)
        if not all(k in table_text for k in [
                "mlcs", "testtiefe", "gxp"]):
            continue

        for row in table.rows:
            for cell in row.cells:
                cell_text = get_cell_text(cell).lower()
                # Echte Checkboxen in Testtiefe-Zelle
                if "testtiefe" in cell_text:
                    cbs = get_checkboxes_from_cell(cell)
                    if cbs:
                        has_testtiefe_in_kap2 = True
                # Text-Checkboxen: "r gering", "r mittel", "r hoch", "r n/a"
                if re.search(
                    r'\br\s+(gering|mittel|hoch|n/a)\b',
                    cell_text
                ):
                    has_testtiefe_in_kap2 = True

    if has_testtiefe_in_kap2:
        return 10  # Testtiefe in Kap. 2 → V10
    else:
        return 11  # Keine Testtiefe in Kap. 2 → V11

# ============================================================
# EXTRAKTIONS-FUNKTIONEN (identisch in allen drei
# Vorgänger-Skripten)
# ============================================================
def extract_mlcs_id(doc):
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if len(cells) < 2:
                continue
            label = get_cell_text(cells[0])
            if "system id" in label.lower() and "mlcs" in label.lower():
                seen = set()
                for i in range(len(cells)-1, 0, -1):
                    val = get_cell_text(cells[i]).strip()
                    if val in seen:
                        continue
                    seen.add(val)
                    val = val.replace("System ID (MLCS):", "").strip()
                    val = re.sub(r"(?i)^mlcs[\s\-:]*", "", val).strip()
                    val = re.sub(r"(?i)^id[\s\-:]*", "", val).strip()
                    val = re.sub(r"gemäß\s+\S+", "", val).strip()
                    val = re.sub(r"\(Info:.*?\)", "", val).strip()
                    val = re.sub(r"\(MLCS.*?\)", "", val,
                                 flags=re.IGNORECASE).strip()
                    if "system-identifier" in val.lower():
                        continue
                    if "cs inventarnummer" in val.lower():
                        continue
                    if val and not _is_boilerplate(val):
                        return val
    return None

def extract_ueberlagertes_mlcs(doc):
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if len(cells) < 2:
                continue
            label = get_cell_text(cells[0])
            if ("schnittstelle zu anderen" in label.lower() and
                    "übergeordneten" in label.lower()):
                seen = set()
                for i in range(len(cells)-1, 0, -1):
                    val = get_cell_text(cells[i]).strip()
                    if val in seen:
                        continue
                    seen.add(val)
                    val = re.sub(
                        r"schnittstelle zu anderen/übergeordneten systemen:?",
                        "", val, flags=re.IGNORECASE
                    ).strip()
                    val = re.sub(r"(?i)^mlcs[\s\-:]*", "", val).strip()
                    val = re.sub(r"(?i)^id[\s\-:]*", "", val).strip()
                    if val and not _is_boilerplate(val):
                        return val
    return None

def extract_systemtyp_zugang(doc):
    """
    Liest Offen/Geschlossen/N-A aus der Systemtyp-Zeile in Kapitel 1
    (Zugangsbeschränkung nach 21 CFR Part 11, §11.3(4)).
    """
    result = {}
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if len(cells) < 2:
                continue
            label = get_cell_text(cells[0])
            if "systemtyp" in label.lower() and "zugang" in label.lower():
                cbs = get_checkboxes_from_cell(cells[1])
                if len(cbs) >= 3:
                    result["Offen"]       = cbs[0]["state"]
                    result["Geschlossen"] = cbs[1]["state"]
                    result["NA"]          = cbs[2]["state"]
                return result
    return result

def extract_version_freigabedatum(doc):
    """
    Liest die höchste Version aus der 'Version / Freigabedatum'-Tabelle auf
    der ersten Seite. Diese Tabelle kann mehrere Revisionszeilen enthalten;
    die aktuellste (höchste) Version steht nicht immer in der ersten
    Datenzeile, sondern manchmal weiter unten (z.B. Zeile 3).
    """
    for table in doc.tables:
        rows = table.rows
        if not rows:
            continue
        header0 = get_cell_text(rows[0].cells[0]) if len(rows[0].cells) >= 1 else ""
        if "version" in header0.lower() and "freigabe" in header0.lower():
            versionen = []
            for row in rows[1:]:
                if len(row.cells) < 1:
                    continue
                val = get_cell_text(row.cells[0]).strip()
                m = re.match(r"^(\d+\.\d+)", val)
                if m:
                    versionen.append(m.group(1))
            if versionen:
                return max(versionen, key=lambda v: [int(x) for x in v.split(".")])
    return None

def extract_anlage(doc):
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if len(cells) < 2:
                continue
            label = get_cell_text(cells[0])
            if ("anlagen-id" in label.lower() or
                    "equipment-nr" in label.lower()):
                seen = set()
                for i in range(len(cells)-1, 0, -1):
                    val = get_cell_text(cells[i]).strip()
                    if val in seen:
                        continue
                    seen.add(val)
                    val = re.sub(
                        r"anlagen-ids/equipment-nr\./qc-id:?",
                        "", val, flags=re.IGNORECASE
                    ).strip()
                    if val and not _is_boilerplate(val):
                        return val
    return None

def extract_schnittstelle(doc):
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if len(cells) < 2:
                continue
            label = get_cell_text(cells[0])
            if "schnittstelle zu anderen" in label.lower():
                seen = set()
                for i in range(len(cells)-1, 0, -1):
                    val = get_cell_text(cells[i]).strip()
                    if val in seen:
                        continue
                    seen.add(val)
                    val = re.sub(
                        r"schnittstelle zu anderen.*?systemen:?",
                        "", val, flags=re.IGNORECASE
                    ).strip()
                    if val and not _is_boilerplate(val):
                        return val
    return None

def extract_besonderheiten(doc):
    """Besonderheiten-Text aus Kapitel 2 - sucht in allen Zellen der Zeile."""
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            for i, cell in enumerate(cells):
                text = get_cell_text(cell)
                if "besonder" in text.lower() and len(text) > 20:
                    # Suche in ALLEN folgenden Zellen der Zeile
                    for j in range(i + 1, len(cells)):
                        val = get_cell_text(cells[j])
                        if val and not _is_boilerplate(val):
                            return val
                    # Fallback: Text steht im Label selbst nach ":"
                    if ":" in text:
                        val = text.split(":", 1)[1].strip()
                        if val and not _is_boilerplate(val):
                            return val
    return None

def extract_begruendung_gxp(doc):
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if len(cells) < 1:
                continue
            label = get_cell_text(cells[0])
            if ("begründung" in label.lower() and
                    "gxp" in label.lower()):
                if ":" in label:
                    text = label.split(":", 1)[1].strip()
                    if text:
                        return text
                if len(cells) > 1:
                    val = get_cell_text(cells[1])
                    if val and not _is_boilerplate(val):
                        return val
                return label
    return None

def extract_neuerstellung(doc):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = get_cell_text(cell)
                if ("neuerstellung" in text.lower() and
                        "änderung" in text.lower()):
                    checkboxes = get_checkboxes_from_cell(cell)
                    if len(checkboxes) >= 2:
                        return {
                            "Neuerstellung": checkboxes[0]["state"],
                            "Revisioniert":  checkboxes[1]["state"],
                        }
    return {}

def extract_gxp_relevan_bc(doc):
    result = {}
    keywords_gxp = ["gxp relevanz", "gxp-relevanz"]
    keywords_bc  = ["business kritikalität", "geschäftskritisch",
                    "business kritisch"]

    for table in doc.tables:
        rows = list(table.rows)
        for row_idx, row in enumerate(rows):
            for cell in row.cells:
                text = get_cell_text(cell).lower()

                if ("GxP_Relevan_JA" not in result and
                        any(k in text for k in keywords_gxp)):
                    for next_row in rows[row_idx+1:row_idx+6]:
                        for nc in next_row.cells:
                            nc_text = get_cell_text(nc).strip().lower()
                            # Marker-Präfix (r/c) aus Text-Checkboxen entfernen, z.B. "c ja" -> "ja"
                            nc_text_clean = re.sub(r"^[rc]\s+", "", nc_text).strip()
                            cbs = get_checkboxes_from_cell(nc)
                            if not cbs:
                                continue
                            if nc_text_clean == "ja":
                                result["GxP_Relevan_JA"]   = cbs[0]["state"]
                            elif nc_text_clean == "nein":
                                result["GxP_Relevan_NEIN"] = cbs[0]["state"]

                if ("BCkritisch" not in result and
                        any(k in text for k in keywords_bc)):
                    for next_row in rows[row_idx+1:row_idx+6]:
                        for nc in next_row.cells:
                            nc_text = get_cell_text(nc).strip().lower()
                            nc_text_clean = re.sub(r"^[rc]\s+", "", nc_text).strip()
                            cbs = get_checkboxes_from_cell(nc)
                            if not cbs:
                                continue
                            if nc_text_clean == "ja":
                                result["BCkritisch"]   = cbs[0]["state"]
                            elif nc_text_clean == "nein":
                                result["BCunkritisch"] = cbs[0]["state"]

    return result

def extract_text_fields(doc, text_field_mapping=None):
    """Nutzt standardmäßig die gemeinsame TEXT_FIELD_MAPPING; ein Modul
    kann bei Bedarf ein eigenes Mapping übergeben (aktuell ungenutzt, da
    TEXT_FIELD_MAPPING in V8/V10/V11 identisch war)."""
    if text_field_mapping is None:
        text_field_mapping = TEXT_FIELD_MAPPING
    result = {}
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if len(cells) < 2:
                continue
            if _ist_rollen_zeile(cells):
                continue  # Deckblatt-Rollenzeile, siehe extract_deckblatt_rollen()
            label_text = get_cell_text(cells[0])
            for search_term, excel_col in text_field_mapping.items():
                if text_matches(search_term, label_text):
                    value = ""
                    for i in range(1, len(cells)):
                        candidate = get_cell_text(cells[i])
                        if candidate and not _is_boilerplate(candidate):
                            value = candidate
                            break
                    if value and excel_col not in result:
                        result[excel_col] = value
                    break
    return result

def extract_checkboxes_formularfelder(doc, checkbox_mapping):
    """checkbox_mapping ist versionsabhängig (CHECKBOX_MAPPING_V8/V10/V11
    aus dem jeweiligen Erweiterungs-Modul) und muss übergeben werden."""
    result = {}
    for table_idx, table in enumerate(doc.tables):
        if table_idx not in checkbox_mapping:
            continue
        table_mapping = checkbox_mapping[table_idx]
        for row_idx, row in enumerate(table.rows):
            if row_idx not in table_mapping:
                continue
            row_mapping = table_mapping[row_idx]
            for cell_idx, cell in enumerate(row.cells):
                if cell_idx not in row_mapping:
                    continue
                cell_mapping = row_mapping[cell_idx]
                checkboxes   = get_checkboxes_from_cell(cell)
                for cb_idx, cb in enumerate(checkboxes):
                    if cb_idx in cell_mapping:
                        col = cell_mapping[cb_idx]
                        if col not in result:
                            result[col] = cb["state"]
    return result

def extract_history(doc):
    result = {}
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if len(cells) < 2:
                continue
            col0 = get_cell_text(cells[0])
            col1 = get_cell_text(cells[1])
            if ("version" in col0.lower() and
                    "freigabe" in col0.lower()):
                continue
            m = re.match(r"^(\d+\.\d+)\s*/\s*(.+)$", col0.strip())
            if m:
                result["Version_Historie"] = m.group(1)
                result["Historie"]         = col1.strip()
                datum = m.group(2).strip()
                if "letztes unterschriftsdatum" not in datum.lower():
                    result["Datum"] = datum
    return result

# ============================================================
# DECKBLATT: Rollen/Namen (Ersteller, SME, SI/PL, TSO, BSO, BQR, CSQ)
# ============================================================
# Bekannte Rollen-Label-Varianten -> Ziel-Spalte in der Master-Excel.
# "SI" und "PL" (Projektleiter) landen beide in der zusammengefassten
# Spalte "SI/PL", weil die Master-Excel dafür nur eine gemeinsame
# Spalte hat.
ROLLEN_LABEL_MAPPING = {
    "ersteller": "Ersteller",
    "sme": "SME",
    "si": "SI/PL",
    "pl": "SI/PL",
    "si/pl": "SI/PL",
    "systemintegrator": "SI/PL",
    "projektleiter": "SI/PL",
    "projektleitung": "SI/PL",
    "tso": "TSO",
    "bso": "BSO",
    "bqr": "BQR",
    "csq": "CSQ",
}

ROLLEN_SPALTEN = ("Ersteller", "SME", "SI/PL", "TSO", "BSO", "BQR", "CSQ")

# Reihenfolge wichtig: "SI/PL" bzw. die Projektleiter-Varianten müssen
# vor dem kürzeren "SI"/"PL" geprüft werden, damit z.B. "SI/PL" nicht
# fälschlich schon als reines "SI" erkannt wird.
_ROLLEN_LABEL_PATTERN = re.compile(
    r'^\s*(Ersteller|SME|SI\s*/\s*PL|Systemintegrator|Projektleit(?:er|ung)|SI|PL|TSO|BSO|BQR|CSQ)\b[:\s]*',
    re.IGNORECASE
)

# Bekannte Word-Platzhaltertexte von noch nicht befüllten
# Content-Control-Dropdownfeldern - dürfen nicht als Name durchgehen.
_ROLLEN_PLATZHALTER = [
    r"^klicken sie",
    r"^choose an item",
    r"^wählen sie",
    r"^auswahl\.*$",
    r"^bitte auswählen",
]

def _ist_rollen_platzhalter(text):
    t = text.lower().strip()
    return any(re.match(p, t) for p in _ROLLEN_PLATZHALTER)

def _ist_bekanntes_rollen_label(text):
    """True, wenn `text` (nach Bereinigung) exakt einem der bekannten
    Rollen-Label entspricht - dient nur der Erkennung, ob eine
    Tabellenzeile überhaupt die Deckblatt-Rollenzeile ist."""
    l_clean = re.sub(r'[^a-zäöüß/]', '', text.strip().lower())
    return l_clean in ROLLEN_LABEL_MAPPING

def _ist_rollen_zeile(cells):
    """
    True, wenn mindestens 3 Zellen dieser Tabellenzeile ein bekanntes
    Rollen-Label (Ersteller/SME/SI-PL/TSO/BSO/BQR/CSQ) enthalten - dann
    ist es die Deckblatt-Rollenzeile. Wichtig auch für
    extract_text_fields(): "Ersteller" würde sonst über den laschen
    Teilstring-Vergleich in text_matches() fälschlich als Treffer für
    den Suchbegriff "Hersteller / SW-Ersteller / Lieferant" durchgehen
    (weil "ersteller" darin als Teilwort steckt) und die Nachbarzelle
    (z.B. "SME") als Hersteller-Wert übernehmen.
    """
    return sum(1 for c in cells if _ist_bekanntes_rollen_label(get_cell_text(c))) >= 3

def _normalisiere_rollen_label(label):
    """
    Ordnet ein Rollen-Label (Feldbezeichnung auf dem Deckblatt) der
    passenden Excel-Spalte zu. Bekannte Kurzformen (z.B. "PL" für
    Projektleiter) werden erkannt und der zusammengefassten Spalte
    "SI/PL" zugeordnet. Nicht erkennbare Funktionsbezeichnungen
    fallen auf "SME" zurück - AUSSER das Label deutet auf eine
    Projektleitungs-Funktion hin (enthält "pl" als eigenständiges
    Kürzel oder "projektleit..."), dann "SI/PL".
    """
    l = label.strip().lower()
    l_clean = re.sub(r'[^a-zäöüß/]', '', l)
    if l_clean in ROLLEN_LABEL_MAPPING:
        return ROLLEN_LABEL_MAPPING[l_clean]
    if re.search(r'\bpl\b', l) or "projektleit" in l:
        return "SI/PL"
    return "SME"

def _get_sdt_text(cell_element, ns):
    """
    Liest den aktuell angezeigten Text aus Content-Control-Feldern
    (w:sdt, z.B. Dropdown-Auswahlfelder) einer Zelle ein. Nötig, weil
    solche Felder ihren Text nicht zuverlässig über die normalen
    Zellen-Absätze (cell.paragraphs) liefern - aus demselben Grund
    liest get_checkboxes_from_cell() Checkbox-Formularfelder direkt
    über die SDT-XML-Struktur statt über den Absatztext.
    """
    texte = []
    for t in cell_element.findall(".//w:sdtContent//w:t", ns):
        if t.text:
            texte.append(t.text)
    return "".join(texte).strip()

def extract_deckblatt_rollen(doc):
    """
    Liest die Namen zu den Rollen Ersteller/SME/SI-PL/TSO/BSO/BQR/CSQ
    vom Deckblatt. Jede Rollen-Zelle besteht aus dem Label
    ("Ersteller" usw.) als Fließtext plus dem Namen direkt daneben
    bzw. als Dropdown-Auswahl darunter/dahinter.

    Eine Tabellenzeile wird als Deckblatt-Rollenzeile behandelt, wenn
    mindestens 3 ihrer Zellen ein bekanntes Rollen-Label enthalten -
    ERST DANN werden alle Zellen dieser Zeile ausgewertet (auch die,
    deren Label NICHT bekannt ist). So greift der Fallback auch für
    abweichend benannte Funktionen, die sonst nie erkannt würden:
    über _normalisiere_rollen_label() geht alles mit "PL"/
    "Projektleit..." nach Spalte "SI/PL", alles andere nicht
    erkennbare nach Spalte "SME".

    Kommt eine Rolle mehrfach vor (z.B. zwei BSO-Felder, weil es zwei
    Business Owner gibt), werden alle für die jeweilige Spalte
    gefundenen Namen mit einem Zeilenumbruch ("\\n" - erscheint in
    Excel als Alt+Enter-Umbruch innerhalb der Zelle) zusammengefügt.
    """
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    treffer = {spalte: [] for spalte in ROLLEN_SPALTEN}

    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if not _ist_rollen_zeile(cells):
                continue  # keine Deckblatt-Rollenzeile
            labels = [get_cell_text(c).strip() for c in cells]

            for cell, label in zip(cells, labels):
                if not label:
                    continue
                m = _ROLLEN_LABEL_PATTERN.match(label)
                # Bekanntes Label mit Text danach (z.B. "Ersteller: Max
                # Mustermann") -> Name schon im Fließtext. Sonst (Label
                # allein, oder Label komplett unbekannt) -> Name steckt
                # im Dropdown-/Content-Control-Feld derselben Zelle.
                name = label[m.end():].strip() if m else ""
                if not name:
                    name = _get_sdt_text(cell._element, ns)
                if not name or _ist_rollen_platzhalter(name):
                    continue  # Feld auf dem Deckblatt noch nicht befüllt
                rollen_label = m.group(1) if m else label
                spalte = _normalisiere_rollen_label(rollen_label)
                if name not in treffer[spalte]:
                    treffer[spalte].append(name)

    return {spalte: "\n".join(namen) for spalte, namen in treffer.items() if namen}

# ============================================================
# VALIDIERUNG: Vollständigkeitsprüfung
# ============================================================
def validiere_kategorien(data, kategorien):
    """
    Prüft je Kategorie, ob genau EIN "r"-Wert (angekreuzt) vorhanden ist.
    c-Werte werden nicht angezeigt - nur Kategoriename, ausgewählter
    Wert (falls genau einer) und Status. `kategorien` kommt aus dem
    jeweiligen Erweiterungs-Modul (VALIDATION_KATEGORIEN_V8/V10/V11).
    """
    print("\n" + "="*55)
    print("✅ VALIDIERUNG: Vollständigkeitsprüfung")
    print("="*55)
    for name, optionen in kategorien:
        gefunden = [label for feld, label in optionen if data.get(feld) == "r"]
        if len(gefunden) == 1:
            print(f"  ✅ {name:<32} = \"{gefunden[0]}\"")
        elif len(gefunden) == 0:
            print(f"  ❗ {name:<32} = KEIN Wert ausgewählt (erwartet: genau 1)")
        else:
            werte = ", ".join(f'"{g}"' for g in gefunden)
            print(f"  ❌ {name:<32} = MEHRERE Werte: {werte} (erwartet: genau 1)")

# ============================================================
# DATEI-VALIDIERUNG (identisch in allen drei Vorgänger-Skripten
# am Anfang von parse_systembewertung_vXX)
# ============================================================
def validiere_docx_datei(docx_path):
    """
    Bereinigt den Pfad (Anführungszeichen vom Drag & Drop) und prüft,
    ob die Datei existiert, eine .docx/.doc-Endung hat und lokal
    verfügbar ist (kein OneDrive-Platzhalter, nicht durch Word gesperrt).

    Gibt den bereinigten Pfad zurück, wenn alles passt, sonst None
    (Fehlermeldung wurde bereits ausgegeben). Wirft FileNotFoundError,
    wenn die Datei nicht existiert (wie im Vorgänger-Verhalten).
    """
    docx_path = docx_path.strip().strip('"').strip("'")

    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"Datei nicht gefunden: {docx_path}")

    ext = os.path.splitext(docx_path)[1].lower()
    if ext not in [".docx", ".doc"]:
        print(f"\n  ❌ Keine Word-Datei: '{os.path.basename(docx_path)}'")
        print(f"     Nur .docx Dateien werden unterstützt.")
        return None

    # Prüfe ob Datei lokal verfügbar (kein OneDrive-Platzhalter)
    try:
        with open(docx_path, 'rb') as f:
            header = f.read(4)
        if header[:2] != b'PK':
            print(f"\n  ❌ Datei ist kein gültiges .docx!")
            print(f"     Möglicherweise nur OneDrive-Platzhalter.")
            print(f"     → Rechtsklick → 'Immer auf diesem Gerät behalten'")
            return None
    except PermissionError:
        print(f"\n  ❌ Datei ist gesperrt!")
        print(f"     ⚠️  WICHTIG: Bitte Word KOMPLETT schließen!")
        print(f"     → Alle Word-Fenster schließen")
        print(f"     → Task-Manager prüfen (Ctrl+Shift+Esc)")
        print(f"     → Nach WINWORD.EXE suchen und beenden")
        print(f"     → Dann Skript erneut starten")
        return None
    except OSError as e:
        print(f"\n  ❌ Datei-Zugriffsfehler: {e}")
        print(f"     OneDrive-Synchronisation abwarten und erneut versuchen.")
        return None

    return docx_path

# ============================================================
# MASTER-EXCEL-OUTPUT (identisch in allen drei Vorgänger-Skripten)
# ============================================================
def write_to_master_excel(data, docx_path, _versuch=1, _max_versuche=3):
    """
    Fügt die extrahierten Daten als neue Zeile in die Excel-Tabelle
    ('Tabelle1') der Master-Excel-Datei ein - per COM-Automatisierung
    (echtes Excel), NICHT per openpyxl.

    Grund: Die Master-Datei enthält eine Excel-Tabelle mit festem Bereich,
    ein Sensitivity-Label (Microsoft Purview Information Protection), eine
    externe Verknüpfung sowie Kommentare. openpyxl erhält beim Neuschreiben
    (load_workbook + save) weder die Tabellen-Bereichserweiterung noch das
    Sensitivity-Label noch die externe Verknüpfung zuverlässig -> Excel
    meldet beim nächsten Öffnen "Arbeitsmappe repariert". Über
    ListObject.ListRows.Add() erweitert echtes Excel die Tabelle korrekt und
    lässt alle anderen Dateibestandteile unangetastet.

    Bei transienten COM-Fehlern (z.B. RPC_E_DISCONNECTED, wenn Excel
    während des Vorgangs die Verbindung verliert - etwa durch OneDrive-
    Sync-Konflikte) wird bis zu _max_versuche-mal automatisch erneut
    versucht, ab dem zweiten Versuch immer mit einer komplett neuen,
    unsichtbaren Excel-Instanz (kein Anhängen an eine ggf. instabile
    laufende Instanz mehr).

    Benötigt: pip install pywin32   (nur unter Windows, mit installiertem Excel)
    """
    try:
        import win32com.client as win32
    except ImportError:
        print("\n  ❌ pywin32 ist nicht installiert.")
        print("     Bitte ausführen: pip install pywin32")
        return None

    if not os.path.exists(MASTER_EXCEL_PFAD):
        print(f"\n  ❌ Master-Excel nicht gefunden:")
        print(f"     {MASTER_EXCEL_PFAD}")
        print("     Bitte prüfen, ob OneDrive/SharePoint-Sync aktiv ist.")
        return None

    excel = None
    wb_com = None
    wir_haben_geoeffnet = False
    wir_haben_excel_gestartet = False

    try:
        try:
            if _versuch > 1:
                # Ab dem zweiten Versuch NIE an eine laufende Instanz anhängen -
                # die könnte genau die Ursache des vorherigen Fehlers sein.
                raise RuntimeError("Erzwinge neue Excel-Instanz für diesen Versuch")
            excel = win32.GetActiveObject("Excel.Application")
            _ = excel.Workbooks.Count  # Gesundheitscheck: haengt eine alte/verwaiste Instanz fest?
        except Exception:
            excel = win32.DispatchEx("Excel.Application")
            wir_haben_excel_gestartet = True

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        ziel_pfad_abs = os.path.abspath(MASTER_EXCEL_PFAD).lower()
        for offene_wb in excel.Workbooks:
            try:
                if offene_wb.FullName.lower() == ziel_pfad_abs:
                    wb_com = offene_wb
                    break
            except Exception:
                continue

        if wb_com is None:
            wb_com = excel.Workbooks.Open(MASTER_EXCEL_PFAD, UpdateLinks=0, ReadOnly=False)
            wir_haben_geoeffnet = True

        ws_com = wb_com.Worksheets(MASTER_SHEET_NAME)

        try:
            tabelle = ws_com.ListObjects(MASTER_TABELLE_NAME)
        except Exception:
            print(f"\n  ❌ Excel-Tabelle '{MASTER_TABELLE_NAME}' nicht auf Sheet '{MASTER_SHEET_NAME}' gefunden.")
            if wir_haben_geoeffnet:
                wb_com.Close(SaveChanges=False)
            if wir_haben_excel_gestartet:
                excel.Quit()
            return None

        header_range = tabelle.HeaderRowRange
        header = [header_range.Cells(1, c).Value for c in range(1, header_range.Columns.Count + 1)]

        neue_zeile = tabelle.ListRows.Add()

        for col_idx, col_name in enumerate(header, start=1):
            if not col_name:
                continue
            script_key = col_name
            for skript_name, master_name in MASTER_SPALTEN_MAPPING.items():
                if master_name == col_name:
                    script_key = skript_name
                    break
            wert = data.get(script_key, "")
            neue_zeile.Range.Cells(1, col_idx).Value = wert

        ziel_zeile = neue_zeile.Range.Row
        wb_com.Save()

        # Duplikat-Hinweis (Dok.-Nr. + Version) für den soeben eingefügten Eintrag.
        # Rein informativ - darf den Haupt-Workflow niemals stören.
        try:
            dok_nr_col_idx  = header.index("Dok. -Nr.") + 1 if "Dok. -Nr." in header else None
            version_col_idx = header.index("Version") + 1 if "Version" in header else None
            if dok_nr_col_idx and version_col_idx:
                neuer_dok_nr = str(neue_zeile.Range.Cells(1, dok_nr_col_idx).Value or "").strip()
                neue_version = str(neue_zeile.Range.Cells(1, version_col_idx).Value or "").strip()
                PLATZHALTER = {"", "qu-ope-xxxxx", "n/a"}
                if neuer_dok_nr and neuer_dok_nr.lower() not in PLATZHALTER:
                    data_range = tabelle.DataBodyRange
                    dok_spalte = data_range.Columns(dok_nr_col_idx).Value
                    ver_spalte = data_range.Columns(version_col_idx).Value
                    if not isinstance(dok_spalte, tuple):
                        dok_spalte = ((dok_spalte,),)
                        ver_spalte = ((ver_spalte,),)
                    treffer = 0
                    for (dv,), (vv,) in zip(dok_spalte, ver_spalte):
                        if str(dv or "").strip() == neuer_dok_nr and str(vv or "").strip() == neue_version:
                            treffer += 1
                    if treffer > 1:
                        print(f"\n  ⚠️  HINWEIS: Möglicher Duplikat-Eintrag!")
                        print(f"     Dok.-Nr. '{neuer_dok_nr}' mit Version '{neue_version}' kommt")
                        print(f"     bereits {treffer}x in der Master-Excel vor (inkl. diesem neuen Eintrag).")
        except Exception:
            pass

        if wir_haben_geoeffnet:
            wb_com.Close(SaveChanges=False)  # bereits gespeichert
        if wir_haben_excel_gestartet:
            excel.Quit()

        if _versuch > 1:
            print(f"     ✅ Erfolgreich im {_versuch}. Versuch.")

        return ziel_zeile

    except Exception as e:
        print(f"\n  ❌ Fehler beim Schreiben in die Master-Excel (COM): {e}")
        try:
            if wir_haben_geoeffnet and wb_com is not None:
                wb_com.Close(SaveChanges=False)
            if wir_haben_excel_gestartet and excel is not None:
                excel.Quit()
        except Exception:
            pass

        if _versuch < _max_versuche:
            print(f"     Erneuter Versuch ({_versuch + 1}/{_max_versuche}) mit frischer Excel-Instanz...")
            time.sleep(3)
            return write_to_master_excel(data, docx_path, _versuch=_versuch + 1, _max_versuche=_max_versuche)

        print("     Falls dies ein 'OLE error 0x800a01a8' o.ä. ist: im Task-")
        print("     Manager prüfen, ob ein unsichtbarer EXCEL.EXE-Prozess von")
        print("     einem früheren Lauf noch offen ist, diesen beenden und")
        print("     erneut versuchen.")
        return None

# ============================================================
# KONSOLEN-VORSCHAU (identisch in allen drei Vorgänger-Skripten,
# stand vorher im jeweiligen __main__-Block)
# ============================================================
def zeige_datenvorschau(data):
    print("\n" + "="*55)
    print("📊 EXTRAHIERTE DATEN (Vorschau):")
    print("="*55)
    for col, value in sorted(data.items()):
        display_val = (
            str(value)[:60] + "..."
            if len(str(value)) > 60
            else str(value)
        )
        print(f"  {col:<35} = {display_val}")
