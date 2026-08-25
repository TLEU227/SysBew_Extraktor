# ============================================================
# sysbew_common.py
# Gemeinsame Basis der Systembewertungs-Extraktoren - 2.6
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
    # Klassifizierung Kapitel 1 (Lokales CS / Multi-Site-CS / Globales CS / Equipment ohne CS)
    "KLASS_Lokal", "KLASS_Multisite", "KLASS_Multisite_NurLokal", "KLASS_Multisite_LokalGlobal",
    "KLASS_Global", "KLASS_Global_1a", "KLASS_Global_1b", "KLASS_Global_2", "KLASS_Global_3",
    "KLASS_OhneCS",
    "KAT1", "KAT3", "KAT4", "KAT5", "KATNA",
    "Ersteller", "SME", "SI/PL", "TSO", "BSO", "BQR", "CSQ",
    "Datum", "Version_Historie", "Historie", "Bearbeiter",
    "Hersteller", "Phenix", "SAP",
    "ERESTYP1", "ERESTYP2", "ERESTYP3", "ERESTYP4", "ERESTYPNA",
    "TTIEFEHOCH", "TTIEFEMITTEL", "TTIEFENIEDRIG",
    # DI EE-Anforderungen (Kapitel 2, Zusammenfassungstabelle)
    "EE_P1", "EE_P2", "EE_P3", "EE_P4", "EE_NA",
    "Z1S1", "Z2S1", "Z3S1", "Z1S2", "Z1S3", "Z2S2", "Z2S3", "Z3S2", "Z3S3",
    # Periodic Review (Kapitel 2, Zusammenfassungstabelle) - 3 echte
    # Checkboxen in der Zelle (QU-SOP-0007359 / QU-SOP-0028559 / freie
    # Angabe), vorher wurde die mittlere (PR_SOP2) uebersehen und
    # "PR_Andere" zeigte faelschlich auf QU-SOP-0028559 statt auf die
    # freie Angabe.
    "PR_SOP", "PR_SOP2", "PR_Andere", "PR-Zyklisch",
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
    # "ja" = Zeile wurde über word_parser_main.py eingetragen, leer/"nein"
    # = manuell erfasst (wird vom Skript nie überschrieben)
    "Python ja/nein",
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

def extract_klassifizierung(doc):
    """
    Liest die Klassifizierung aus Kapitel 1 (Lokales CS / Multi-Site-CS
    inkl. "nur lokal"/"lokal und global" / Globales CS inkl. Klasse
    1a/1b/2/3 / Equipment ohne CS). Laut Formular ist hier
    Mehrfachauswahl möglich.

    Die eigentliche Checkbox-Zeile folgt direkt auf die Zeile mit dem
    Label "Klassifizierung: (Mehrfachauswahl möglich)" - deshalb wird
    per Textsuche zunächst diese Label-Zeile gesucht und dann die
    Folgezeile ausgewertet (Kapitel 1 steht in einer eigenen Tabelle,
    unabhängig von Tabellenindex/KI-Kapitel der jeweiligen Version -
    an einem echten Dokument bestätigt: 4 Spalten mit 1/3/5/1
    Checkboxen).
    """
    result = {}
    for table in doc.tables:
        rows = list(table.rows)
        for row_idx, row in enumerate(rows):
            cells = row.cells
            if not cells:
                continue
            label = get_cell_text(cells[0]).lower()
            if "klassifizierung" not in label or "mehrfachauswahl" not in label:
                continue
            if row_idx + 1 >= len(rows):
                continue
            werte = rows[row_idx + 1].cells

            cbs = get_checkboxes_from_cell(werte[0]) if len(werte) > 0 else []
            if cbs:
                result["KLASS_Lokal"] = cbs[0]["state"]

            cbs = get_checkboxes_from_cell(werte[1]) if len(werte) > 1 else []
            if len(cbs) >= 3:
                result["KLASS_Multisite"]               = cbs[0]["state"]
                result["KLASS_Multisite_NurLokal"]       = cbs[1]["state"]
                result["KLASS_Multisite_LokalGlobal"]    = cbs[2]["state"]

            cbs = get_checkboxes_from_cell(werte[2]) if len(werte) > 2 else []
            if len(cbs) >= 5:
                result["KLASS_Global"]    = cbs[0]["state"]
                result["KLASS_Global_1a"] = cbs[1]["state"]
                result["KLASS_Global_1b"] = cbs[2]["state"]
                result["KLASS_Global_2"]  = cbs[3]["state"]
                result["KLASS_Global_3"]  = cbs[4]["state"]

            cbs = get_checkboxes_from_cell(werte[3]) if len(werte) > 3 else []
            if cbs:
                result["KLASS_OhneCS"] = cbs[0]["state"]

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
        if _ist_rollen_tabelle(table):
            continue  # Deckblatt-Rollentabelle, siehe extract_deckblatt_rollen()
        for row in table.rows:
            cells = row.cells
            if len(cells) < 2:
                continue
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
# Reale Tabellenstruktur (in allen bisher geprüften V11-Dokumenten
# identisch, vermutlich auch V8/V10): eine Unterschriften-Tabelle mit
# je 2 Zeilen pro Rolle - Zeile 1 = Rollen-Label (Spalte 0, z.B. "TSO
# (Technical System Owner) (FBC ENG FF)") + Name (Spalte 1), Zeile 2 =
# Bestätigungstext (Spalte 0 und 1 enthalten denselben langen
# Fließtext - wird übersprungen). "Projektleiter/SME" ist im Dokument
# EINE kombinierte Rolle: der Name landet dann in BEIDEN Excel-Spalten
# "SI/PL" und "SME".
#
# Bekannte Rollen-Label-Varianten -> Ziel-Spalte(n) in der Master-
# Excel. Werte sind Listen, weil "Projektleiter/SME" auf zwei Spalten
# zugleich zeigt.
ROLLEN_LABEL_MAPPING = {
    "autor": ["Ersteller"],
    "ersteller": ["Ersteller"],
    "sme": ["SME"],
    "si": ["SI/PL"],
    "pl": ["SI/PL"],
    "si/pl": ["SI/PL"],
    "systemintegrator": ["SI/PL"],
    "projektleiter": ["SI/PL"],
    "projektleiterin": ["SI/PL"],
    "projektleitung": ["SI/PL"],
    "projektleiter/sme": ["SI/PL", "SME"],
    "projektleiterin/sme": ["SI/PL", "SME"],
    "tso": ["TSO"],
    "bso": ["BSO"],
    "bqr": ["BQR"],
    "csq": ["CSQ"],
}

ROLLEN_SPALTEN = ("Ersteller", "SME", "SI/PL", "TSO", "BSO", "BQR", "CSQ")

# Erfasst nur den Rollen-Kernbegriff am Zeilenanfang von Spalte 0,
# der ggf. folgende Klammer-Zusatz (Abteilung o.ä.) wird NICHT erfasst
# und bleibt unberücksichtigt. Reihenfolge wichtig: die kombinierte
# Form "Projektleiter/SME" muss vor "Projektleiter" allein geprüft
# werden, sonst würde nur "Projektleiter" greifen.
_ROLLEN_LABEL_PATTERN = re.compile(
    r'^\s*(Autor(?:in)?|Ersteller|'
    r'Projektleiter(?:in)?\s*/\s*SME|Projektleiter(?:in)?|'
    r'Systemintegrator|SI\s*/\s*PL|SME|SI|PL|TSO|BSO|BQR|CSQ)\b',
    re.IGNORECASE
)

# Bekannte Word-Platzhaltertexte von noch nicht befüllten Feldern -
# dürfen nicht als Name durchgehen.
_ROLLEN_PLATZHALTER = [
    r"^klicken sie",
    r"^choose an item",
    r"^wählen sie",
    r"^auswahl\.*$",
    r"^bitte auswählen",
    r"^n/?a$",
]

def _ist_rollen_platzhalter(text):
    t = text.lower().strip()
    return any(re.match(p, t) for p in _ROLLEN_PLATZHALTER)

def _ist_rollen_tabelle(table):
    """
    True, wenn mindestens 3 Zeilen dieser Tabelle mit einem bekannten
    Rollen-Label beginnen (Spalte 0) - dann ist es die Deckblatt-
    Unterschriften-/Rollentabelle. Wichtig auch für
    extract_text_fields(): diese Tabelle wird dort komplett
    übersprungen, damit z.B. "Autor ..." nicht versehentlich über den
    laschen Teilstring-Vergleich in text_matches() mit einem anderen
    Suchbegriff kollidiert.
    """
    treffer = 0
    for row in table.rows:
        cells = row.cells
        if not cells:
            continue
        if _ROLLEN_LABEL_PATTERN.match(get_cell_text(cells[0])):
            treffer += 1
    return treffer >= 3

def _normalisiere_rollen_label(rollen_kern):
    """
    Ordnet den erfassten Rollen-Kernbegriff (z.B. "TSO",
    "Projektleiter/SME", "Autor") der/den passenden Excel-Spalte(n)
    zu. `rollen_kern` ist entweder ein von _ROLLEN_LABEL_PATTERN
    erfasster bekannter Begriff, oder - falls die Zeile in einer
    erkannten Rollentabelle steht, aber KEIN bekanntes Label hat - der
    komplette (abweichend benannte) Zellentext. Nicht erkennbare
    Funktionsbezeichnungen fallen auf "SME" zurück, AUSSER der Text
    deutet auf eine Projektleitungs-Funktion hin (enthält "pl" als
    eigenständiges Kürzel oder "projektleit..."), dann "SI/PL".
    """
    l_clean = re.sub(r'[^a-zäöüß/]', '', rollen_kern.strip().lower())
    if l_clean in ROLLEN_LABEL_MAPPING:
        return ROLLEN_LABEL_MAPPING[l_clean]
    l = rollen_kern.strip().lower()
    if re.search(r'\bpl\b', l) or "projektleit" in l:
        return ["SI/PL"]
    return ["SME"]

def _get_sdt_text(cell_element, ns):
    """
    Liest den aktuell angezeigten Text aus Content-Control-Feldern
    (w:sdt, z.B. Dropdown-Auswahlfelder) einer Zelle ein - Fallback,
    falls eine Namens-Zelle über die normalen Zellen-Absätze
    (cell.paragraphs) leer erscheint, der Name aber in einem
    Formularfeld steckt (in den bisher geprüften Dokumenten stehen
    die Namen als reiner Fließtext, dieser Fallback greift daher
    normalerweise nicht, schadet aber nicht).
    """
    texte = []
    for t in cell_element.findall(".//w:sdtContent//w:t", ns):
        if t.text:
            texte.append(t.text)
    return "".join(texte).strip()

def extract_deckblatt_rollen(doc):
    """
    Liest die Namen zu den Rollen Ersteller/SME/SI-PL/TSO/BSO/BQR/CSQ
    aus der Deckblatt-Unterschriftentabelle (siehe Modul-Kommentar
    oben für die Tabellenstruktur).

    Nicht erkannte Rollenbezeichnungen (z.B. abweichend benannte
    Funktionen) werden trotzdem verarbeitet, sobald die Zeile in einer
    als Rollentabelle erkannten Tabelle steht - über
    _normalisiere_rollen_label() geht alles mit "PL"/"Projektleit..."
    nach Spalte "SI/PL", alles andere nicht erkennbare nach "SME".

    Kommt eine Rolle mehrfach vor (z.B. zwei BSO-Zeilen, weil es zwei
    Business Owner gibt), werden alle für die jeweilige Spalte
    gefundenen Namen mit einem Zeilenumbruch ("\\n" - erscheint in
    Excel als Alt+Enter-Umbruch innerhalb der Zelle) zusammengefügt.
    """
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    treffer = {spalte: [] for spalte in ROLLEN_SPALTEN}

    for table in doc.tables:
        if not _ist_rollen_tabelle(table):
            continue
        rows = list(table.rows)
        for row_idx, row in enumerate(rows):
            cells = row.cells
            if len(cells) < 2:
                continue
            label = get_cell_text(cells[0]).strip()
            if not label:
                continue
            name = get_cell_text(cells[1]).strip()
            if not name:
                name = _get_sdt_text(cells[1]._element, ns)
            # Bestätigungstext-Zeilen: Label und "Name" sind derselbe
            # lange Fließtext, oder der Wert ist für einen Namen viel
            # zu lang -> keine echte Namenszeile, überspringen.
            if (not name or len(name) > 60 or
                    name[:30].strip() == label[:30].strip() or
                    _ist_rollen_platzhalter(name)):
                continue
            m = _ROLLEN_LABEL_PATTERN.match(label)
            rollen_kern = m.group(1) if m else label
            spalten = _normalisiere_rollen_label(rollen_kern)

            # Sonderfall: Das Label steht manchmal nur als "SME"
            # (ohne "Projektleiter/"), obwohl es sich laut dem
            # Bestätigungstext der Folgezeile trotzdem um die
            # kombinierte Rolle "Projektleiter/SME" handelt (in den
            # geprüften Dokumenten uneinheitlich gehandhabt). Steht
            # dort "Projektleiter", zählt der Name zusätzlich als
            # SI/PL.
            if spalten == ["SME"] and row_idx + 1 < len(rows):
                bestaetigung = get_cell_text(rows[row_idx + 1].cells[0])
                if "projektleit" in bestaetigung.lower():
                    spalten = ["SME", "SI/PL"]

            for spalte in spalten:
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
# VALIDIERUNG: Mehrfachauswahl-Kategorien (Klassifizierung,
# DI EE-Anforderungen, Periodic Review - laut Formular ist hier
# AUSDRÜCKLICH Mehrfachauswahl möglich, mehrere angekreuzte Werte
# sind also KEIN Konflikt, anders als bei validiere_kategorien()).
# Identisch in V8/V10/V11, da diese Tabellen(-bereiche) versions-
# unabhängig gleich aufgebaut sind - deshalb hier zentral definiert
# statt in den Erweiterungs-Modulen.
# ============================================================
MEHRFACHAUSWAHL_KATEGORIEN = [
    ("Klassifizierung", [
        ("KLASS_Lokal", "Lokales CS"),
        ("KLASS_Multisite", "Multi-Site-CS"),
        ("KLASS_Multisite_NurLokal", "nur lokal"),
        ("KLASS_Multisite_LokalGlobal", "lokal und global"),
        ("KLASS_Global", "Globales CS"),
        ("KLASS_Global_1a", "Klasse 1a"),
        ("KLASS_Global_1b", "Klasse 1b"),
        ("KLASS_Global_2", "Klasse 2"),
        ("KLASS_Global_3", "Klasse 3"),
        ("KLASS_OhneCS", "Equipment ohne CS"),
    ]),
    ("DI EE-Anforderungen", [
        ("EE_P1", "P1"), ("EE_P2", "P2"), ("EE_P3", "P3"),
        ("EE_P4", "P4"), ("EE_NA", "N/A"),
    ]),
    ("Periodic Review", [
        ("PR_SOP", "QU-SOP-0007359"),
        ("PR_SOP2", "QU-SOP-0028559"),
        ("PR_Andere", "andere/freie Angabe"),
        ("PR-Zyklisch", "zyklische Requalifizierung"),
    ]),
]

# ============================================================
# NUR FÜR DIE VORSCHAU (zeige_datenvorschau): zusätzliche
# gruppierbare Felder, die bewusst NICHT in den Validierungsblöcken
# auftauchen, weil ihre Konfliktprüfung wegen bekannter Lücken
# unzuverlässig wäre (fehlende Testtiefe-N/A-Spalte, fehlende 3.
# SOP-Spalte bei Validierung/Qualifizierung - siehe README.md,
# "Bekannte Einschränkungen"). Für die reine Anzeige als eine Zeile
# statt drei/zwei Einzelfeldern reicht das aber.
# ============================================================
ANZEIGE_ZUSATZ_KATEGORIEN = [
    ("Testtiefe", [
        ("TTIEFENIEDRIG", "Gering"),
        ("TTIEFEMITTEL", "Mittel"),
        ("TTIEFEHOCH", "Hoch"),
    ]),
    ("Validierung/Qualifizierung nach SOP", [
        ("QUAL", "QU-SOP-0021736 (Qualifizierung)"),
        ("VAL", "QU-SOP-0049866 (Validierung)"),
    ]),
]

def validiere_mehrfachauswahl_kategorien(data, kategorien=MEHRFACHAUSWAHL_KATEGORIEN):
    """
    Wie validiere_kategorien(), aber für Kategorien mit laut Formular
    erlaubter Mehrfachauswahl: zeigt ALLE angekreuzten Werte einer
    Kategorie an (kein "❌ Konflikt" bei mehreren Treffern, das ist
    hier normal) - nur "KEIN Wert ausgewählt" wird als Hinweis
    markiert.
    """
    print("\n" + "="*55)
    print("✅ VALIDIERUNG: Mehrfachauswahl-Kategorien")
    print("="*55)
    for name, optionen in kategorien:
        gefunden = [label for feld, label in optionen if data.get(feld) == "r"]
        if gefunden:
            werte = ", ".join(f'"{g}"' for g in gefunden)
            print(f"  ✅ {name:<32} = {werte}")
        else:
            print(f"  ❗ {name:<32} = KEIN Wert ausgewählt")

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
# Reihenfolge UND Gliederung der Konsolen-Vorschau: entspricht der
# Abschnittsfolge im Template (Deckblatt -> Kapitel 1 -> Kapitel 2 ->
# Rest -> Historie). Jeder Abschnitt bekommt in zeige_datenvorschau()
# eine eigene Überschrift - aber NUR, wenn mindestens eines seiner
# Felder tatsächlich Daten enthält (kein leerer Abschnitt in der
# Ausgabe).
VORSCHAU_ABSCHNITTE = [
    ("Deckblatt – Rollen/Unterschriften", [
        "Ersteller", "SME", "SI/PL", "TSO", "BSO", "BQR", "CSQ",
    ]),
    ("Deckblatt – Identifikation", [
        # MLCS-ID bewusst DIREKT vor dem Systemnamen (AS/BDIS-Name) -
        # beide gehoeren fachlich zusammen ("welches System ist das").
        "MLCSID", "AS/BDIS-Name", "UeberlagerteMLCS", "Dok. -Nr.", "Hyperlink",
        "Erkannte_Version", "Version", "Anlage",
        "API", "Betrieb", "Gebaeude", "BE", "Raum", "PLSTA",
        "DokNummerVorQualiPSO", "Lieferantennummer", "Schnittstelle",
    ]),
    ("Deckblatt – Beschreibung/Hersteller", [
        "Kurzbeschreibung", "SW-Version / Typ:", "SW-Name:",
        "SW-Hersteller", "Hersteller", "Phenix", "SAP",
    ]),
    ("Kapitel 1 – Systembewertung", [
        "Neuerstellung/Änderung",
        "Systemtyp (Zugangsbeschränkung)",
        "Klassifizierung",
        "GxP-Relevanz",
        "Business Kritisch",
    ]),
    ("Kapitel 2 – Zusammenfassungstabelle", [
        "GxP-Kritikalität",
        "GxP_Produktqualitaet", "GxP_Patientensicherheit", "GxP_Datenintegritaet",
        "CS-Typ", "Systemtyp_CE",
        "ERES-Typ",
        "GAMP5 Software-Kategorie",
        "KI-Reifegrad",
        "DI EE-Anforderungen",
        "Gerätekategorie",
        "Periodic Review",
        "Vereinfachte Qualifizierung",
        "Validierung/Qualifizierung nach SOP",
        "Testtiefe", "Testtiefe-Matrix",
        "Besonderheiten",
    ]),
    # Entspricht dem Kapitel "Informationen und Bemerkungen" im
    # Template (Tabelle 16) - inkl. der 4 generischen "BemerkungX"-
    # Spalten der Master-Excel, die laut Fachbereich konkret folgende
    # Bedeutung haben: Bemerkung1=Prozessbeschreibung,
    # Bemerkung2=Daten, Bemerkung3=Audit Trail, Bemerkung4=Parameter
    # (siehe auch FELD_LABELS in webapp/app.py fuer die
    # entsprechend beschrifteten Editor-Felder).
    ("Informationen und Bemerkungen", [
        "Steuerung erfolgt über?", "Prozessbeschreibung", "Daten",
        "Parameter", "Alarme (GxP-relevant)", "Chargenprotokoll",
        "Audit Trail (AT)", "Benutzer-verwaltung?",
        "Schnittstellen mit PLS", "Angeschlossenes Equipment",
        "Sonstiges", "KI Bewertung",
        "Bemerkung1", "Bemerkung2", "Bemerkung3", "Bemerkung4",
    ]),
    ("Dokumentenhistorie", [
        "Bearbeiter", "Datum", "Version_Historie", "Historie",
    ]),
    ("Intern", [
        "Python ja/nein",
    ]),
]

# Flache Reihenfolge (nur die Feld-/Kategorienamen, ohne Überschriften) -
# wird u.a. genutzt, um am Ende unbekannte Felder zu erkennen.
VORSCHAU_REIHENFOLGE = [
    feld for _, felder in VORSCHAU_ABSCHNITTE for feld in felder
]

def zeige_datenvorschau(data, kategorien=None):
    """
    Zeigt alle extrahierten Felder in der Konsole an.

    `kategorien` (typischerweise VALIDATION_KATEGORIEN_VX +
    MEHRFACHAUSWAHL_KATEGORIEN zusammen) fasst alle Felder, die zu
    einer bekannten Kategorie gehören (z.B. GxP-C/GxP-M/GxP-m2/GxP-NA),
    zu EINER Zeile "<Kategorie> = <ausgewähltes Label>" zusammen,
    statt jedes Einzelfeld als r/c-Zeile anzuzeigen. Sind mehrere
    Werte einer Kategorie angekreuzt, werden alle kommasepariert
    angezeigt (z.B. bei den Mehrfachauswahl-Kategorien normal, bei den
    "genau 1 erwartet"-Kategorien ein Hinweis auf einen Konflikt -
    Details dazu liefert der nachfolgende VALIDIERUNG-Block).

    Alle übrigen Felder (Text, Namen, Daten, nicht gruppierte
    Checkboxen) bleiben wie gewohnt einzeln stehen - deren
    Checkbox-Rohwerte ("r"/"c") werden dabei in "ja"/"nein" übersetzt.

    Die Ausgabe ist in Abschnitte gegliedert (VORSCHAU_ABSCHNITTE), die
    der Reihenfolge im Template entsprechen: zuerst die Personen vom
    Deckblatt, dann Deckblatt-Identifikation, Kapitel 1, Kapitel 2 usw.
    Jeder Abschnitt bekommt eine eigene Überschrift - eine Überschrift
    erscheint aber nur, wenn der Abschnitt mindestens ein Feld mit
    tatsächlichem Inhalt enthält (kein leerer Abschnitt in der
    Ausgabe). Felder/Kategorien, die (noch) keinem Abschnitt
    zugeordnet sind, landen alphabetisch sortiert in einem
    abschließenden Abschnitt "Weitere Felder", damit nichts
    stillschweigend verschwindet.
    """
    kategorien = kategorien or []
    gruppierte_felder = set()
    kategorie_werte = {}
    for name, optionen in kategorien:
        gruppierte_felder.update(feld for feld, _ in optionen)
        gefunden = [label for feld, label in optionen if data.get(feld) == "r"]
        kategorie_werte[name] = ", ".join(gefunden) if gefunden else "-"

    print("\n" + "="*55)
    print("📊 EXTRAHIERTE DATEN (Vorschau):")
    print("="*55)

    zeilen = {name: wert for name, wert in kategorie_werte.items()}
    for col, value in data.items():
        if col in gruppierte_felder:
            continue
        if value == "r":
            zeilen[col] = "ja"
        elif value == "c":
            # Systemtyp_CE ist kein eigenstaendig gestelltes Ja/Nein,
            # sondern nur ein abgeleitetes Merkmal ("gehoert CS-Typ zu
            # LCE/PCS/EE?") - bei allen anderen CS-Typen (CIS, S0-S2,
            # N/A) ist die Frage schlicht nicht einschlaegig, "nein"
            # waere hier irrefuehrend.
            zeilen[col] = "-" if col == "Systemtyp_CE" else "nein"
        else:
            zeilen[col] = (
                str(value)[:60] + "..."
                if len(str(value)) > 60
                else str(value)
            )

    unbekannt = sorted(col for col in zeilen if col not in VORSCHAU_REIHENFOLGE)
    abschnitte = VORSCHAU_ABSCHNITTE + [("Weitere Felder", unbekannt)]

    for titel, felder in abschnitte:
        vorhandene = [f for f in felder if f in zeilen]
        if not vorhandene:
            continue
        print(f"\n--- {titel} ---")
        for col in vorhandene:
            print(f"  {col:<35} = {zeilen[col]}")
