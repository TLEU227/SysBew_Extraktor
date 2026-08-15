# ============================================================
# word_parser_v11.py
# V11-Erweiterung des Systembewertung-Extraktors - 3.0
#
# Enthält NUR das, was für Template V11 spezifisch ist. Alles
# Gemeinsame steht in sysbew_common.py.
#
# Kein eigenständiges Drag & Drop mehr (kein __main__-Block) -
# wird ausschließlich von word_parser_main.py importiert.
#
# Versionshistorie vor dem Umbau in Erweiterungs-Module: siehe
# help.txt, Abschnitt 4 (dort als word_parser_v11_formularfelder
# bis 2.8 protokolliert). Diese Datei setzt strukturell bei 3.0 neu
# auf (Umbau auf Main-Datei + Erweiterungen), Extraktionslogik/
# Verhalten ist unverändert zu 2.8.
# ============================================================

import re
import os

from sysbew_common import (
    get_cell_text,
    get_checkboxes_from_cell,
    berechne_systemtyp_ce,
    extract_mlcs_id,
    extract_ueberlagertes_mlcs,
    extract_systemtyp_zugang,
    extract_klassifizierung,
    extract_version_freigabedatum,
    extract_anlage,
    extract_schnittstelle,
    extract_besonderheiten,
    extract_begruendung_gxp,
    extract_neuerstellung,
    extract_gxp_relevan_bc,
    extract_text_fields,
    extract_checkboxes_formularfelder,
    extract_history,
    extract_deckblatt_rollen,
    parse_betrieb,
    parse_dok_nr,
    parse_sw_version,
    parse_gkat_subtypen,
    parse_gxp_einfluss,
)

# ============================================================
# VALIDIERUNG: Kategorien, bei denen genau EIN "r"-Wert erwartet wird
# ============================================================
# Ausgeschlossen (noch unvollständig getrackt, siehe help.txt):
#   - Validierung/Qualifizierung nach SOP (QUAL/VAL) - fehlende 3. SOP-Spalte
#   (V11 hat keine direkte Testtiefe-Zelle wie V8/V10, sondern die
#   Z-Felder-Matrix - siehe eigene Kategorie unten)
VALIDATION_KATEGORIEN_V11 = [
    ("Neuerstellung/Änderung", [
        ("Neuerstellung", "Neuerstellung"),
        ("Revisioniert", "Änderung – im Einsatz/Aktualisierung"),
    ]),
    ("Systemtyp (Zugangsbeschränkung)", [
        ("Offen", "offen"),
        ("Geschlossen", "geschlossen"),
        ("NA", "N/A"),
    ]),
    ("GxP-Relevanz", [
        ("GxP_Relevan_JA", "Ja"),
        ("GxP_Relevan_NEIN", "Nein"),
    ]),
    ("Business Kritisch", [
        ("BCkritisch", "Ja"),
        ("BCunkritisch", "Nein"),
    ]),
    ("GxP-Kritikalität", [
        ("GxP-C", "Critical"),
        ("GxP-M", "Major"),
        ("GxP-m2", "minor"),
        ("GxP-NA", "N/A"),
    ]),
    ("CS-Typ", [
        ("Systemtyp_CIS", "CIS"),
        ("Subtyp_LCE", "CE-LCE"),
        ("Subtyp_PCS", "CE-PCS"),
        ("Subtyp_EE", "CE-EE"),
        ("VNAP_S0", "S0"),
        ("VNAP_S1", "S1"),
        ("VNAP_S2", "S2"),
        ("Subtyp_NA", "N/A"),
    ]),
    ("ERES-Typ", [
        ("ERESTYP1", "Typ 1"),
        ("ERESTYP2", "Typ 2"),
        ("ERESTYP3", "Typ 3"),
        ("ERESTYP4", "Typ 4"),
        ("ERESTYPNA", "N/A"),
    ]),
    ("GAMP5 Software-Kategorie", [
        ("KAT1", "SW-Kat 1"),
        ("KAT3", "SW-Kat 3"),
        ("KAT4", "SW-Kat 4"),
        ("KAT5", "SW-Kat 5"),
        ("KATNA", "N/A"),
    ]),
    ("Gerätekategorie", [
        ("GKATA", "A"),
        ("GKATB", "B"),
        ("GKATC", "C"),
        ("GKATNA", "N/A"),
    ]),
    ("Vereinfachte Qualifizierung", [
        ("VQ", "Ja"),
        ("NVQ", "Nein"),
    ]),
    ("KI-Reifegrad", [
        ("KI1", "I"), ("KI2", "II"), ("KI3", "III"),
        ("KI4", "IV"), ("KI5", "V"), ("KI6", "VI"),
        ("KINA", "N/A"),
    ]),
    ("Testtiefe-Matrix", [
        ("Z1S1", "Z1S1"), ("Z1S2", "Z1S2"), ("Z1S3", "Z1S3"),
        ("Z2S1", "Z2S1"), ("Z2S2", "Z2S2"), ("Z2S3", "Z2S3"),
        ("Z3S1", "Z3S1"), ("Z3S2", "Z3S2"), ("Z3S3", "Z3S3"),
    ]),
]

# ============================================================
# MAPPING: Checkboxen Kapitel 2 (V11: KEINE direkte Testtiefe-Zelle,
# stattdessen Z-Felder-Matrix in Kapitel 8; KI-Kapitel in Zelle 8)
# ============================================================
CHECKBOX_MAPPING_V11 = {
    6: {
        # Periodic Review gemäß (Zusammenfassungstabelle Kapitel 2, Zeile
        # vor der Hauptzeile 6) - an einem echten Dokument bestätigt: die
        # beiden Checkbox-Gruppen liegen (nach Zellverschmelzung) an
        # Zellindex 3 (QU-SOP-0007359 / freie Angabe) und 6 (zyklische
        # Requalifizierung).
        4: {
            3: {0: "PR_SOP", 1: "PR_Andere"},
            6: {0: "PR-Zyklisch"},
        },
        6: {
            0: {0: "GxP-C", 1: "GxP-M", 2: "GxP-m2", 3: "GxP-NA"},
            1: {
                0: "Systemtyp_CIS", 1: "Subtyp_LCE", 2: "Subtyp_PCS",
                3: "Subtyp_EE",     4: "VNAP_S0",    5: "VNAP_S1",
                6: "VNAP_S2",       7: "Subtyp_NA",
            },
            2: {0: "ERESTYP1", 1: "ERESTYP2", 2: "ERESTYP3",
                3: "ERESTYP4", 4: "ERESTYPNA"},
            3: {0: "KAT1", 1: "KAT3", 2: "KAT4", 3: "KAT5", 4: "KATNA"},
            5: {0: "EE_P1", 1: "EE_P2", 2: "EE_P3", 3: "EE_P4", 4: "EE_NA"},
            6: {0: "GKATA", 1: "GKATB", 2: "GKATC", 3: "GKATNA"},
            7: {0: "QUAL", 1: "VAL"},
            8: {0: "KI1", 1: "KI2", 2: "KI3", 3: "KI4",
                4: "KI5", 5: "KI6", 6: "KINA"},
        },
        8: {
            0: {0: "BCkritisch", 1: "BCunkritisch"},
            6: {0: "VQ", 1: "NVQ"},
        },
    },
}

# ============================================================
# V11-spezifische Ableitungsfunktion (Herstellerfeld) - unterscheidet
# sich von V8/V10 durch eine zusätzliche Bereinigung der abschließenden
# QualiPSO-ID im SW-Hersteller-Text.
# ============================================================
def parse_hersteller(raw):
    result = {}
    if not raw:
        return result

    # SW-Hersteller aufbauen: schrittweise bereinigen
    sw = raw
    sw = re.sub(r'\s*\n\s*OOZ[A-Z0-9]*', '', sw).strip()
    sw = re.sub(r'\s*OOZ[A-Z0-9]+\s*', ' ', sw).strip()
    sw = re.sub(
        r'(PU-Lieferant|Systemsoftware|Lieferant)\s*:\s*',
        '', sw, flags=re.IGNORECASE
    ).strip()
    sw = re.sub(
        r'\s*/\s*QTP-Customer\s*ID\s*:.*$', '', sw,
        flags=re.IGNORECASE
    ).strip()
    sw = re.sub(
        r'\s*QualiPSO[-\s]*ID\s*:?\s*[A-Z0-9]*\s*$', '', sw,
        flags=re.IGNORECASE
    ).strip()
    result["SW-Hersteller"] = sw

    # Hersteller: Prefix entfernen und nur Firmennamen behalten
    hersteller_raw = re.split(
        r'[/,]|QualiPSO|QTP|Phenix|Ursprünglich|OOZ', raw, maxsplit=1
    )[0]
    hersteller_raw = re.sub(
        r'^(PU-Lieferant|Systemsoftware|Lieferant|Hersteller)\s*:\s*',
        '', hersteller_raw, flags=re.IGNORECASE
    ).strip().rstrip(":")
    result["Hersteller"] = hersteller_raw.strip()

    # QualiPSO-ID oder QTP-Customer ID → Lieferantennummer
    m = re.search(
        r'(?:QualiPSO|QTP)[-\s]*(?:Customer\s*)?ID\s*:?\s*([A-Z0-9]+)',
        raw, re.IGNORECASE
    )
    if m:
        result["Lieferantennummer"] = m.group(1).strip()

    # Phenix-ID
    m = re.search(r'Phenix[-\s]*ID\s*:?\s*(?:ID\s*)?([0-9]+)',
                  raw, re.IGNORECASE)
    if m:
        result["Phenix"] = m.group(1).strip()
    else:
        # OOZ-Nummer direkt (ohne Prefix) → Phenix
        m = re.search(r'\b(OOZ[A-Z0-9]+)\b', raw, re.IGNORECASE)
        if m and "Lieferantennummer" not in result:
            result["Phenix"] = m.group(1).strip()

    return result

# ============================================================
# NUR IN V11: Testtiefe aus der Z-Felder-Matrix in Kapitel 8
# ============================================================
def extract_z_felder(doc):
    """
    Z-Felder aus nested Testtiefe-Tabelle in Kapitel 8.
    Struktur: Tabelle 12 (8.0/8.1) enthält nested table mit Matrix.
              Kat1+3   Kat4    Kat5
    Critical  Z1S1     Z2S1    Z3S1
    Major     Z1S2     Z2S2    Z3S2
    Minor     Z1S3     Z2S3    Z3S3
    """
    z_label_mapping = {
        "critical": ("Z1S1", "Z2S1", "Z3S1"),
        "major":    ("Z1S2", "Z2S2", "Z3S2"),
        "minor":    ("Z1S3", "Z2S3", "Z3S3"),
    }

    def process_matrix_table(table):
        """Verarbeite die eigentliche 5×5 Matrix-Tabelle."""
        result = {}
        for row in table.rows:
            cells = row.cells
            if len(cells) < 2:
                continue
            label = get_cell_text(cells[0]).lower().strip()
            for key, (col1, col2, col3) in z_label_mapping.items():
                if label == key:
                    row_states = []
                    for cell in cells[1:4]:  # Spalten 1,2,3
                        cbs = get_checkboxes_from_cell(cell)
                        if cbs:
                            row_states.append(cbs[0]["state"])
                    if len(row_states) >= 3:
                        result[col1] = row_states[0]
                        result[col2] = row_states[1]
                        result[col3] = row_states[2]
                    break
        return result

    def find_nested_tables(table):
        """Sammle alle nested tables aus einer Tabelle."""
        nested = []
        for row in table.rows:
            for cell in row.cells:
                for nt in cell.tables:
                    nested.append(nt)
        return nested

    # Suche in allen Tabellen nach der Testtiefe-Matrix
    for table in doc.tables:
        # Prüfe ob diese Tabelle Kap. 8 ist (enthält "8.0" oder "8.1")
        table_text = ""
        for row in table.rows:
            for cell in row.cells:
                table_text += get_cell_text(cell).lower() + " "

        if "8.0" not in table_text and "festlegung der testtiefe" not in table_text:
            continue

        # Suche nested tables in dieser Tabelle
        for nested in find_nested_tables(table):
            nested_text = ""
            for row in nested.rows:
                for cell in row.cells:
                    nested_text += get_cell_text(cell).lower() + " "

            # Ist es die Matrix? (hat critical, major, minor)
            if all(k in nested_text for k in ["critical", "major", "minor"]):
                result = process_matrix_table(nested)
                if result:
                    return result

    return {}

def berechne_testtiefe(data):
    hoch = ["Z2S1", "Z3S1", "Z3S2"]
    mittel = ["Z1S1", "Z2S2", "Z3S3"]
    gering = ["Z1S2", "Z1S3", "Z2S3"]
    for f in hoch:
        if data.get(f) == "r":
            return {"TTIEFEHOCH": "r", "TTIEFEMITTEL": "c", "TTIEFENIEDRIG": "c"}
    for f in mittel:
        if data.get(f) == "r":
            return {"TTIEFEHOCH": "c", "TTIEFEMITTEL": "r", "TTIEFENIEDRIG": "c"}
    for f in gering:
        if data.get(f) == "r":
            return {"TTIEFEHOCH": "c", "TTIEFEMITTEL": "c", "TTIEFENIEDRIG": "r"}
    return {"TTIEFEHOCH": "c", "TTIEFEMITTEL": "c", "TTIEFENIEDRIG": "c"}

# ============================================================
# V11-HAUPTFUNKTION
# ============================================================
def parse_systembewertung_v11(doc, docx_path):
    """
    Extrahiert alle Felder eines bereits erkannten V11-Dokuments.
    `doc` ist ein bereits von word_parser_main.py geöffnetes
    docx.Document-Objekt, `docx_path` nur für Dok.-Nr.-Ableitung
    aus dem Dateinamen.
    """
    data = {"Erkannte_Version": "V11"}

    print("  → Extrahiere MLCS-ID...")
    mlcs_id = extract_mlcs_id(doc)
    if mlcs_id:
        data["MLCSID"] = mlcs_id

    print("  → Extrahiere Anlage/Equipment-Nr...")
    anlage = extract_anlage(doc)
    if anlage:
        data["Anlage"] = anlage

    print("  → Extrahiere Schnittstelle...")
    schnittstelle = extract_schnittstelle(doc)
    if schnittstelle:
        data["Schnittstelle"] = schnittstelle

    print("  → Extrahiere überlagertes MLCS...")
    ueberlagertes_mlcs = extract_ueberlagertes_mlcs(doc)
    if ueberlagertes_mlcs:
        data["UeberlagerteMLCS"] = ueberlagertes_mlcs

    print("  → Extrahiere Offen/Geschlossen/N-A (Kapitel 1 Systemtyp)...")
    data.update(extract_systemtyp_zugang(doc))

    print("  → Extrahiere Klassifizierung (Lokal/Multi-Site/Global)...")
    data.update(extract_klassifizierung(doc))

    print("  → Extrahiere Textfelder...")
    data.update(extract_text_fields(doc))

    print("  → Verarbeite Betrieb / API / Gebäude...")
    betrieb_raw = data.pop("_betrieb_raw", None)
    if betrieb_raw:
        data.update(parse_betrieb(betrieb_raw))

    print("  → Verarbeite Hersteller / Lieferant...")
    hersteller_raw = data.pop("_hersteller_raw", None)
    if hersteller_raw:
        data.update(parse_hersteller(hersteller_raw))

    print("  → Extrahiere Dok.-Nr. aus Dateiname...")
    dok_nr = parse_dok_nr(os.path.basename(docx_path))
    if dok_nr:
        data["Dok. -Nr."] = dok_nr

    print("  → Extrahiere Checkboxen (Kapitel 2)...")
    data.update(extract_checkboxes_formularfelder(doc, CHECKBOX_MAPPING_V11))

    print("  → Extrahiere Neuerstellung/Revisioniert (text-basiert)...")
    data.update(extract_neuerstellung(doc))

    print("  → Extrahiere GxP-Relevanz + BC (text-basiert)...")
    data.update(extract_gxp_relevan_bc(doc))

    print("  → Extrahiere Z-Felder (text-basiert)...")
    data.update(extract_z_felder(doc))

    print("  → Berechne Testtiefe aus Z-Feldern...")
    data.update(berechne_testtiefe(data))

    print("  → Berechne Systemtyp_CE...")
    data["Systemtyp_CE"] = berechne_systemtyp_ce(data)

    print("  → Extrahiere Besonderheiten / SW-Version / GKAT...")
    besonderheiten = extract_besonderheiten(doc)
    if besonderheiten:
        data["Besonderheiten"] = besonderheiten
        data.update(parse_sw_version(besonderheiten))
        data.update(parse_gkat_subtypen(besonderheiten))

    print("  → Analysiere GxP-Begründung...")
    begruendung = extract_begruendung_gxp(doc)
    data.update(parse_gxp_einfluss(begruendung))

    print("  → Extrahiere Dokumentenhistorie...")
    data.update(extract_history(doc))

    print("  → Extrahiere Rollen/Namen vom Deckblatt...")
    data.update(extract_deckblatt_rollen(doc))

    version_freigabe = extract_version_freigabedatum(doc)
    data["Version"] = version_freigabe if version_freigabe else data.get("Version_Historie", "")

    print("  → Setze Bemerkungen...")
    data["Bemerkung1"] = data.get("Prozessbeschreibung", "")
    data["Bemerkung2"] = data.get("Daten", "")
    data["Bemerkung3"] = data.get("Audit Trail (AT)", "")
    data["Bemerkung4"] = data.get("Parameter", "")

    print(f"  ✅ Extraktion abgeschlossen: {len(data)} Felder gefunden")
    return data
