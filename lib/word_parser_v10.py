# ============================================================
# word_parser_v10.py
# V10-Erweiterung des Systembewertung-Extraktors - 3.0
#
# Enthält NUR das, was für Template V10 spezifisch ist. Alles
# Gemeinsame steht in sysbew_common.py.
#
# Kein eigenständiges Drag & Drop mehr (kein __main__-Block) -
# wird ausschließlich von word_parser_main.py importiert.
#
# Versionshistorie vor dem Umbau in Erweiterungs-Module: siehe
# README.md (dort als word_parser_v10_formularfelder bis 1.8
# protokolliert). Diese Datei setzt strukturell bei 2.0 neu auf
# (Umbau auf Main-Datei + Erweiterungen). 3.0: Deckblatt-Rollen,
# Klassifizierung, DI EE-Anforderungen, Periodic Review ergänzt
# (siehe README.md).
# ============================================================

import re
import os

from sysbew_common import (
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
# Ausgeschlossen (noch unvollständig getrackt, siehe README.md):
#   - Testtiefe (Gering/Mittel/Hoch) - fehlende N/A-Spalte
#   - Validierung/Qualifizierung nach SOP (QUAL/VAL) - fehlende 3. SOP-Spalte
VALIDATION_KATEGORIEN_V10 = [
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
]

# ============================================================
# MAPPING: Checkboxen Kapitel 2 (V10: Testtiefe direkt in Zelle 4
# UND KI-Kapitel in Zelle 8, anders als V8/V11)
# ============================================================
CHECKBOX_MAPPING_V10 = {
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
            # V10: Testtiefe direkt in Zelle 4! (kein eigenes NA-Feld in der
            # Master-Excel für Testtiefe-N/A vorhanden -> Index 3 bewusst nicht gemappt,
            # "NA" ist für Kapitel-1 Systemtyp offen/geschlossen/N-A reserviert)
            4: {0: "TTIEFENIEDRIG", 1: "TTIEFEMITTEL", 2: "TTIEFEHOCH"},
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
# V10-spezifische Ableitungsfunktion (Herstellerfeld)
# ============================================================
def parse_hersteller(raw):
    result = {}
    if not raw:
        return result

    # SW-Hersteller: OOZ-Nummern und Zeilenumbrüche entfernen
    # SW-Hersteller: Prefix + OOZ + QTP-Teile entfernen
    sw_hersteller = re.sub(r'\s*\n\s*OOZ[A-Z0-9]*', '', raw).strip()
    sw_hersteller = re.sub(r'\s*OOZ[A-Z0-9]+\s*', ' ', sw_hersteller).strip()
    sw_hersteller = re.sub(
        r'(PU-Lieferant|Systemsoftware|Lieferant)\s*:\s*',
        '', sw_hersteller, flags=re.IGNORECASE
    ).strip()
    sw_hersteller = re.sub(
        r'\s*/\s*QTP-Customer\s*ID\s*:.*$', '', sw_hersteller,
        flags=re.IGNORECASE
    ).strip()
    result["SW-Hersteller"] = sw_hersteller

    # Hersteller: Prefix wie "PU-Lieferant:", "Systemsoftware:" entfernen
    # und nur den Firmennamen behalten
    hersteller_raw = re.split(
        r'[/,]|QualiPSO|QTP|Phenix|Ursprünglich|OOZ', raw, maxsplit=1
    )[0]
    # Prefix wie "PU-Lieferant: " entfernen
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
# V10-HAUPTFUNKTION
# ============================================================
def parse_systembewertung_v10(doc, docx_path):
    """
    Extrahiert alle Felder eines bereits erkannten V10-Dokuments.
    `doc` ist ein bereits von word_parser_main.py geöffnetes
    docx.Document-Objekt, `docx_path` nur für Dok.-Nr.-Ableitung
    aus dem Dateinamen.
    """
    data = {"Erkannte_Version": "V10"}

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

    print("  → Extrahiere Checkboxen (Kapitel 2 inkl. Testtiefe)...")
    data.update(extract_checkboxes_formularfelder(doc, CHECKBOX_MAPPING_V10))

    print("  → Extrahiere Neuerstellung/Revisioniert (text-basiert)...")
    data.update(extract_neuerstellung(doc))

    print("  → Extrahiere GxP-Relevanz + BC (text-basiert)...")
    data.update(extract_gxp_relevan_bc(doc))

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
