# ============================================================
# masterform_import.py
# Import des dekodierten "Fill-a-Masterform"-Exports - 1.0
#
# HINTERGRUND: Das andere Team (Projekt "Fill-a-Masterform") hat uns
# einen eigenen, bereits dekodierten Auszug ihrer/unserer Systemdaten
# gegeben (Excel mit Klartext-Spalten wie "gxp_kritikalitaet": "Major"
# statt Checkbox-Flags, siehe Kollaborations-Vorschlag). Dieses Modul
# ist das GEGENSTUECK zu export_dekodiert.py: dort kodieren WIR unsere
# Checkbox-Spalten fuer ANDERE in Klartext, hier dekodieren WIR
# fremden Klartext zurueck in unsere Checkbox-Felder, um daraus eine
# neue Systembewertung (SB) im Web-Editor vorbefuellen zu koennen -
# als DRITTE Startquelle neben "aus ML" und "von Null" (siehe
# webapp/app.py, Routen /masterform/...).
#
# Bewusst NUR LESEND (kein Schreibzugriff, keine laufende Verbindung -
# einfache Datei, die bei Bedarf hochgeladen wird) und bewusst
# TRANSPARENT bei allem, was sich nicht zweifelsfrei einer Checkbox
# zuordnen laesst: nicht erkannte/mehrfach markierte Werte werden NIE
# stillschweigend geraten, sondern als `hinweise` zurueckgegeben, die
# die Aufrufer-Seite (Web-Editor) dem/der Bearbeitenden anzeigen soll -
# dieselbe Philosophie wie bei der Konsolen-Validierung ("MEHRERE
# Werte") und bei export_dekodiert.py.
#
# Schema-Stand (Spalten der Fill-a-Masterform-Datei, Stand der
# Uebergabe): siehe ERWARTETE_SPALTEN unten. Das Schema kann sich beim
# anderen Team aendern - lese_masterform_export() bricht deshalb bei
# fehlenden Spalten NICHT ab, sondern warnt nur (Zeilen bekommen fuer
# fehlende Spalten einfach None, zeile_zu_sysbew_daten() behandelt das
# wie "keine Angabe").
#
# Testdatei/Schema-Referenz: assets/templates_xlsx/
# fill_a_masterform_vorlage.xlsx (erzeugt von
# erzeuge_masterform_vorlage.py DIREKT aus den Mapping-Tabellen unten -
# bei Aenderungen hier also dort einfach neu ausfuehren).
# ============================================================

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl

# ============================================================
# ERWARTETE SPALTEN (Stand der Uebergabe von master_db_bereinigt.xlsx)
# ============================================================
ERWARTETE_SPALTEN = [
    "excel_zeile", "mlcs_id", "bereich", "gebaeude", "dok_version",
    "dok_nummer", "systemname", "anlage", "raum", "kurzbeschreibung",
    "sw_name", "sw_version", "sw_hersteller", "hersteller",
    "lieferantennummer", "gxp_relevant", "gxp_kritikalitaet",
    "systemtyp_cis", "systemtyp_ce", "subtyp", "vnap_stufe",
    "doku_status", "gamp_kategorie", "eres_typ", "testtiefe",
    "zone_stufe", "business_critical", "dokumentart", "geraetekategorie",
    "vq_nvq", "qualifizierung_erforderlich", "validierung_erforderlich",
    "ki_einstufung", "subtyp_mehrfach_markiert",
    "geraetekategorie_mehrfach_markiert",
]

# ============================================================
# EINFACHE 1:1-TEXTFELDER (Fill-a-Masterform-Spalte -> unser Feldname)
# Nur Felder, die tatsaechlich eine eigene Editor-/Template-Zelle
# haben (siehe FELDUEBERSICHT.md) - sw_name/sw_version z.B. bewusst
# NICHT hier, siehe _besonderheiten_ergaenzen().
# ============================================================
PLAIN_FELD_MAPPING = {
    "mlcs_id":            "MLCSID",
    "bereich":            "Betrieb",
    "gebaeude":           "Gebaeude",
    "dok_version":        "Version",
    "dok_nummer":         "Dok. -Nr.",
    "systemname":         "AS/BDIS-Name",
    "anlage":             "Anlage",
    "raum":                "Raum",
    "kurzbeschreibung":   "Kurzbeschreibung",
    "sw_hersteller":      "SW-Hersteller",
    "hersteller":         "Hersteller",
    "lieferantennummer":  "Lieferantennummer",
}

# ============================================================
# KATEGORIE-MAPPINGS (Klartext-Wert, kleingeschrieben verglichen -> unser Checkbox-Feld)
# ============================================================
GXP_RELEVANZ      = {"ja": "GxP_Relevan_JA", "nein": "GxP_Relevan_NEIN"}
GXP_KRITIKALITAET = {"critical": "GxP-C", "major": "GxP-M", "minor": "GxP-m2", "n/a": "GxP-NA"}
BUSINESS_CRITICAL = {"ja": "BCkritisch", "nein": "BCunkritisch"}
ERES_TYP          = {"typ 1": "ERESTYP1", "typ 2": "ERESTYP2", "typ 3": "ERESTYP3",
                      "typ 4": "ERESTYP4", "n/a": "ERESTYPNA"}
GAMP_KATEGORIE    = {"1": "KAT1", "3": "KAT3", "4": "KAT4", "5": "KAT5", "n/a": "KATNA"}
# Hauptkategorie A/B/C/N/A direkt; B1/B2/B3/C1/C2/C3 sind bei uns keine
# eigenen Checkboxen (siehe FELDUEBERSICHT.md) - werden auf die
# jeweilige Oberkategorie abgebildet, die Sub-Angabe selbst wandert
# zusaetzlich (siehe _geraetekategorie_decodieren) in "Besonderheiten".
GERAETEKATEGORIE  = {
    "a": "GKATA", "b": "GKATB", "c": "GKATC", "n/a": "GKATNA",
    "b1": "GKATB", "b2": "GKATB", "b3": "GKATB",
    "c1": "GKATC", "c2": "GKATC", "c3": "GKATC",
}
# Fill-a-Masterform nennt hier den Checkbox-Feldnamen selbst
# ("VQ"/"NVQ"), nicht "Ja"/"Nein".
VQ_NVQ            = {"vq": "VQ", "nvq": "NVQ"}
KI_EINSTUFUNG     = {"i": "KI1", "ii": "KI2", "iii": "KI3", "iv": "KI4", "v": "KI5",
                      "vi": "KI6", "n/a": "KINA"}
# Fill-a-Masterform nennt die Werte offenbar nach dem Checkbox-Feldnamen
# selbst (TTIEFENIEDRIG -> "niedrig") statt nach unserem Anzeige-Label
# ("Gering") - deshalb eigene Tabelle statt Wiederverwendung von
# sysbew_common.ANZEIGE_ZUSATZ_KATEGORIEN.
TESTTIEFE         = {"niedrig": "TTIEFENIEDRIG", "mittel": "TTIEFEMITTEL", "hoch": "TTIEFEHOCH"}
DOKUMENTART       = {"neuerstellung": "Neuerstellung", "revisioniert": "Revisioniert",
                      "änderung": "Revisioniert", "aenderung": "Revisioniert"}
SUBTYP            = {"pcs": "Subtyp_PCS", "lce": "Subtyp_LCE", "ee": "Subtyp_EE", "n/a": "Subtyp_NA"}
VNAP_STUFE        = {"s0": "VNAP_S0", "s1": "VNAP_S1", "s2": "VNAP_S2"}
ZONE_FELDER       = ["Z1S1", "Z1S2", "Z1S3", "Z2S1", "Z2S2", "Z2S3", "Z3S1", "Z3S2", "Z3S3"]

def _kategorie_decodieren(data, alle_felder, mapping, rohwert, hinweise, kategorie_name):
    """Setzt alle Felder der Kategorie erst auf "c" (wie eine echte
    ML-Zeile), dann - falls der rohe Klartext-Wert bekannt ist - genau
    ein Feld auf "r". Unbekannte/leere Werte setzen NICHTS auf "r",
    sondern (bei unbekanntem, nicht-leerem Wert) einen Hinweis, statt
    stillschweigend zu raten."""
    for f in alle_felder:
        data[f] = "c"
    if rohwert in (None, ""):
        return
    treffer = mapping.get(str(rohwert).strip().lower())
    if treffer:
        data[treffer] = "r"
    else:
        hinweise.append(
            f"Unbekannter Wert '{rohwert}' für {kategorie_name} - keine "
            f"Checkbox automatisch gesetzt, bitte im Editor prüfen/nachtragen."
        )

def _cs_typ_decodieren(row, data, hinweise):
    """"CS-Typ" ist im Template EIN Kategorie-Feld (Systemtyp_CIS /
    Subtyp_LCE/PCS/EE / VNAP_S0-S2 / Subtyp_NA - genau 1 erwartet),
    das Fill-a-Masterform in mehrere Spalten zerlegt hat
    (systemtyp_cis, subtyp, vnap_stufe). Wird hier wieder
    zusammengefuehrt."""
    alle = ["Systemtyp_CIS", "Subtyp_LCE", "Subtyp_PCS", "Subtyp_EE",
            "VNAP_S0", "VNAP_S1", "VNAP_S2", "Subtyp_NA"]
    for f in alle:
        data[f] = "c"
    if row.get("subtyp_mehrfach_markiert"):
        hinweise.append(
            "CS-Typ (Systemtyp/Subtyp/VNAP-Stufe): im Fill-a-Masterform-"
            "Datensatz als mehrfach markiert gekennzeichnet - keine "
            "Checkbox automatisch gesetzt, bitte Kapitel 1/2 manuell prüfen."
        )
        return
    gesetzt = []
    if row.get("systemtyp_cis") is True:
        gesetzt.append("Systemtyp_CIS")
    subtyp = str(row.get("subtyp") or "").strip().lower()
    if subtyp:
        treffer = SUBTYP.get(subtyp)
        if treffer:
            gesetzt.append(treffer)
        else:
            hinweise.append(f"Unbekannter subtyp-Wert '{row.get('subtyp')}' - nicht auf CS-Typ abgebildet.")
    vnap = str(row.get("vnap_stufe") or "").strip().lower()
    if vnap:
        treffer = VNAP_STUFE.get(vnap)
        if treffer:
            gesetzt.append(treffer)
        else:
            hinweise.append(f"Unbekannter vnap_stufe-Wert '{row.get('vnap_stufe')}' - nicht auf CS-Typ abgebildet.")
    if not gesetzt and row.get("systemtyp_cis") is False and not subtyp and not vnap:
        gesetzt.append("Subtyp_NA")
    for f in gesetzt:
        data[f] = "r"
    if len(gesetzt) > 1:
        hinweise.append(
            f"CS-Typ: mehrere Checkboxen gleichzeitig gesetzt ({', '.join(gesetzt)}) - "
            f"im Template ist eigentlich genau 1 vorgesehen, bitte prüfen."
        )

def _geraetekategorie_decodieren(row, data, hinweise):
    for f in ("GKATA", "GKATB", "GKATC", "GKATNA"):
        data[f] = "c"
    if row.get("geraetekategorie_mehrfach_markiert"):
        hinweise.append(
            "Gerätekategorie: im Fill-a-Masterform-Datensatz als mehrfach "
            "markiert gekennzeichnet - keine Checkbox automatisch gesetzt, "
            "bitte Kapitel 2 manuell prüfen."
        )
        return
    roh = row.get("geraetekategorie")
    wert = str(roh or "").strip().lower()
    if not wert:
        return
    treffer = GERAETEKATEGORIE.get(wert)
    if treffer:
        data[treffer] = "r"
        # B1/B2/B3/C1/C2/C3 sind nur auf die Oberkategorie (B/C)
        # abgebildet - die konkrete Sub-Angabe selbst nicht verlieren.
        if wert not in ("a", "b", "c", "n/a"):
            bestehend = data.get("Besonderheiten") or ""
            zusatz = f"Gerätekategorie-Detail (Fill-a-Masterform): {roh}"
            data["Besonderheiten"] = f"{bestehend}\n{zusatz}" if bestehend else zusatz
    else:
        hinweise.append(f"Unbekannter geraetekategorie-Wert '{roh}'.")

def _zone_stufe_decodieren(data, wert, hinweise):
    for f in ZONE_FELDER:
        data[f] = "c"
    if not wert:
        return
    wert = str(wert).strip().upper()
    if wert in ZONE_FELDER:
        data[wert] = "r"
    else:
        hinweise.append(f"Unbekannte zone_stufe '{wert}' - keine Testtiefe-Matrix-Checkbox gesetzt.")

def _besonderheiten_ergaenzen(row, data, hinweise):
    """sw_name/sw_version haben keine eigene Editor-/Template-Zelle
    (werden bei uns nur RUeCKWaeRTS per Regex aus "Besonderheiten"
    abgeleitet, siehe FELDUEBERSICHT.md) - deshalb hier informativ ins
    Freitextfeld "Besonderheiten" geschrieben (dort frei editierbar,
    nichts geht verloren).

    doku_status/qualifizierung_erforderlich/validierung_erforderlich
    landen bewusst NICHT im Dokument selbst (Besonderheiten wird 1:1
    in die Zusammenfassungstabelle Kapitel 2 uebernommen) - ihr Bezug
    zu einer unserer Checkboxen ist nicht zweifelsfrei genug, um sie
    automatisch zu setzen, UND es sind reine Prozess-/Status-Metadaten
    ohne erkennbaren fachlichen Mehrwert im Dokumenttext selbst (in der
    Beispieldatei praktisch bei jeder Zeile identisch). Sie werden
    stattdessen nur als Hinweis fuer die/den Bearbeitende(n) im
    Web-Editor angezeigt, damit nichts unbemerkt verloren geht, ohne
    das erzeugte Dokument mit Wiederholtext zu fuellen."""
    teile = []
    if row.get("sw_name"):
        teile.append(f"SW-Name (Fill-a-Masterform): {row['sw_name']}")
    if row.get("sw_version"):
        teile.append(f"SW-Version (Fill-a-Masterform): {row['sw_version']}")
    if teile:
        bestehend = data.get("Besonderheiten") or ""
        zusatz = "\n".join(teile)
        data["Besonderheiten"] = f"{bestehend}\n{zusatz}" if bestehend else zusatz

    for feld, label in (
        ("doku_status", "Doku-Status"),
        ("qualifizierung_erforderlich", "Qualifizierung erforderlich"),
        ("validierung_erforderlich", "Validierung erforderlich"),
    ):
        wert = row.get(feld)
        if wert not in (None, ""):
            hinweise.append(
                f"'{feld}' aus dem Import nicht automatisch einer Checkbox "
                f"zugeordnet (Bezug unklar) - Wert laut Import: {label} = "
                f"{wert} (nicht ins Dokument übernommen, bitte bei Bedarf "
                f"manuell im Editor ergänzen)."
            )

def zeile_zu_sysbew_daten(row):
    """Dekodiert eine Fill-a-Masterform-Zeile (Dict wie von
    lese_masterform_export() geliefert) in ein data-Dict im selben
    Format wie eine ML-Zeile (Skript-Feldnamen, Checkbox-Spalten als
    "r"/"c") - kann direkt an
    webapp.app._neues_dokument_aus_db_zeile() uebergeben werden, genau
    wie eine echte ML-Zeile.

    Gibt (data, hinweise) zurueck - `hinweise` ist eine Liste von
    Klartext-Warnungen zu allem, was NICHT zweifelsfrei uebernommen
    werden konnte (unbekannte Werte, als mehrfach markiert
    gekennzeichnete Kategorien, nicht automatisch zuordenbare Felder -
    siehe Modul-Kommentar). Der Web-Editor zeigt diese Hinweise dem/
    der Bearbeitenden an, statt sie zu verschlucken."""
    hinweise = []
    data = {}
    for quelle, ziel in PLAIN_FELD_MAPPING.items():
        wert = row.get(quelle)
        if wert not in (None, ""):
            data[ziel] = str(wert)

    _kategorie_decodieren(data, ["GxP_Relevan_JA", "GxP_Relevan_NEIN"],
                          GXP_RELEVANZ, row.get("gxp_relevant"), hinweise, "GxP-Relevanz")
    _kategorie_decodieren(data, ["GxP-C", "GxP-M", "GxP-m2", "GxP-NA"],
                          GXP_KRITIKALITAET, row.get("gxp_kritikalitaet"), hinweise, "GxP-Kritikalität")
    _kategorie_decodieren(data, ["BCkritisch", "BCunkritisch"],
                          BUSINESS_CRITICAL, row.get("business_critical"), hinweise, "Business Kritisch")
    _kategorie_decodieren(data, ["ERESTYP1", "ERESTYP2", "ERESTYP3", "ERESTYP4", "ERESTYPNA"],
                          ERES_TYP, row.get("eres_typ"), hinweise, "ERES-Typ")
    _kategorie_decodieren(data, ["KAT1", "KAT3", "KAT4", "KAT5", "KATNA"],
                          GAMP_KATEGORIE, row.get("gamp_kategorie"), hinweise, "GAMP5 Software-Kategorie")
    _kategorie_decodieren(data, ["VQ", "NVQ"],
                          VQ_NVQ, row.get("vq_nvq"), hinweise, "Vereinfachte Qualifizierung")
    _kategorie_decodieren(data, ["KI1", "KI2", "KI3", "KI4", "KI5", "KI6", "KINA"],
                          KI_EINSTUFUNG, row.get("ki_einstufung"), hinweise, "KI-Reifegrad")
    _kategorie_decodieren(data, ["TTIEFENIEDRIG", "TTIEFEMITTEL", "TTIEFEHOCH"],
                          TESTTIEFE, row.get("testtiefe"), hinweise, "Testtiefe")
    _kategorie_decodieren(data, ["Neuerstellung", "Revisioniert"],
                          DOKUMENTART, row.get("dokumentart"), hinweise, "Neuerstellung/Änderung")
    _cs_typ_decodieren(row, data, hinweise)
    _geraetekategorie_decodieren(row, data, hinweise)
    _zone_stufe_decodieren(data, row.get("zone_stufe"), hinweise)
    _besonderheiten_ergaenzen(row, data, hinweise)

    return data, hinweise

def lese_masterform_export(pfad):
    """Liest eine hochgeladene Fill-a-Masterform-Datei (erstes Sheet,
    Header in Zeile 1 - im Gegensatz zu db_reader.read_master_rows()
    KEINE benannte Excel-Tabelle notwendig/erwartet). Gibt (header,
    rows) zurueck, `rows` inkl. "_zeile" (1-basiert wie in der
    Quelldatei, fuer Rueckverfolgung).

    Bricht bei abweichendem Schema NICHT ab (das Schema liegt beim
    anderen Team und kann sich aendern) - fehlende erwartete Spalten
    werden nur als Warnung ausgegeben, fehlende Werte pro Zeile sind
    ohnehin schon der Normalfall (None)."""
    wb = openpyxl.load_workbook(pfad, data_only=True, read_only=True)
    try:
        ws = wb.worksheets[0]
        zeilen_iter = ws.iter_rows(values_only=True)
        header = list(next(zeilen_iter))
        fehlend = [s for s in ERWARTETE_SPALTEN if s not in header]
        if fehlend:
            print(
                f"⚠️  Fill-a-Masterform-Import: erwartete Spalten fehlen "
                f"(Schema beim anderen Team evtl. geändert): {', '.join(fehlend)}"
            )
        rows = []
        for excel_row_idx, werte in enumerate(zeilen_iter, start=2):
            if all(w in (None, "") for w in werte):
                continue
            eintrag = dict(zip(header, werte))
            eintrag["_zeile"] = excel_row_idx
            rows.append(eintrag)
        return header, rows
    finally:
        wb.close()

def filter_rows(rows, suchtext=""):
    """Einfache Volltextsuche, analog db_reader.filter_rows()."""
    suchtext = (suchtext or "").strip().lower()
    if not suchtext:
        return rows
    spalten = ("mlcs_id", "systemname", "anlage", "dok_nummer", "gebaeude", "bereich")
    treffer = []
    for row in rows:
        for spalte in spalten:
            wert = row.get(spalte)
            if wert and suchtext in str(wert).lower():
                treffer.append(row)
                break
    return treffer

def get_row_by_index(rows, zeile):
    for row in rows:
        if row.get("_zeile") == zeile:
            return row
    return None
