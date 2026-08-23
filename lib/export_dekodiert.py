# ============================================================
# export_dekodiert.py
# Dekodierter Export der Master-Excel fuer externe Tools - 1.0 (Prototyp)
#
# HINTERGRUND: Ein anderes Team baut parallel ein zweites Tool
# ("Fill-a-Masterform") und wollte Systembewertungen_GESAMT.xlsx
# direkt als Startbestand importieren. Problem: ein grosser Teil der
# Spalten (Checkbox-Flags "c"/"r") ergibt erst als GRUPPE einen
# fachlichen Wert (z.B. GxP-C/GxP-M/GxP-m2/GxP-NA -> "GxP-Kritikalitaet:
# Major"). Ein direkter Import wuerde dieses Gruppen-Wissen beim
# externen Tool fest verdrahten und bei jeder Spaltenaenderung hier
# stillschweigend brechen.
#
# Dieses Skript ist ein PROTOTYP fuer eine stabilere, entkoppelte
# Schnittstelle: es liest die Master-Excel (read-only, wie db_reader.py)
# und gibt pro Zeile einen FLACHEN Datensatz zurueck, in dem jede
# Checkbox-GRUPPE zu einem einzigen Klartext-Feld zusammengefasst ist
# (z.B. "gxp_kritikalitaet": "Major") - alle anderen (bereits
# lesbaren) Spalten bleiben unter ihrem Original-Namen unveraendert.
#
# Bewusst KEIN Live-System/keine laufende API - nur ein bei Bedarf neu
# erzeugter, versionierter Datei-Export (JSON und/oder CSV), wie im
# Kollaborations-Vorschlag angeregt. Rein lesend, schreibt nichts.
#
# Nutzung:
#   python export_dekodiert.py ausgabe.json
#   python export_dekodiert.py ausgabe.json --csv ausgabe.csv
# ============================================================

import argparse
import csv
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import db_reader
import sysbew_common as common
from word_parser_v8 import VALIDATION_KATEGORIEN_V8
from word_parser_v10 import VALIDATION_KATEGORIEN_V10
from word_parser_v11 import VALIDATION_KATEGORIEN_V11

# Menschenlesbare (deutsche) Kategorienamen -> stabile, ASCII-/snake_case
# Export-Schluessel. Namen, die hier NICHT auftauchen (sollte nicht
# vorkommen, siehe _alle_kategorien()), werden 1:1 als Kategoriename
# uebernommen.
KATEGORIEN_EXPORT_NAMEN = {
    "Neuerstellung/Änderung":               "neuerstellung_aenderung",
    "Systemtyp (Zugangsbeschränkung)":      "systemtyp_zugang",
    "GxP-Relevanz":                         "gxp_relevanz",
    "Business Kritisch":                    "business_kritisch",
    "GxP-Kritikalität":                     "gxp_kritikalitaet",
    "CS-Typ":                                "cs_typ",
    "ERES-Typ":                              "eres_typ",
    "GAMP5 Software-Kategorie":             "gamp5_kategorie",
    "Gerätekategorie":                       "geraetekategorie",
    "Vereinfachte Qualifizierung":          "vereinfachte_qualifizierung",
    "KI-Reifegrad":                          "ki_reifegrad",
    "Klassifizierung":                       "klassifizierung",
    "DI EE-Anforderungen":                   "di_ee_anforderungen",
    "Periodic Review":                       "periodic_review",
    "Testtiefe":                             "testtiefe",
    "Validierung/Qualifizierung nach SOP":  "validierung_qualifizierung_sop",
}

def _alle_kategorien():
    """Vereinigung ALLER Checkbox-Kategorien ueber V8/V10/V11 hinweg
    (Feldnamen sind versionsunabhaengig gleich benannt, z.B. bedeutet
    "GxP-M" in jeder Version dasselbe) plus der versionsunabhaengigen
    Mehrfachauswahl-/Anzeige-Kategorien aus sysbew_common.py. Eine
    reine Zeilen-Datenbank wie die Master-Excel hat kein "Erkannte
    Version"-Feld, das zuverlaessig genug waere, um pro Zeile nur eine
    Teilmenge zu waehlen - deshalb einfach alle bekannten Gruppen
    anwenden; nicht zutreffende Felder sind ohnehin leer/"c" und tragen
    nichts zur dekodierten Kategorie bei.

    "Testtiefe-Matrix" (die rohe Z1S1...Z3S3-Matrix) wird bewusst
    NICHT dekodiert - die bereits vorhandene "Testtiefe"-Kategorie
    (Gering/Mittel/Hoch) deckt denselben Sachverhalt bereits in
    Klartext ab.
    """
    kategorien = {}
    for liste in (VALIDATION_KATEGORIEN_V8, VALIDATION_KATEGORIEN_V10, VALIDATION_KATEGORIEN_V11):
        for name, optionen in liste:
            if name == "Testtiefe-Matrix":
                continue
            kategorien.setdefault(name, optionen)
    for name, optionen in common.MEHRFACHAUSWAHL_KATEGORIEN:
        kategorien.setdefault(name, optionen)
    for name, optionen in common.ANZEIGE_ZUSATZ_KATEGORIEN:
        kategorien.setdefault(name, optionen)
    return kategorien

def zeile_dekodieren(row):
    """Baut aus einer rohen ML-Zeile (Dict {Spaltenname: Wert}, siehe
    db_reader.read_master_rows()) einen dekodierten Datensatz: jede
    Checkbox-Kategorie wird zu EINEM Feld mit den angekreuzten
    Klartext-Label(s) (mehrere durch "; " getrennt, falls tatsaechlich
    mehrere angekreuzt sind - bei einer "genau 1 erwartet"-Kategorie
    ist das ein Hinweis auf einen Dateneingabe-Konflikt und wird
    bewusst NICHT verschleiert). Kein Treffer -> None. Alle anderen
    Spalten (bereits Klartext) bleiben unter ihrem Original-Namen
    erhalten."""
    kategorien = _alle_kategorien()
    alle_checkbox_felder = set()
    dekodiert = {}
    for name, optionen in kategorien.items():
        alle_checkbox_felder.update(feld for feld, _ in optionen)
        gefunden = [label for feld, label in optionen if row.get(feld) == "r"]
        export_key = KATEGORIEN_EXPORT_NAMEN.get(name, name)
        dekodiert[export_key] = "; ".join(gefunden) if gefunden else None

    for spalte, wert in row.items():
        if spalte in alle_checkbox_felder or spalte == "_zeile":
            continue
        dekodiert[spalte] = wert
    return dekodiert

def dekodierte_zeilen():
    """Liest die Master-Excel und gibt alle Zeilen dekodiert zurueck
    (Liste von Dicts)."""
    _, rows = db_reader.read_master_rows()
    return [zeile_dekodieren(row) for row in rows]

def exportieren(json_pfad=None, csv_pfad=None):
    """Schreibt den dekodierten Export nach `json_pfad` und/oder
    `csv_pfad` (mindestens einer sollte angegeben werden). Gibt die
    dekodierten Zeilen ausserdem zurueck."""
    zeilen = dekodierte_zeilen()

    if json_pfad:
        with open(json_pfad, "w", encoding="utf-8") as f:
            json.dump(zeilen, f, ensure_ascii=False, indent=2)

    if csv_pfad:
        # Spaltenreihenfolge: alle tatsaechlich vorkommenden Schluessel,
        # in Auftrittsreihenfolge (Zeilen koennen sich in optionalen
        # Feldern unterscheiden, falls die Excel-Struktur mal variiert).
        alle_spalten, gesehen = [], set()
        for zeile in zeilen:
            for spalte in zeile:
                if spalte not in gesehen:
                    gesehen.add(spalte)
                    alle_spalten.append(spalte)
        with open(csv_pfad, "w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=alle_spalten)
            writer.writeheader()
            for zeile in zeilen:
                writer.writerow(zeile)

    return zeilen

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Dekodierter Export der Master-Excel (Klartext statt Checkbox-Flags)."
    )
    parser.add_argument("json_datei", nargs="?", default="systembewertungen_dekodiert.json",
                         help="Zieldatei fuer den JSON-Export (Default: systembewertungen_dekodiert.json)")
    parser.add_argument("--csv", metavar="CSV_DATEI", default=None,
                         help="Zusaetzlich als CSV exportieren")
    args = parser.parse_args()

    zeilen = exportieren(json_pfad=args.json_datei, csv_pfad=args.csv)
    print(f"✅ {len(zeilen)} Zeilen dekodiert exportiert nach: {args.json_datei}")
    if args.csv:
        print(f"   sowie: {args.csv}")
