# ============================================================
# word_parser_main.py
# Systembewertung Extraktor - Main-Datei - 2.1
#
# EINZIGER Drag & Drop-Einstiegspunkt: Word-Datei auf DIESES Skript
# ziehen (oder als Kommandozeilenargument übergeben). Die
# Template-Version (V7/V8/V10/V11) wird automatisch erkannt
# (sysbew_common.detect_template_version) und intern an die
# passende Erweiterung weitergereicht:
#   - word_parser_v8.py   (deckt auch V7 ab, siehe README.md)
#   - word_parser_v10.py
#   - word_parser_v11.py
#
# Die drei Erweiterungs-Module sind reine Bibliotheks-Module ohne
# eigenes Drag & Drop mehr - alle Aufrufe laufen über diese Datei.
#
# Diese Datei ist bewusst die einzige, die direkt im (öffentlichen)
# Ordner neben der Master-Excel liegt - sysbew_common.py und die drei
# Erweiterungs-Module liegen im Unterordner "lib/" (weniger sichtbar/
# weniger versehentliches Anfassen). Der Unterordner wird unten
# manuell in sys.path aufgenommen, da Python diesen sonst nicht
# automatisch durchsucht.
# ============================================================

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "lib"))

from docx import Document

import sysbew_common as common
import word_parser_v8 as v8
import word_parser_v10 as v10
import word_parser_v11 as v11

# Version-Nummer -> (Extraktionsfunktion, Validierungskategorien)
_VERSIONS = {
    8:  (v8.parse_systembewertung_v8,  v8.VALIDATION_KATEGORIEN_V8),
    10: (v10.parse_systembewertung_v10, v10.VALIDATION_KATEGORIEN_V10),
    11: (v11.parse_systembewertung_v11, v11.VALIDATION_KATEGORIEN_V11),
}

def parse_systembewertung(docx_path):
    docx_path = common.validiere_docx_datei(docx_path)
    if docx_path is None:
        return {}, None

    print(f"\n📄 Lese Datei: {os.path.basename(docx_path)}")
    doc = Document(docx_path)

    print("  → Prüfe Template-Version...")
    template_version = common.detect_template_version(doc)

    parse_fn, validation_kategorien = _VERSIONS[template_version]
    print(f"  ✅ Template-Version: V{template_version} erkannt")

    data = parse_fn(doc, docx_path)
    if data:
        # Kennzeichnet Zeilen, die über dieses Skript (statt manuell)
        # in die Master-Excel eingetragen wurden.
        data["Python ja/nein"] = "ja"
    return data, validation_kategorien

# ============================================================
# START: Drag & Drop ODER Kommandozeile
# ============================================================
if __name__ == "__main__":

    if len(sys.argv) >= 2:
        docx_path = sys.argv[1]
    else:
        print("="*55)
        print("  Systembewertung Extraktor")
        print("="*55)
        print("  ⚠️  WICHTIG: Word-Datei muss geschlossen sein!")
        print("="*55)
        print("  Tipp: Word-Datei auf dieses Script ziehen!")
        print("="*55)
        docx_path = input("\n📂 Pfad zur Word-Datei: ").strip().strip('"')

    try:
        data, validation_kategorien = parse_systembewertung(docx_path)

        if not data:
            print("\n  Keine Daten extrahiert — kein Excel erstellt.")
        else:
            common.zeige_datenvorschau(data)
            common.validiere_kategorien(data, validation_kategorien)
            common.validiere_mehrfachauswahl_kategorien(data)

            print("\n  → Füge Zeile in Master-Excel ein...")
            ziel_zeile = common.write_to_master_excel(data, docx_path)
            if ziel_zeile:
                print(f"\n✅ In Master-Excel eingefügt (Zeile {ziel_zeile}):")
                print(f"   {common.MASTER_EXCEL_PFAD}")
            else:
                print("\n❌ Master-Excel konnte nicht aktualisiert werden.")

    except FileNotFoundError as e:
        print(f"\n❌ Fehler: {e}")
    except Exception as e:
        print(f"\n❌ Unerwarteter Fehler: {e}")
        import traceback
        traceback.print_exc()

    print("\n" + "="*55)
    print("  Drücke ENTER zum Beenden...")
    input()
