# SysBew_Extraktor

Extrahiert Daten aus Sanofi-Systembewertungen (Word, `.docx`, Templates
V7/V8/V10/V11) und trägt sie als neue Zeile in die zentrale
SharePoint-Master-Excel (`Systembewertungen_GESAMT.xlsx`, Sheet
`SysBew`) ein.

## Benutzung

Word-Datei auf **`word_parser_main.py`** ziehen (Drag & Drop) oder:

```
python word_parser_main.py "Systembewertung.docx"
```

Die Template-Version wird automatisch erkannt; es gibt nur noch diesen
einen Einstiegspunkt.

Ausführliche Bedienungsanleitung, Voraussetzungen und Versionshistorie:
siehe [`help.txt`](help.txt) (liegt im produktiven Einsatz zusätzlich
neben der Master-Excel unter `!Systembewertungen_CS\00_Serienbrief\`).

## Aufbau

| Datei | Zweck |
|---|---|
| `word_parser_main.py` | Einziger Drag-&-Drop-Einstiegspunkt. Erkennt die Template-Version und reicht an das passende Erweiterungs-Modul weiter. |
| `sysbew_common.py` | Gemeinsame Basis: Excel-Spalten, Master-Excel-Konfiguration, alle Hilfs- und Extraktionsfunktionen, die nicht vom Template-Aufbau abhängen, sowie das Schreiben in die Master-Excel per COM-Automatisierung. |
| `word_parser_v8.py` | Erweiterung für Template V8 (und V7, siehe `help.txt`) – kein KI-Kapitel. |
| `word_parser_v10.py` | Erweiterung für Template V10 – KI-Kapitel, Testtiefe direkt in Kapitel 2. |
| `word_parser_v11.py` | Erweiterung für Template V11 – KI-Kapitel, Testtiefe als Matrix in Kapitel 8. |
| `help.txt` | Ausführliche Anwender-Dokumentation für den produktiven Einsatz neben der Master-Excel. |

Die drei Erweiterungs-Module sind reine Bibliotheksmodule ohne eigenes
Drag & Drop mehr – alle Aufrufe laufen ausschließlich über
`word_parser_main.py`.

## Voraussetzungen

- Python ≥ 3.8
- `python-docx`, `pywin32`
- Microsoft Excel lokal installiert (Master-Excel-Insert läuft per
  COM-Automatisierung) – nur unter Windows lauffähig

```
pip install python-docx pywin32
```
