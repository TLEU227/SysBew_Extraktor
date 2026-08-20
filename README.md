# SysBew_Extraktor

Extrahiert Daten aus Sanofi-Systembewertungen (Word, `.docx`, Templates
V7/V8/V10/V11) und trägt sie als neue Zeile in die zentrale
SharePoint-Master-Excel (`Systembewertungen_GESAMT.xlsx`, Sheet
`SysBew`, Excel-Tabelle `Tabelle1`) ein.

Diese Datei ist die einzige Dokumentation des Projekts – sie liegt im
Repo **und** wird unverändert mit in den produktiven Ordner kopiert
(siehe [Dateien zum Publizieren](#dateien-zum-publizieren--ordnerstruktur)).

## Benutzung

Word-Datei auf **`word_parser_main.py`** ziehen (Drag & Drop) oder:

```
python word_parser_main.py "Systembewertung.docx"
```

Das ist der **einzige** Einstiegspunkt. Die Template-Version wird
automatisch erkannt und intern an das passende Erweiterungs-Modul
weitergereicht – es gibt keine Auswahl "welches Skript für welche
Version" mehr.

1. Word-Datei ist geschlossen (nicht in Word geöffnet).
2. Word-Datei per Drag & Drop auf `word_parser_main.py` ziehen.
3. Konsolenausgabe verfolgen – jeder Extraktionsschritt wird
   protokolliert, am Ende erscheint entweder `✅ In Master-Excel
   eingefügt (Zeile N)` oder eine Fehlermeldung mit Hinweis zur
   Behebung.
4. Bei "Master-Excel ist geöffnet/gesperrt": Datei in Excel schließen,
   dann ENTER drücken – das Skript versucht es erneut.
5. Bei "pywin32 ist nicht installiert": einmalig `pip install
   pywin32` ausführen.

## Arbeitsweise

1. Das Skript prüft zuerst die Template-Version des Dokuments:
   - **V8** = kein KI-Kapitel im Dokument
   - **V10** = KI-Kapitel vorhanden + Testtiefe als direkte
     Checkboxen/Text in Kapitel 2 ODER N/A bei "Vereinfachte
     Qualifizierung"
   - **V11** = "GAMP 5 (2nd Edition)" im Text ODER Testtiefe nur als
     Matrix in Kapitel 8

   **Sonderfall V7:** strukturell identisch zu V8 (kein KI-Kapitel,
   gleicher Tabellenaufbau) – einziger Unterschied sind SOP-Referenzen
   mit Standort-Präfix (z. B. `FRA-QU-SOP-...` statt `QU-SOP-...`).
   V7-Dokumente werden daher von der V8-Erweiterung verarbeitet. Die
   Spalte `Erkannte_Version` wird trotzdem korrekt mit `V7` befüllt,
   da zusätzlich der im Dokument selbst genannte Basis-Versionshinweis
   ausgelesen wird (`Neuerstellung auf Basis der ... Version 7.0` im
   Text der Tabelle "Version / Freigabedatum") – unabhängig von der
   strukturellen Versionserkennung, die weiterhin V8-Logik nutzt.
   *Bestätigt an den echten Master-Template-PDFs (Versionen 4.0–8.0):
   Versionen 4.0/5.0/6.0 sind die FRA-präfigierte Variante ("V7"),
   Version 8.0 die generische ("V8"); ab Version 5.0 wechselt zudem
   der Unterschriften-Vermerk von "siehe GEODE+" auf "siehe CMS".*

2. Alle relevanten Felder werden aus den Word-Tabellen extrahiert:
   Kopfdaten, Checkboxen (echte Word-Formularfelder **und**
   Klartext-Checkboxen `r`/`c` aus Seriendruck-Dokumenten),
   GxP-Klassifizierung, Klassifizierung (Lokal/Multi-Site/Global,
   inkl. Klasse 1a–3), Periodic Review, DI EE-Anforderungen,
   Testtiefe, Rollen/Namen vom Deckblatt (Ersteller/SME/SI-PL/TSO/
   BSO/BQR/CSQ), Kategorisierung, Historie usw.

3. Ergebnis wird in die Master-Excel geschrieben: neue letzte Zeile
   in `Systembewertungen_GESAMT.xlsx`, Sheet `SysBew`, Excel-Tabelle
   `Tabelle1` – per COM-Automatisierung durch ein im Hintergrund
   gestartetes echtes Excel (kein Neuschreiben der Datei). Dadurch
   bleiben Sensitivity-Label, externe Verknüpfung, Kommentare etc.
   unangetastet.

4. **Duplikat-Hinweis:** Nach dem Einfügen prüft das Skript, ob die
   Kombination aus "Dok. -Nr." und "Version" des neuen Eintrags
   bereits an anderer Stelle in der Master-Excel vorkommt, und gibt
   bei Fund eine Konsolen-Warnung aus. Rein informativ – verhindert
   das Einfügen nicht. Der Platzhalterwert `QU-OPE-xxxxx` wird dabei
   bewusst ignoriert.

5. **Spalte `Python ja/nein`:** wird automatisch mit `ja` befüllt,
   wenn die Zeile über dieses Skript eingetragen wurde. Manuell
   erfasste Zeilen bleiben unangetastet (leer/`nein`).

## Aufbau

| Datei | Zweck |
|---|---|
| `word_parser_main.py` | Einziger Drag-&-Drop-Einstiegspunkt. Erkennt die Template-Version und reicht an das passende Erweiterungs-Modul weiter. Liegt bewusst direkt im (öffentlichen) Ordner, nicht im Unterordner. |
| `lib/sysbew_common.py` | Gemeinsame Basis: Excel-Spalten, Master-Excel-Konfiguration, alle Hilfs- und Extraktionsfunktionen, die nicht vom Template-Aufbau abhängen, sowie das Schreiben in die Master-Excel per COM-Automatisierung. |
| `lib/word_parser_v8.py` | Erweiterung für Template V8 (und V7). |
| `lib/word_parser_v10.py` | Erweiterung für Template V10. |
| `lib/word_parser_v11.py` | Erweiterung für Template V11. |

Die drei Erweiterungs-Module sind reine Bibliotheksmodule ohne eigenes
Drag & Drop mehr – alle Aufrufe laufen ausschließlich über
`word_parser_main.py`. Sie liegen zusammen mit `sysbew_common.py` im
Unterordner `lib/`, damit im (öffentlichen) Hauptordner nur die eine
Datei sichtbar ist, auf die tatsächlich gezogen wird. `word_parser_main.py`
nimmt `lib/` beim Start selbst in den Python-Importpfad auf (`sys.path`)
– das muss beim Kopieren nicht händisch gemacht werden.

## Dateien zum Publizieren / Ordnerstruktur

Für den produktiven Einsatz müssen **`word_parser_main.py`, diese
README.md und der komplette Unterordner `lib/`** in den Zielordner
kopiert werden:

```
!Systembewertungen_CS\00_Serienbrief\
├── Systembewertungen_GESAMT.xlsx     (bereits vorhanden, nicht anfassen)
├── README.md                          <- diese Datei
├── word_parser_main.py                <- Drag & Drop-Ziel
└── lib\
    ├── sysbew_common.py
    ├── word_parser_v8.py
    ├── word_parser_v10.py
    └── word_parser_v11.py
```

Der Ordner ist öffentlich einsehbar - deshalb liegt im Hauptordner nur
`word_parser_main.py` (das eigentliche Drag-&-Drop-Ziel), die
restlichen vier Dateien liegen unauffälliger im Unterordner `lib/`.
Wird `lib/` versehentlich mitgezogen oder umbenannt, meldet
`word_parser_main.py` einen `ModuleNotFoundError` beim Start - dann
prüfen, ob der Ordner `lib/` noch direkt neben `word_parser_main.py`
liegt und exakt so heißt. Die drei
alten Einzelskripte (`word_parser_v8/v10/v11_formularfelder_vX.X.py`)
werden nicht mehr benötigt und sollten beim Update entfernt werden,
damit niemand versehentlich noch das alte, nicht mehr gepflegte
Skript benutzt.

## Voraussetzungen

- Python ≥ 3.8 (getestet mit 3.12)
- `python-docx` (getestet mit 1.2.0), `pywin32` (getestet mit 312)
- Microsoft Excel lokal installiert (Master-Excel-Insert läuft per
  COM-Automatisierung) – nur unter Windows lauffähig

```
pip install python-docx pywin32
```

## Versionshistorie

### Main-Datei + Erweiterungen (aktuelle Architektur)

- Umbau der drei bisher eigenständigen Skripte
  (`word_parser_v8/v10/v11_formularfelder`) in eine gemeinsame Basis
  (`sysbew_common.py`) + drei schlanke Erweiterungsmodule + eine
  einzige Main-Datei (`word_parser_main.py`) mit automatischer
  Versionserkennung.
- Neue Funktion `extract_deckblatt_rollen()`: liest die Namen zu
  Ersteller/SME/SI-PL/TSO/BSO/BQR/CSQ von der Deckblatt-
  Unterschriftentabelle. "Projektleiter/SME" ist im Dokument teils
  eine kombinierte Rolle (Name landet dann in beiden Spalten SI/PL
  und SME – erkannt auch, wenn das Label nur "SME" lautet, aber der
  Bestätigungstext der Folgezeile "Projektleiter/SME" nennt).
  Unbekannte Funktionsbezeichnungen fallen auf SME zurück, außer sie
  deuten auf Projektleitung hin (→ SI/PL). Mehrfachnennungen (z. B.
  zwei BSO) werden mit Zeilenumbruch in einer Zelle zusammengeführt.
- Bugfix (bereits im Vorgänger-Code vorhanden, jetzt behoben): die
  Zelle "Ersteller"/"Autor" wurde von `extract_text_fields()` über
  einen zu laschen Teilstring-Vergleich fälschlich als Treffer für
  "Hersteller / SW-Ersteller / Lieferant" gewertet und übernahm die
  Nachbarzelle als Hersteller-Wert. Die Deckblatt-Rollentabelle wird
  jetzt davon ausgeschlossen.
- Drei bisher komplett unextrahierte Bereiche ergänzt (gegen echte
  Dokumente verifiziert):
  - **DI EE-Anforderungen** (P1–P4/N/A) – war in der
    Zusammenfassungstabelle Kapitel 2 zwischen Testtiefe und
    Gerätekategorie schlicht übersprungen worden.
  - **Periodic Review gemäß** (QU-SOP-0007359 / freie Angabe /
    zyklische Requalifizierung) – eigene Zeile in derselben Tabelle,
    bisher komplett ungemappt.
  - **Klassifizierung** aus Kapitel 1 (Lokales CS / Multi-Site-CS
    inkl. nur lokal/lokal und global / Globales CS inkl. Klasse
    1a/1b/2/3 / Equipment ohne CS).
  - Neue Spalte `Python ja/nein`: wird automatisch mit `ja` befüllt.
- Spaltenname `SI` auf `SI/PL` korrigiert (entsprach nicht dem
  echten Master-Excel-Spaltenkopf).
- Konsolen-Vorschau (`📊 EXTRAHIERTE DATEN`) überarbeitet:
  - Rohe Checkbox-Werte `r`/`c` werden als `ja`/`nein` angezeigt.
  - Gruppierbare Checkbox-Felder (z. B. GxP-Kritikalität,
    Klassifizierung, DI EE-Anforderungen, ERES-Typ, ...) erscheinen
    als eine Zeile mit dem/den ausgewählten Wert(en) statt als
    Einzelfelder.
  - Reihenfolge und Gliederung folgen jetzt dem Template: die
    Ausgabe ist in Abschnitte unterteilt (Deckblatt – Rollen,
    Deckblatt – Identifikation, Deckblatt – Beschreibung/Hersteller,
    Kapitel 1, Kapitel 2 – Zusammenfassungstabelle, Kapitel 2 –
    Systembeschreibung, Sonstiges, Dokumentenhistorie), jeweils mit
    eigener Überschrift statt einer alphabetischen Gesamtliste. Ein
    Abschnitt erscheint nur, wenn er tatsächlich befüllte Felder
    enthält.

### word_parser_v10_formularfelder (Vorgänger-Skript, bis 1.8)

- **1.0** Erste stabile Version. Text-Checkbox-Erkennung (Klartext
  `r`/`c` aus Seriendruck-Dokumenten) als Fallback zu echten
  Word-Formularfeld-Checkboxen (SDT). Neue Spalte `Erkannte_Version`.
  Direktes Einfügen als neue Zeile in die Master-Excel statt eigener
  Excel-Datei. Umstellung von openpyxl auf COM-Automatisierung, um
  Master-Excel-Struktur (Excel-Tabelle, Sensitivity-Label, externe
  Verknüpfung) nicht zu beschädigen. MLCS-ID: Präfix "MLCS" wird
  entfernt.
- **1.1** Neue Felder `Version`, `Offen`/`Geschlossen`/`NA`,
  `UeberlagerteMLCS`. Bugfix: Namenskonflikt zwischen Testtiefe-N/A
  und Kapitel-1-Systemtyp-N/A behoben.
- **1.2** Bugfix "Version": höchste Version über ALLE Zeilen der
  Tabelle ermittelt, nicht nur die erste.
- **1.3** Duplikat-Hinweis nach dem Einfügen.
- **1.4** Bugfix MLCSID: Präfix "MLCS ID" wird korrekt entfernt.
- **1.5** Bugfix COM-Lifecycle: `excel.Quit()` nur noch, wenn das
  Skript selbst eine neue Instanz gestartet hat.
- **1.6** Gesundheitscheck für angehängte Excel-Instanz (verwaiste
  Prozesse, "OLE error 0x800a01a8").
- **1.7** Automatische Wiederholung bei transienten COM-Fehlern
  (z. B. RPC_E_DISCONNECTED durch OneDrive/SharePoint-Sync).
- **1.8** Konsolen-Block "VALIDIERUNG: Vollständigkeitsprüfung" für
  11 Checkbox-Kategorien.

`word_parser_v11_formularfelder` (2.0–2.8) und
`word_parser_v8_formularfelder` (1.0–1.10) durchliefen dieselben
Bugfixes/Erweiterungen parallel; V11 berechnet die Testtiefe
zusätzlich aus der Z-Felder-Matrix in Kapitel 8 statt aus einer
direkten Zelle in Kapitel 2, V8 erkennt zusätzlich den Sonderfall V7
über `extract_template_basis_version()`.

## Bekannte Einschränkungen

- Bei der Fehlermeldung "OLE error 0x800a01a8" ("Object required"):
  im Task-Manager nach hängenden, unsichtbaren EXCEL.EXE-Prozessen
  suchen und diese beenden, dann erneut versuchen. Das Skript prüft
  dies zwar automatisch vor dem Schreibvorgang, ein manueller Check
  schadet im Zweifel aber nicht.
- Die Vollständigkeitsprüfung (VALIDIERUNG-Block) zeigt bei
  Dokumenten mit "GxP-Relevanz = Nein" für alle nachfolgenden
  Kategorien "❗ KEIN Wert ausgewählt" an, da die Systembewertung laut
  Formularlogik dort vorzeitig endet. Das ist KEIN Fehler, sondern
  erwartet.
- Testtiefe (Gering/Mittel/Hoch) und Validierung/Qualifizierung nach
  SOP (QUAL/VAL) sind bewusst NICHT Teil der Vollständigkeitsprüfung,
  da beide Gruppen aktuell nicht alle im Dokument vorhandenen
  Checkbox-Optionen vollständig auf Excel-Spalten abbilden (fehlende
  Testtiefe-N/A-Spalte, fehlende 3. SOP-Spalte QU-SOP-0028559 bei
  Validierung/Qualifizierung) – eine Prüfung würde dort zu falschen
  Warnungen führen.
- Externe Verknüpfungen in der Master-Excel werden beim Öffnen per
  COM bewusst NICHT automatisch aktualisiert (`UpdateLinks=0`), um
  Störungen durch Verknüpfungs-Dialoge zu vermeiden.
- Die Deckblatt-Rollen- und Klassifizierungs-Extraktion wurde gegen
  echte V11-Dokumente verifiziert; für V8/V10 liegen aktuell nur die
  leeren Master-Template-PDFs zum Strukturabgleich vor, noch keine
  echten ausgefüllten Dokumente zum End-to-End-Test.
