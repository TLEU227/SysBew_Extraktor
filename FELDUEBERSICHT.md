# Feldübersicht: Master-Excel ↔ Web-Editor ↔ Word-Template

Nachschlagewerk für alle Felder, die im System vorkommen: welche gibt
es, woher kommen sie, und welche davon landen tatsächlich im
automatisch erzeugten Word-Dokument (und wo genau)?

**Kurze Antwort auf "welche DB(s)":** Es gibt genau **eine** echte
Datenbank - die Master-Excel (`Systembewertungen_GESAMT.xlsx`,
Sheet `SysBew`). Ihre komplette Spaltenliste steht in
`lib/sysbew_common.py` als `EXCEL_COLUMNS` (143 Spalten). Zusätzlich
kennt **nur der Web-Editor** (`webapp/app.py`) 17 weitere Felder, die
es in der Master-Excel nicht gibt (siehe Abschnitt "Web-Editor-only"
unten) - kein zweites Datenbank-Schema, nur ein paar zusätzliche
Formularfelder für Angaben, die im Template gebraucht werden, aber in
der Excel nie eine eigene Spalte hatten.

Diese Übersicht wurde automatisch gegen den aktuellen Code geprüft
(nicht nur gegen Kommentare/Erinnerung geschrieben) - siehe
"Wie diese Übersicht erstellt wurde" ganz unten, falls sie mal wieder
aktualisiert werden muss.

## Feldname = Excel-Spaltenname?

**Ja, mit genau einer Ausnahme.** Der interne Feldname (Python-Dict-
Schlüssel, wie in dieser Übersicht verwendet) ist für **alle**
Master-Excel-Felder identisch mit dem tatsächlichen Spaltenkopf in
`Systembewertungen_GESAMT.xlsx` - Groß-/Kleinschreibung und
Sonderzeichen (Leerzeichen, Bindestrich vs. Unterstrich) inklusive.
Einzige Ausnahme (`MASTER_SPALTEN_MAPPING` in `lib/sysbew_common.py`):

| Interner Feldname | Master-Excel-Spaltenkopf |
|---|---|
| `Erkannte_Version` | `Erkannte Version2` |

Die Web-Editor-only-Felder (siehe Abschnitt unten) haben dagegen
**keine** Master-Excel-Spalte - für sie stellt sich die Frage nicht.

## Spaltenerklärung

| Spalte | Bedeutung |
|---|---|
| **Feld** | interner Name = Master-Excel-Spaltenname (falls nicht anders vermerkt) |
| **Im Editor?** | erscheint als Eingabefeld im Web-Editor (`webapp/app.py`)? |
| **Ins Dokument geschrieben?** | liest `lib/template_filler.py` diesen Wert beim Erzeugen einer NEUEN Systembewertung? |
| **Wo im Dokument** | welche Stelle/welches Kapitel im erzeugten .docx |

`✅` = ja · `—` = nein/entfällt

---

## Deckblatt – Rollen/Unterschriften

| Feld | Im Editor? | Ins Dokument geschrieben? | Wo im Dokument |
|---|---|---|---|
| `Ersteller` | ✅ | ✅ | Deckblatt – Unterschriftentabelle |
| `Ersteller_Abteilung` *(Web-Editor-only)* | ✅ | ✅ | Deckblatt – ersetzt "(Site/Unit)" im Label |
| `SME` | ✅ | ✅ | Deckblatt – an SI/PL-Zeile angehängt, falls abweichend |
| `SI/PL` | ✅ | ✅ | Deckblatt – Unterschriftentabelle |
| `SI_PL_Abteilung` *(Web-Editor-only)* | ✅ | ✅ | Deckblatt – ersetzt "(Site/Unit)" im Label |
| `TSO` | ✅ | ✅ | Deckblatt – Unterschriftentabelle |
| `TSO_Abteilung` *(Web-Editor-only)* | ✅ | ✅ | Deckblatt – ersetzt "(Site/Unit)" im Label |
| `BSO` | ✅ | ✅ | Deckblatt – Unterschriftentabelle |
| `BSO_Abteilung` *(Web-Editor-only)* | ✅ | ✅ | Deckblatt – ersetzt "(Site/Unit)" im Label |
| `BQR` | ✅ | ✅ | Deckblatt – Unterschriftentabelle |
| `BQR_Abteilung` *(Web-Editor-only)* | ✅ | ✅ | Deckblatt – ersetzt "(Site/Unit)" im Label |
| `CSQ` | ✅ | ✅ | Deckblatt – Unterschriftentabelle |
| `CSQ_Abteilung` *(Web-Editor-only)* | ✅ | ✅ *(nur falls ausgefüllt)* | Deckblatt – ersetzt das im Label standardmäßig fest eingetragene "(FBC Quality Q&V CSV)"; ohne Eingabe bleibt der Standardtext stehen |

## Deckblatt – Identifikation / Beschreibung / Hersteller

| Feld | Im Editor? | Ins Dokument geschrieben? | Wo im Dokument |
|---|---|---|---|
| `MLCSID` | ✅ | ✅ | Kapitel 2 – Zusammenfassungstabelle |
| `UeberlagerteMLCS` | ✅ | ✅ | Kapitel 2 – Zusammenfassungstabelle (kombiniert mit `Schnittstelle`) |
| `Erkannte_Version` | — (automatisch V11) | — | reine Excel-Metadaten ("Erkannte Version2") |
| `Dok. -Nr.` | ✅ *(kombiniert mit Version, siehe Editor "Dok.-Nr. / Version")* | — | **nicht im Dokumenttext** - nur für Dateiname-Vorschlag + optionalen Master-Excel-Eintrag |
| `Version` | — | — | nur beim LESEN eines bestehenden Dokuments relevant (höchste vorhandene Version); für ein neues Dokument irrelevant |
| `Version_Historie` | ✅ *(Teil von "Dok.-Nr. / Version")* | ✅ | Dokumentenhistorie-Tabelle (Version/Datum) |
| `Datum` | ✅ *(Teil von "Dok.-Nr. / Version")* | ✅ | Dokumentenhistorie-Tabelle |
| `Historie` (Grund der Erstellung/Änderung) | ✅ (eigenes Feld + Textbaustein-Vorschläge) | ✅ | NUR Dokumentenhistorie-Tabelle - Kapitel 1 ("Grund der Systembewertung") bekommt seit dem Praxistest-Feedback stattdessen das eigene, kurze Feld `CCNr_Rahmen` (siehe "Web-Editor-only" unten), da dort laut Template nur CC-Nr. + Rahmen der Erstellung erwartet wird, nicht der lange Historie-Text |
| `Bearbeiter` | — | — | reine Master-Excel-Spalte, keine Zelle im Template |
| `AS/BDIS-Name` | ✅ | ✅ | Kapitel 1 |
| `Anlage` | ✅ | ✅ | Kapitel 2 – Zusammenfassungstabelle |
| `API` | — | — | nur aus `Betrieb` abgeleitete Teilzeichenkette, rein fürs Filtern in der Excel |
| `Betrieb` | ✅ | ✅ | Kapitel 1 (kombiniert mit `Gebaeude` zu "Einsatzbereich") |
| `Gebaeude` | ✅ (Anzeige "Gebäude") | ✅ | Kapitel 1 (kombiniert mit `Betrieb`) |
| `BE` | — | — | reine Master-Excel-Spalte, keine Zelle im Template |
| `Raum` | — | — | reine Master-Excel-Spalte, keine Zelle im Template |
| `PLSTA` | — | — | reine Master-Excel-Spalte (Teilanlagen-Nummer PLS), keine Zelle im Template |
| `DokNummerVorQualiPSO` | ✅ (Anzeige "Vorherige Doc-ID") | — | **nicht im Dokumenttext** - nur Excel-Bookkeeping; wird beim Start aus der DB automatisch mit alter Dok.-Nr./Version befüllt |
| `Lieferantennummer` | ✅ | ✅ | Kapitel 1 (an Hersteller-Feld angehängt: "QualiPSO-ID: ...") |
| `Schnittstelle` | ✅ (Anzeige "Schnittstelle", jetzt Textarea) | ✅ | Kapitel 2 – Zusammenfassungstabelle |
| `Kurzbeschreibung` | ✅ | ✅ | Kapitel 1 |
| `SW-Version / Typ:` | — | — | beim Lesen per Regex aus `Besonderheiten` abgeleitet; keine eigene Zelle im Template |
| `SW-Name:` | — | — | wie oben |
| `SW-Hersteller` | ✅ | ✅ | Kapitel 1 (an Hersteller-Feld angehängt, falls abweichend) |
| `Hersteller` | ✅ | ✅ | Kapitel 1 |
| `Phenix` | — | — | Phenix-Nummern existieren laut Fachbereich nicht mehr |
| `SAP` | — | — | reine Master-Excel-Spalte, keine Zelle im Template |

## Kapitel 1

| Feld | Im Editor? | Ins Dokument geschrieben? | Wo im Dokument |
|---|---|---|---|
| `Neuerstellung` | ✅ (Kategorie "Neuerstellung/Änderung") | ✅ | Kapitel 1 |
| `Revisioniert` | ✅ | ✅ | Kapitel 1 |
| `Offen` | ✅ (Kategorie "Systemtyp (Zugangsbeschränkung)") | ✅ | Kapitel 1 |
| `Geschlossen` | ✅ | ✅ | Kapitel 1 |
| `NA` | ✅ | ✅ | Kapitel 1 |
| `SystemtypZugang_Begruendung` *(Web-Editor-only)* | ✅ (Begründung, optional) | ✅ *(nur falls ausgefüllt)* | Kapitel 1 - als zusätzliche Zeile unter Offen/Geschlossen/N/A angehängt |
| `GxP_Relevan_JA` | ✅ (Kategorie "GxP-Relevanz") | ✅ | nur Kapitel 1 (im Gegensatz zu Business Kritisch hat GxP-Relevanz KEINE eigene Checkbox in der Kapitel-2-Zusammenfassungstabelle) |
| `GxP_Relevan_NEIN` | ✅ | ✅ | nur Kapitel 1 |
| `BCkritisch` | ✅ (Kategorie "Business Kritisch") | ✅ | Kapitel 1 + Kapitel 2 (Checkbox) |
| `BCunkritisch` | ✅ | ✅ | Kapitel 1 + Kapitel 2 (Checkbox) |

### Kapitel 1 – Klassifizierung (Mehrfachauswahl)

| Feld | Im Editor? | Ins Dokument geschrieben? | Wo im Dokument |
|---|---|---|---|
| `KLASS_Lokal` | ✅ | ✅ | Kapitel 1 – Klassifizierung |
| `KLASS_Multisite` | ✅ | ✅ | Kapitel 1 – Klassifizierung |
| `KLASS_Multisite_NurLokal` | ✅ | ✅ | Kapitel 1 – Klassifizierung |
| `KLASS_Multisite_LokalGlobal` | ✅ | ✅ | Kapitel 1 – Klassifizierung |
| `KLASS_Global` | ⚠️ *(nicht direkt, siehe unten)* | ✅ | Kapitel 1 – Klassifizierung **+** steuert die "N/A - Weiter mit Kap. 4"-Checkbox in Kapitel 3 (angekreuzt, wenn NICHT "r" - Kapitel 3 fragt nur die Detailfrage innerhalb von "Globales CS" ab) |
| `KLASS_Global_1a` | ✅ (Kategorie "Globale CS-Klasse (Kapitel 3)") | ✅ | Kapitel 1 – Klassifizierung **+** Kapitel 3 (Detailfrage) |
| `KLASS_Global_1b` | ✅ | ✅ | Kapitel 1 – Klassifizierung **+** Kapitel 3 |
| `KLASS_Global_2` | ✅ | ✅ | Kapitel 1 – Klassifizierung **+** Kapitel 3 |
| `KLASS_Global_3` | ✅ | ✅ | Kapitel 1 – Klassifizierung **+** Kapitel 3 |
| `KLASS_OhneCS` | ✅ | ✅ | Kapitel 1 – Klassifizierung |
| `KLASS_Global_NA` *(Web-Editor-only)* | ✅ (Option "N/A - kein Globales CS") | ❌ | rein UI - markiert im Editor nur, dass keine der 4 Klassen zutrifft; nicht Teil von EXCEL_COLUMNS, wird beim Erzeugen nicht weiterverarbeitet |

`KLASS_Global` selbst hat im Web-Editor **keine eigene Checkbox** mehr:
seit der Kapitel-3-Restrukturierung wird es automatisch gesetzt, sobald
eine der 4 Klassen (`KLASS_Global_1a/1b/2/3`) gewählt wird (siehe
`_dokument_erzeugen_und_senden()`) - eine Klasse ohne "Globales CS" ist
laut Formular ohnehin nicht möglich. `KLASS_Lokal`/`KLASS_Multisite`
(+ Unteroptionen)/`KLASS_OhneCS` bleiben wie bisher eigene Checkboxen
in der Kategorie "Klassifizierung", nur ohne die 4 Klassen und ohne
"Globales CS" selbst (die stehen jetzt in der eigenen Kategorie
"Globale CS-Klasse (Kapitel 3)" direkt darunter). Diese Aufteilung
gilt nur für den Web-Editor (`webapp/app.py` `_KATEGORIEN_ZUSATZ`) -
`sysbew_common.MEHRFACHAUSWAHL_KATEGORIEN["Klassifizierung"]` selbst
(für die Konsistenzprüfung beim Einlesen fertiger Dokumente) enthält
weiterhin alle 10 echten Checkboxen unverändert.

## Kapitel 2 – Zusammenfassungstabelle (Checkboxen)

Alle Felder in diesem Abschnitt laufen über dieselbe generische
Checkbox-Mapping-Funktion (`fill_checkboxes_formularfelder`) und
landen **ausschließlich** in der Zusammenfassungstabelle Kapitel 2 -
außer wo explizit ein zusätzliches Detail-Kapitel vermerkt ist (diese
fragen laut Template dieselbe Information nochmal separat ab und
werden deshalb zusätzlich mitbefüllt).

| Feld | Im Editor? | Zusätzlich noch wo? |
|---|---|---|
| `GxP-C` / `GxP-M` / `GxP-m2` / `GxP-NA` | ✅ (Kategorie "GxP-Kritikalität") | **+** Kapitel 2 Detail (GxP-Risikoklassifizierung); `GxP-C`/`-M`/`-m2` zusätzlich für Testtiefe-Berechnung genutzt |
| `Systemtyp_CIS`, `Subtyp_LCE`, `Subtyp_PCS`, `Subtyp_EE`, `VNAP_S0/S1/S2`, `Subtyp_NA` | ✅ (Kategorie "CS-Typ") | nur Kapitel 2 |
| `ERESTYP1-4`, `ERESTYPNA` | ✅ (Kategorie "ERES-Typ") | **+** Kapitel 6 (Detail); bei Typ 4 zusätzlich `ERES4_SIG_*` (Art der Signatur, Web-Editor-only) |
| `KAT1`, `KAT3`, `KAT4`, `KAT5` | ✅ (Kategorie "GAMP5 Software-Kategorie") | **+** Kapitel 7 (Detail); zusätzlich für Testtiefe-Berechnung genutzt |
| `KATNA` | ✅ (Kategorie "GAMP5 Software-Kategorie") | nur Kapitel 2 - **kein** Kapitel-7-Gegenstück für N/A |
| `EE_P1-4`, `EE_NA` | ✅ (Kategorie "DI EE-Anforderungen") | nur Kapitel 2 |
| `GKATA`, `GKATB`, `GKATC`, `GKATNA` | ✅ (Kategorie "Gerätekategorie") | nur Kapitel 2 - **`GKATB1/B2/B3`, `GKATC1/C2/C3` haben KEINE eigene Checkbox** im Template, nur die Oberkategorie; Subkategorisierung gehört ins Freitextfeld "Besonderheiten" (dafür gibt es dort eigene Textbaustein-Vorschläge) |
| `PR_SOP`, `PR_SOP2`, `PR_Andere`, `PR-Zyklisch` | ✅ (Kategorie "Periodic Review") | nur Kapitel 2; `PR_Andere_Text` (Web-Editor-only) ersetzt den Blanko-Platzhalter hinter `PR_Andere` |
| `VQ`, `NVQ` | ✅ (Kategorie "Vereinfachte Qualifizierung") | nur Kapitel 2 |
| `QUAL`, `VAL` | ✅ (Kategorie "Validierung/Qualifizierung nach SOP", Mehrfachauswahl) | nur Kapitel 2 |
| `KI1-6`, `KINA` | — (nicht mehr direkt im Editor, siehe unten) | **+** Kapitel 9.1 |
| `TTIEFEHOCH/MITTEL/NIEDRIG`, `Z1S1...Z3S3` | — (nicht im Formular) | **werden automatisch BERECHNET** aus GxP-Kritikalität + GAMP5-Kategorie (`fill_testtiefe`) und in Kapitel 2 + der Z-Felder-Matrix in Kapitel 8 eingetragen - nicht direkt eingelesen |

## Kapitel 2 – Informationen und Bemerkungen (Beschreibungstabelle)

| Feld | Im Editor? | Wo im Dokument |
|---|---|---|
| `Prozessbeschreibung` | ✅ (Textbausteine) | Kapitel 2, Beschreibungstabelle - inkl. `Steuerung erfolgt über?` (falls vorhanden) |
| `Daten` | ✅ (Textbausteine) | Beschreibungstabelle |
| `Parameter` | ✅ (Textbausteine) | Beschreibungstabelle |
| `Alarme (GxP-relevant)` | ✅ (Textbausteine) | Beschreibungstabelle |
| `Chargenprotokoll` | ✅ (Textbausteine) | Beschreibungstabelle |
| `Audit Trail (AT)` | ✅ (Textbausteine) | Beschreibungstabelle |
| `Benutzer-verwaltung?` | ✅ (Textbausteine) | Beschreibungstabelle |
| `Schnittstellen mit PLS` | ✅ (Textbausteine) | Beschreibungstabelle |
| `DatenflussAbbildung` *(Web-Editor-only)* | ✅ (Textbaustein) | Beschreibungstabelle, Zeile "Datenfluss / Abbildung:" - Grafiken kann die App nicht einfügen, nur Verweistext |
| `Angeschlossenes Equipment` | ✅ (Textbausteine) | Beschreibungstabelle |
| `Sonstiges` | ✅ (Textbausteine) | Beschreibungstabelle |
| `KI Bewertung` | ✅ (Textbausteine) | Beschreibungstabelle |
| `Besonderheiten` | ✅ (Textbausteine) | Kapitel 2 – Zusammenfassungstabelle (ergänzt den Hinweistext zur GxP-Subkategorisierung) |
| `Steuerung erfolgt über?` | — (nicht mehr separat abgefragt) | wird der Prozessbeschreibung vorangestellt, **falls** im Draft-Dict vorhanden (z. B. aus einem alten Master-Excel-Datensatz) |
| `Bemerkung1` | — (nicht im Editor, siehe unten) | wird beim Start aus einem Datenbank-Eintrag SOFORT in `Prozessbeschreibung` gemergt (nicht erst beim Erzeugen) |
| `Bemerkung2` | — (nicht im Editor, siehe unten) | wird beim Start aus einem Datenbank-Eintrag SOFORT in `Daten` gemergt |
| `Bemerkung3` | — (nicht im Editor, siehe unten) | wird beim Start aus einem Datenbank-Eintrag SOFORT in `Audit Trail (AT)` gemergt |
| `Bemerkung4` | — (nicht im Editor, siehe unten) | wird beim Start aus einem Datenbank-Eintrag SOFORT in `Parameter` gemergt |
| `Hyperlink` | — (nicht im Editor) | **nicht im Dokument** - reiner QualiPSO-Verweis für die Excel |

**Zu den 4 "BemerkungX"-Feldern:** Diese vier Master-Excel-Spalten
sind inhaltlich dasselbe wie `Prozessbeschreibung`/`Daten`/
`Audit Trail (AT)`/`Parameter` - nur unter einem generischen Namen.
Startet man eine neue Systembewertung aus einem Datenbank-Eintrag,
der noch (alte) BemerkungX-Werte enthält, werden diese **sofort beim
Öffnen des Editors** in das jeweilige Hauptfeld gemergt (nicht erst
beim Erzeugen des Dokuments) - dadurch gibt es dort nur noch EIN
sichtbares, bearbeitbares Feld statt zwei getrennter Werte, die sich
sonst am Ende unbemerkt im Dokument verdoppelt hätten (siehe
`_neues_dokument_aus_db_zeile()` in `webapp/app.py`).

## Kapitel 2 (Detail) – GxP-Risikoklassifizierung

| Feld | Im Editor? | Wo im Dokument |
|---|---|---|
| `GxP_Produktqualitaet` | ✅ (Teil der GxP-Begründung) | Kapitel 2 (Detail) - Begründungstext |
| `GxP_Patientensicherheit` | ✅ | Kapitel 2 (Detail) - Begründungstext |
| `GxP_Datenintegritaet` | ✅ | Kapitel 2 (Detail) - Begründungstext |

## Reine Master-Excel-Felder (kein Bezug zum Web-Editor/Template)

Diese Felder tauchen weder im Web-Editor auf, noch werden sie beim
Erzeugen eines neuen Dokuments gebraucht - reine Altlasten/Bookkeeping
in der Master-Excel, ohne Zelle im Template:

`Bedien-SOP`, `SOP-Titel`, `PNK`, `Systemtyp_CE` (berechnete
Übersichtsspalte), `Python ja/nein` (interne Kennzeichnung "per
Skript eingetragen"). `Bearbeiter` siehe oben bei "Deckblatt –
Identifikation".

## Web-Editor-only (nicht Teil der Master-Excel)

Diese 17 Felder gibt es **nur** im Web-Editor - kein eigenes
"zweites DB-Schema", sondern zusätzliche Formularfelder, deren Werte
zwar ins erzeugte Word-Dokument geschrieben werden, aber (wie
gewollt) nie eine eigene Master-Excel-Spalte hatten (Ausnahme:
`KLASS_Global_NA`, siehe unten - das wird NICHT ins Dokument
geschrieben, rein UI):

`Ersteller_Abteilung`, `SI_PL_Abteilung`, `TSO_Abteilung`,
`BSO_Abteilung`, `BQR_Abteilung`, `CSQ_Abteilung`, `PR_Andere_Text`,
`ERES4_SIG_ID_PW`, `ERES4_SIG_BIOMETRISCH`, `ERES4_SIG_TOKEN_PW`,
`DatenflussAbbildung`, `SystemtypZugang_Begruendung`, `KI_Einsatz_Ja`,
`KI_Einsatz_Nein`, `KI_Einsatz_Begruendung`, `CCNr_Rahmen`,
`KLASS_Global_NA`.

**`KI_Einsatz_Ja`/`KI_Einsatz_Nein`/`KI_Einsatz_Begruendung`** (neu):
ersetzen `KI-Reifegrad` (`KI1-6`/`KINA`) als Editor-Frage - siehe
"Nicht automatisch befüllbare Kapitel" unten. Bei "Nein" wird beim
Erzeugen automatisch `KINA` gesetzt und `KI_Einsatz_Begruendung` an
`KI Bewertung` angehängt; bei "Ja" bleibt `KI1-6`/`KINA` in Kapitel 2
bewusst leer.

**`CCNr_Rahmen`** (neu): eigenes, kurzes Feld für Kapitel 1 ("Grund
der Systembewertung") - CC-Nr. + Rahmen der Erstellung (z. B.
"CC-2024-01234 - Periodic Review"), NICHT der lange Freitext aus
`Historie` (der landet ausschließlich in der Dokumentenhistorie-
Tabelle). Direkt im Editor neben "Historie" platziert.

**`KLASS_Global_NA`** (neu): Radiobutton-Option "N/A - kein Globales
CS" in der neuen Kategorie "Globale CS-Klasse (Kapitel 3)" - markiert
lediglich, dass keine der 4 Klassen zutrifft (Lokales CS/Multi-Site-
CS/Equipment ohne CS stattdessen). Anders als die übrigen Web-Editor-
only-Felder wird dieser Wert NICHT ins Dokument geschrieben - er ist
rein informativ für die Editor-Oberfläche, da Kapitel 3 sein "N/A"
bereits automatisch aus `KLASS_Global != "r"` ableitet (siehe
`template_filler.fill_kapitel3`).

## Externe Datenquelle: Fill-a-Masterform-Import

Der Web-Editor kann eine neue Systembewertung auch aus einem
Klartext-Export des anderen Teams ("Fill-a-Masterform") vorbefüllen
(dritte Startquelle, `/masterform`). Dort werden fremde Spaltennamen
(`gxp_kritikalitaet`, `eres_typ`, `subtyp` usw.) auf genau die oben
gelisteten internen Felder/Checkboxen abgebildet - siehe
`lib/masterform_import.py` (Modul-Kommentar + Mapping-Tabellen) und
README.md, Abschnitt "Fill-a-Masterform-Import" für Details und dafür,
welche Import-Felder bewusst NICHT automatisch einer Checkbox
zugeordnet werden.

## Nicht automatisch befüllbare Kapitel (bewusste Lücke)

**Kapitel 5** (Entscheidungsbaum Gerätekategorie/CS-Typ selbst, inkl.
der zugehörigen Ja/Nein-Antworten) wird **nicht** automatisch
befüllt - aus dem in der Zusammenfassungstabelle gespeicherten
Endergebnis lässt sich der zugrunde liegende Entscheidungsweg nicht
eindeutig rekonstruieren (mehrere Antwortpfade können zum selben
Endergebnis führen - ein Raten wäre in einem GxP-Dokument nicht
vertretbar). Muss nach dem automatischen Erzeugen manuell in Word
ergänzt werden.

**Kapitel 9** (KI): aus demselben Grund wird im Editor nur die
einfache Frage "Kommt KI zum Einsatz?" (Ja/Nein, webapp-only
`KI_Einsatz_Ja`/`KI_Einsatz_Nein`) gestellt statt der konkreten
`KI-Reifegrad`-Stufe (I-VI) - die lässt sich ohne die Detailfragen
9.2-9.5 (verbotene Praktiken, Autonomie-/Steuerungsdesignstufe) nicht
zuverlässig ableiten. "Nein" befüllt automatisch 9.1 + `KINA`
(Kapitel 2) und hängt eine angegebene Begründung an `KI Bewertung`
an; "Ja" befüllt nur 9.1 - die genaue Stufe (Kapitel 2 `KI1-6` sowie
Kapitel 9.2-9.5) bleibt wie Kapitel 5 der manuellen Nacharbeit
vorbehalten.

## Wie diese Übersicht erstellt wurde

Nicht aus Erinnerung/Kommentaren zusammengeschrieben, sondern gegen
den tatsächlichen Code geprüft: ein Testlauf hat `lib/template_filler.py`
mit allen bekannten Feldern (143 Master-Excel-Spalten +
17 Web-Editor-only) gefüttert und pro `fill_*`-Funktion protokolliert,
welche Felder tatsächlich gelesen werden. Bei einer künftigen
Template-Änderung (z. B. V12) lohnt es sich, diese Übersicht auf
demselben Weg neu zu erzeugen, statt sie händisch nachzupflegen -
sonst veraltet sie unbemerkt.
