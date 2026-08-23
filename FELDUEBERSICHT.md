# Feldübersicht: Master-Excel ↔ Web-Editor ↔ Word-Template

Nachschlagewerk für alle Felder, die im System vorkommen: welche gibt
es, woher kommen sie, und welche davon landen tatsächlich im
automatisch erzeugten Word-Dokument (und wo genau)?

**Kurze Antwort auf "welche DB(s)":** Es gibt genau **eine** echte
Datenbank - die Master-Excel (`Systembewertungen_GESAMT.xlsx`,
Sheet `SysBew`). Ihre komplette Spaltenliste steht in
`lib/sysbew_common.py` als `EXCEL_COLUMNS` (143 Spalten). Zusätzlich
kennt **nur der Web-Editor** (`webapp/app.py`) 12 weitere Felder, die
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
| `Historie` (Grund der Erstellung/Änderung) | ✅ (eigenes Feld + Textbaustein-Vorschläge) | ✅ | Dokumentenhistorie-Tabelle **+** Kapitel 1 ("Grund der Systembewertung" neben Neuerstellung/Änderung) |
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
| `KLASS_Global` | ✅ | ✅ | Kapitel 1 – Klassifizierung **+** steuert die "N/A - Weiter mit Kap. 4"-Checkbox in Kapitel 3 (angekreuzt, wenn NICHT "r" - Kapitel 3 fragt nur die Detailfrage innerhalb von "Globales CS" ab) |
| `KLASS_Global_1a` | ✅ | ✅ | Kapitel 1 – Klassifizierung **+** Kapitel 3 (Detailfrage) |
| `KLASS_Global_1b` | ✅ | ✅ | Kapitel 1 – Klassifizierung **+** Kapitel 3 |
| `KLASS_Global_2` | ✅ | ✅ | Kapitel 1 – Klassifizierung **+** Kapitel 3 |
| `KLASS_Global_3` | ✅ | ✅ | Kapitel 1 – Klassifizierung **+** Kapitel 3 |
| `KLASS_OhneCS` | ✅ | ✅ | Kapitel 1 – Klassifizierung |

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
| `KI1-6`, `KINA` | ✅ (Kategorie "KI-Reifegrad") | **+** Kapitel 9.1 (nur die Ja/Nein-Frage "kommt KI zum Einsatz" wird daraus abgeleitet, 9.2-9.5 NICHT) |
| `TTIEFEHOCH/MITTEL/NIEDRIG`, `Z1S1...Z3S3` | — (nicht im Formular) | **werden automatisch BERECHNET** aus GxP-Kritikalität + GAMP5-Kategorie (`fill_testtiefe`) und in Kapitel 2 + der Z-Felder-Matrix in Kapitel 8 eingetragen - nicht direkt eingelesen |

## Kapitel 2 – Informationen und Bemerkungen (Beschreibungstabelle)

| Feld | Im Editor? | Wo im Dokument |
|---|---|---|
| `Prozessbeschreibung` | ✅ (Textbausteine) | Kapitel 2, Beschreibungstabelle - inkl. `Steuerung erfolgt über?` (falls vorhanden) und `Bemerkung1` vorangestellt/angehängt |
| `Daten` | ✅ (Textbausteine) | Beschreibungstabelle - inkl. `Bemerkung2` |
| `Parameter` | ✅ (Textbausteine) | Beschreibungstabelle - inkl. `Bemerkung4` |
| `Alarme (GxP-relevant)` | ✅ (Textbausteine) | Beschreibungstabelle |
| `Chargenprotokoll` | ✅ (Textbausteine) | Beschreibungstabelle |
| `Audit Trail (AT)` | ✅ (Textbausteine) | Beschreibungstabelle - inkl. `Bemerkung3` |
| `Benutzer-verwaltung?` | ✅ (Textbausteine) | Beschreibungstabelle |
| `Schnittstellen mit PLS` | ✅ (Textbausteine) | Beschreibungstabelle |
| `DatenflussAbbildung` *(Web-Editor-only)* | ✅ (Textbaustein) | Beschreibungstabelle, Zeile "Datenfluss / Abbildung:" - Grafiken kann die App nicht einfügen, nur Verweistext |
| `Angeschlossenes Equipment` | ✅ (Textbausteine) | Beschreibungstabelle |
| `Sonstiges` | ✅ (Textbausteine) | Beschreibungstabelle |
| `KI Bewertung` | ✅ (Textbausteine) | Beschreibungstabelle |
| `Besonderheiten` | ✅ (Textbausteine) | Kapitel 2 – Zusammenfassungstabelle (ergänzt den Hinweistext zur GxP-Subkategorisierung) |
| `Steuerung erfolgt über?` | — (nicht mehr separat abgefragt) | wird der Prozessbeschreibung vorangestellt, **falls** im Draft-Dict vorhanden (z. B. aus einem alten Master-Excel-Datensatz) |
| `Bemerkung1` | — (nicht im Editor) | an `Prozessbeschreibung` angehängt |
| `Bemerkung2` | — (nicht im Editor) | an `Daten` angehängt |
| `Bemerkung3` | — (nicht im Editor) | an `Audit Trail (AT)` angehängt |
| `Bemerkung4` | — (nicht im Editor) | an `Parameter` angehängt |
| `Hyperlink` | — (nicht im Editor) | **nicht im Dokument** - reiner QualiPSO-Verweis für die Excel |

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

Diese 12 Felder gibt es **nur** im Web-Editor - kein eigenes
"zweites DB-Schema", sondern zusätzliche Formularfelder, deren Werte
zwar ins erzeugte Word-Dokument geschrieben werden, aber (wie
gewollt) nie eine eigene Master-Excel-Spalte hatten:

`Ersteller_Abteilung`, `SI_PL_Abteilung`, `TSO_Abteilung`,
`BSO_Abteilung`, `BQR_Abteilung`, `CSQ_Abteilung`, `PR_Andere_Text`,
`ERES4_SIG_ID_PW`, `ERES4_SIG_BIOMETRISCH`, `ERES4_SIG_TOKEN_PW`,
`DatenflussAbbildung`, `SystemtypZugang_Begruendung`.

## Nicht automatisch befüllbare Kapitel (bewusste Lücke)

**Kapitel 5** (Entscheidungsbaum Gerätekategorie/CS-Typ selbst, inkl.
der zugehörigen Ja/Nein-Antworten) wird **nicht** automatisch
befüllt - aus dem in der Zusammenfassungstabelle gespeicherten
Endergebnis lässt sich der zugrunde liegende Entscheidungsweg nicht
eindeutig rekonstruieren (mehrere Antwortpfade können zum selben
Endergebnis führen - ein Raten wäre in einem GxP-Dokument nicht
vertretbar). Muss nach dem automatischen Erzeugen manuell in Word
ergänzt werden. Ebenso die Detailfragen 9.2-9.5 (Kapitel 9, verbotene
Praktiken/Autonomie-/Steuerungsdesignstufe) - nur 9.1 wird aus dem
KI-Reifegrad abgeleitet.

## Wie diese Übersicht erstellt wurde

Nicht aus Erinnerung/Kommentaren zusammengeschrieben, sondern gegen
den tatsächlichen Code geprüft: ein Testlauf hat `lib/template_filler.py`
mit allen bekannten Feldern (143 Master-Excel-Spalten +
12 Web-Editor-only) gefüttert und pro `fill_*`-Funktion protokolliert,
welche Felder tatsächlich gelesen werden. Bei einer künftigen
Template-Änderung (z. B. V12) lohnt es sich, diese Übersicht auf
demselben Weg neu zu erzeugen, statt sie händisch nachzupflegen -
sonst veraltet sie unbemerkt.
