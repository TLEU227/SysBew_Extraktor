# Power Automate: Tägliche Änderungs-Mail für den Ordner "Qualifizierung"

Diese Anleitung beschreibt den Aufbau eines **Power Automate Cloud Flows**,
der einmal täglich per E-Mail meldet, welche Dateien im SharePoint-Ordner
`Qualifizierung` **neu erstellt** bzw. **geändert** wurden – inklusive
Angabe, von wem.

> Hinweis: Ein Power-Automate-Flow lässt sich nicht als Datei "installieren"
> wie ein Skript – er wird im Power Automate-Portal (https://make.powerautomate.com)
> zusammengeklickt bzw. importiert. Diese Anleitung ist deshalb bewusst so
> detailliert, dass der Flow 1:1 nachgebaut werden kann. Alle Aktionen,
> Feldnamen und Ausdrücke sind exakt benannt, damit nichts frei interpretiert
> werden muss.

## Status: erfolgreich importiert ✅ (mit einer Testmail bestätigt)

Der Import von `Qualifizierung_Aenderungsmail.zip` wurde am 22.08.2026
erfolgreich im Ziel-Tenant durchgeführt und mit einem echten Testlauf
bestätigt (Datei wurde korrekt als "neu erstellt" erkannt, inkl.
Ersteller-Name).

- **v5:** Die von "Create HTML table" erzeugten Tabellen hatten keine
  Rahmen/Zellenabstände (Text wirkte "zusammengequetscht"). Die beiden
  Tabellen bekommen jetzt per `replace(...)` nachträglich Inline-CSS
  (Rahmen, Padding, `word-wrap`) verpasst, bevor sie in die Mail
  eingebaut werden – bewusst **Inline-Styles statt `<style>`-Block**, da
  Outlook (insbesondere die Desktop-App mit Word-Rendering) `<style>`-
  Blöcke in HTML-Mails häufig ignoriert.
- **v6:** Zwei neue Spalten in beiden Tabellen:
  - **"Link"** – Text-Link `"Datei öffnen"`, der direkt auf `{Link}` (die
    von SharePoint gelieferte Öffnen-URL der Datei) verweist.
  - **"QualiPSO"** – zeigt den **aktuellen** Wert des SharePoint-Metadatenfelds
    `QualiPSO` (bewusst kein Alt/Neu-Vergleich – das würde eine
    zusätzliche Versionsverlauf-Abfrage pro Datei erfordern; siehe
    Rücksprache mit dem Nutzer). **Annahme:** interner Feldname ist exakt
    `QualiPSO` (= Anzeigename, da keine Sonderzeichen/Leerzeichen enthalten).
    Falls das Feld in der Mail leer bleibt: internen Feldnamen im
    Bibliotheks-Spaltenmanagement prüfen (kann von der Anzeigebezeichnung
    abweichen) und in beiden "HTML-Tabelle"-Aktionen anpassen.
- **v7:** `QualiPSO` ist **kein Ja/Nein-Feld**, sondern ein Auswahlfeld
  mit mehreren Stufen (z. B. "Review 1", "Review 2", "Freigegeben"). Die
  Formel zeigt daher jetzt einfach den Rohwert direkt an
  (`@item()?['QualiPSO']`) statt ihn auf Ja/Nein zu erzwingen.
- **v8 (aktuell):**
  - **Link-Spalte reparlert:** "Create HTML table" kodiert Zellinhalte
    HTML-sicher (`<` → `&lt;` usw.), wodurch der eingebettete
    `<a href="...">`-Tag nicht als Link gerendert wurde, sondern als
    kompletter Text sichtbar war. Die Mailtext-Zusammensetzung dekodiert
    jetzt `&lt;` `&gt;` `&quot;` `&amp;` in beiden Tabellen wieder zurück,
    nachdem die Styling-Ersetzung gelaufen ist – dadurch wird nur der
    absichtlich eingefügte `<a>`-Tag zu einem echten, kurzen Text-Link
    ("Datei öffnen").
  - **Umlaute korrigiert:** Description und Mailtext verwendeten teils
    ASCII-Ersatzschreibweisen (`ae`/`oe`/`ue` statt `ä`/`ö`/`ü`, z. B.
    "Aenderungen", "taeglich") – jetzt durchgängig echte Umlaute.
- **v9:** Anhand der echten Rohausgabe identifiziert – das Feld heißt
  intern **nicht** `QualiPSO`, sondern
  `OData__x00dc_betragennachQualiPSO` (SharePoint kodiert das "Ü" aus
  "Übertragen nach QualiPSO" als `_x00dc_`), und ist zusätzlich ein
  **Lookup-Feld** (Objekt mit `Id`/`Value`, kein einfacher Text). Formel
  korrigiert auf `@item()?['OData__x00dc_betragennachQualiPSO']?['Value']`.
- **v10 (aktuell):** Zweiter Empfänger ergänzt –
  `emailMessage/To` ist jetzt `thomas.leuckel@sanofi.com;karlheinz.preuss@engineo.com`
  (mehrere Empfänger per Semikolon getrennt).

## Fertige Import-Datei

Im Ordner liegt `Qualifizierung_Aenderungsmail.zip` – ein **"Package
(Legacy)"** mit exakt dem in dieser Anleitung beschriebenen Flow.

**Import:** https://make.powerautomate.com → **Meine Flows** → **Import**
→ **Import Package (Legacy)** → Datei hochladen → bei den beiden
Verbindungen (SharePoint, Office 365 Outlook) jeweils eine bestehende
Verbindung auswählen oder neu anmelden → **Import**.

> **Versionsverlauf (Stand: erfolgreich importiert):**
> - **v1 (verworfen):** komplett von Hand geraten → Import schlug fehl
>   ("Something went wrong").
> - **v2 (verworfen):** anhand eines echten Tenant-Exports (anderer, real
>   funktionierender Flow) nachgebaut, inkl. `authentication`-Parameter
>   und echter Bibliotheks-ID → Import lief so weit, dass ein Entwurf
>   angelegt wurde, aber die SharePoint-/Outlook-**Connections** wurden
>   nicht korrekt zugeordnet.
> - **v3 (aktuell):** Ursache gefunden, indem der tatsächlich importierte
>   und danach wieder exportierte Flow analysiert wurde. Der Fix: der
>   logische Verbindungsname braucht ein **`_1`-Suffix**
>   (`shared_sharepointonline_1` / `shared_office365_1` statt nur
>   `shared_sharepointonline` / `shared_office365`) – das war der Grund,
>   warum die Connection-Zuordnung nicht griff. Zusätzlich wurden die
>   Parameter `id`/`nestingLimit` bei "Dateien abrufen" entfernt (waren
>   ungültig und wurden beim Import ohnehin verworfen); die Ordner-Eingrenzung
>   läuft stattdessen komplett über den nachgeschalteten "Array
>   filtern"-Schritt auf `{Path}`. Die Connection- und API-Ressourcen-IDs
>   in dieser Datei sind außerdem die **echten, im Tenant bestätigten
>   GUIDs** (aus zwei realen Exports übernommen), keine Zufallswerte mehr.
>
> Die Aktion "Dateien abrufen (nur Eigenschaften)" liest dadurch aktuell
> die **gesamte Bibliothek** "Shared Documents" auf der Site
> `ICFLantusBB` (nicht nur den Unterordner) und filtert erst im zweiten
> Schritt auf `/P-2024-05_Engineo/Qualifizierung`. Funktioniert, ist bei
> sehr großen Bibliotheken aber weniger effizient als eine direkte
> Ordner-Einschränkung (SharePoint-Connector unterstützt das für diese
> Aktion offenbar nicht über die hier verfügbaren Parameter).
>
> Der Flow enthält außerdem eine **Beschreibung** (Flow-Details in Power
> Automate → Reiter "Details"/"About") mit Kurzfassung von Zweck und
> Ablauf. Hinweis: Power Automate begrenzt dieses Feld auf **1024
> Zeichen** (`WorkflowDescriptionTooLong`, falls überschritten) – die
> Beschreibung ist entsprechend kurz gehalten, Details stehen weiterhin
> hier im README.
>
> - **v4 (aktuell):** v3 scheiterte beim Speichern mit
>   `WorkflowDescriptionTooLong` (1260 statt max. 1024 Zeichen) – die
>   Beschreibung wurde auf ca. 590 Zeichen gekürzt.
>
> Sollte der Import trotzdem noch nicht ganz durchlaufen: bitte die genaue
> Fehlermeldung bzw. einen Screenshot der betroffenen Aktion in "Code view"
> schicken – dann lässt sich gezielt nachbessern. Alternativ bleibt die
> Schritt-für-Schritt-Anleitung unten als garantiert funktionierender Weg.

## Betroffene SharePoint-Struktur

- **Site:** `https://sanofi.sharepoint.com/sites/ICFLantusBB`
- **Bibliothek:** `Freigegebene Dokumente` / `Shared Documents`
- **Überwachter Ordner:**
  `/sites/ICFLantusBB/Shared Documents/P-2024-05_Engineo/Qualifizierung`
- **Empfänger der Mail:** `thomas.leuckel@sanofi.com`
- **Turnus:** täglich (Uhrzeit frei wählbar, siehe Schritt 1)

> Den exakten Server-relativen Pfad vor dem Aufbau bitte einmal in
> SharePoint prüfen: Ordner "Qualifizierung" öffnen → *Details* (Info-Icon
> oben rechts) → dort steht der vollständige Pfad. Falls er von obigem Pfad
> abweicht, im Flow entsprechend anpassen.

## Voraussetzungen

- Power Automate-Lizenz, die den **SharePoint**- und den
  **Office 365 Outlook**-Connector abdeckt (i. d. R. in M365 enthalten,
  Standard-Connectors, keine Premium-Lizenz nötig).
- Das Konto, unter dem der Flow läuft, braucht mindestens **Leserechte**
  auf die Site `ICFLantusBB`.
- Empfehlung: Flow nicht unter dem persönlichen Konto, sondern (falls
  vorhanden) unter einem Team-/Funktionskonto anlegen, damit er nicht
  ausfällt, wenn eine Person das Unternehmen verlässt.

## Funktionsprinzip

Der Flow läuft 1×/Tag, holt alle Dateien aus dem Ordner (inkl.
Unterordner), filtert die, die seit dem letzten Lauf (≈ letzte 24h)
**erstellt** oder **geändert** wurden, trennt sie in zwei Listen ("Neu" /
"Geändert") und verschickt sie als zwei HTML-Tabellen per Mail.

---

## Schritt-für-Schritt-Aufbau

### 1. Flow anlegen

1. https://make.powerautomate.com öffnen → **Erstellen** → **Automatisierter
   Cloud Flow ohne Trigger auslösen** (bzw. *"Scheduled cloud flow"*).
2. Name: `Qualifizierung – tägliche Änderungsmail`
3. Trigger: **Wiederholung** (`Recurrence`)
   - Intervall: `1`
   - Häufigkeit: `Tag`
   - Erweiterte Optionen → **Uhrzeit** z. B. `06:00`, **Zeitzone**
     `W. Europe Standard Time` (damit die Uhrzeit unabhängig vom
     Rechenzentrum stimmt).

### 2. Variable für den Vergleichszeitpunkt

Aktion **Variable initialisieren** (`Initialize variable`):

| Feld | Wert |
|---|---|
| Name | `varSeit` |
| Typ | Zeichenfolge (String) |
| Wert | `@{addDays(utcNow(), -1)}` |

Das ist der Zeitstempel "vor 24 Stunden" (UTC), gegen den später verglichen
wird. Läuft der Flow einmal täglich ohne Aussetzer, deckt das lückenlos
den Zeitraum seit dem letzten Lauf ab.

### 3. Alle Dateien im Ordner holen

Aktion **SharePoint → Dateien abrufen (nur Eigenschaften)**
(`Get files (properties only)`):

| Feld | Wert |
|---|---|
| Websiteadresse | `https://sanofi.sharepoint.com/sites/ICFLantusBB` |
| Bibliotheksname | `Dokumente` (bzw. `Shared Documents`) |
| Ordner-Id / Limit Entries to Folder (erweiterte Optionen) | `/P-2024-05_Engineo/Qualifizierung` |
| Include Nested Items (erweiterte Optionen) | **Ja** – damit auch Dateien in Unterordnern von "Qualifizierung" erfasst werden |
| Top Count (erweiterte Optionen) | z. B. `5000` (SharePoint-Standardlimit; bei mehr Dateien Pagination in den Flow-Einstellungen einschalten, siehe [Hinweise](#hinweise--stolperfallen)) |

### 4. Nur geänderte/neue Dateien der letzten 24h filtern

Aktion **Array filtern** (`Filter array`):

- **Von:** Ausgabe von Schritt 3 (`value` der "Dateien abrufen"-Aktion)
- Umschalten auf **Erweiterter Modus** und folgenden Ausdruck einfügen:

```
@or(
  greaterOrEquals(item()?['Modified'], variables('varSeit')),
  greaterOrEquals(item()?['Created'], variables('varSeit'))
)
```

Ergebnis: alle Dateien, die seit `varSeit` entweder angelegt oder
bearbeitet wurden.

### 5. Trennen in "Neu erstellt" und "Geändert"

Zwei weitere **Filter array**-Aktionen, jeweils **von der Ausgabe aus
Schritt 4**:

**5a. Neu erstellt:**
```
@greaterOrEquals(item()?['Created'], variables('varSeit'))
```

**5b. Geändert (aber nicht neu):**
```
@and(
  greaterOrEquals(item()?['Modified'], variables('varSeit')),
  less(item()?['Created'], variables('varSeit'))
)
```

### 6. Abbrechen, wenn nichts passiert ist

Aktion **Bedingung** (`Condition`):

```
@or(
  greater(length(body('Neu_erstellt')), 0),
  greater(length(body('Geändert_aber_nicht_neu')), 0)
)
```

Alle folgenden Schritte (7–9) kommen in den **Ja**-Zweig, damit an Tagen
ohne Änderungen keine (leere) Mail verschickt wird.

### 7. HTML-Tabellen bauen

Zwei Aktionen **HTML-Tabelle erstellen** (`Create HTML table`),
Tabellenformat je auf **Benutzerdefiniert** stellen:

**7a. "Neue Dateien"** – von: Ausgabe aus Schritt 5a

| Spalte | Wert (Ausdruck) |
|---|---|
| Dateiname | `item()?['{FilenameWithExtension}']` |
| Ordner | `item()?['{Path}']` |
| Erstellt von | `item()?['Author']?['DisplayName']` |
| Erstellt am | `formatDateTime(item()?['Created'], 'dd.MM.yyyy HH:mm')` |
| Link | `item()?['{Link}']` |

**7b. "Geänderte Dateien"** – von: Ausgabe aus Schritt 5b

| Spalte | Wert (Ausdruck) |
|---|---|
| Dateiname | `item()?['{FilenameWithExtension}']` |
| Ordner | `item()?['{Path}']` |
| Geändert von | `item()?['Editor']?['DisplayName']` |
| Geändert am | `formatDateTime(item()?['Modified'], 'dd.MM.yyyy HH:mm')` |
| Link | `item()?['{Link}']` |

> Die exakten dynamischen Feldnamen (`{FilenameWithExtension}`, `{Path}`,
> `{Link}`, `Author`, `Editor`) am besten über den **Dynamischer
> Inhalt**-Picker aus der Ausgabe von Schritt 3 auswählen statt
> abzutippen – Power Automate benennt intern gekapselte Felder mit
> geschweiften Klammern. `Author`/`Editor` liefern ein Personenobjekt mit
> u. a. `DisplayName` und `Email`.

### 8. Mailtext zusammensetzen

Aktion **Kompilieren** (`Compose`), Eingabe:

```html
<h3>Ordner "Qualifizierung" – Änderungen vom @{formatDateTime(utcNow(),'dd.MM.yyyy')}</h3>

<p><b>Neu erstellte Dateien (@{length(body('Neu_erstellt'))})</b></p>
@{if(greater(length(body('Neu_erstellt')), 0), body('HTML_Tabelle_Neue_Dateien'), '<p><i>Keine neuen Dateien.</i></p>')}

<p><b>Geänderte Dateien (@{length(body('Geändert_aber_nicht_neu'))})</b></p>
@{if(greater(length(body('Geändert_aber_nicht_neu')), 0), body('HTML_Tabelle_Geänderte_Dateien'), '<p><i>Keine geänderten Dateien.</i></p>')}

<p style="color:grey;font-size:11px">Automatische Nachricht von Power Automate – Flow "Qualifizierung – tägliche Änderungsmail".</p>
```

### 9. Mail versenden

Aktion **Office 365 Outlook → E-Mail senden (V2)**
(`Send an email (V2)`):

| Feld | Wert |
|---|---|
| An | `thomas.leuckel@sanofi.com` |
| Betreff | `Qualifizierung: Dateiänderungen vom @{formatDateTime(utcNow(),'dd.MM.yyyy')}` |
| Textkörper | Ausgabe von Schritt 8 |
| Ist HTML | `Ja` (in den erweiterten Optionen aktivieren) |

### 10. Speichern und testen

1. Flow speichern.
2. **Testen → Manuell** ausführen.
3. Kontrollieren, ob im Ordner "Qualifizierung" innerhalb der letzten 24h
   Dateien angelegt/geändert wurden – nur dann kommt (bewusst) eine Mail.
   Zum Testen notfalls kurz eine Testdatei anlegen/umbenennen.

---

## Hinweise / Stolperfallen

- **Mehr als 5000 Dateien:** SharePoint-Listenansichten sind bei mehr als
  5000 Elementen limitiert. Falls der Ordner "Qualifizierung" (inkl.
  Unterordner) das überschreitet: in den Flow-Einstellungen der
  "Dateien abrufen"-Aktion **Pagination = Ein** setzen und einen
  höheren *Threshold* (z. B. 100000) angeben.
- **Zeitzone:** `utcNow()` liefert UTC. Die Anzeige-Formatierung in
  Schritt 7/8 (`formatDateTime`) kann bei Bedarf mit
  `convertFromUTC(..., 'W. Europe Standard Time')` auf MEZ/MESZ
  umgerechnet werden, falls die Uhrzeiten in der Mail lokal angezeigt
  werden sollen.
- **Gelöschte Dateien:** Dieser Flow meldet nur neue/geänderte Dateien,
  keine Löschungen. Falls gewünscht, zusätzlich mit dem SharePoint-Trigger
  *"Wenn eine Datei gelöscht wird"* separat abbilden (eigener, einfacher
  Zusatz-Flow).
- **"Author"/"Editor" leer bei Sync-Client-Uploads:** Bei Dateien, die per
  OneDrive-Sync-Client oder Explorer-Ansicht hochgeladen wurden, sind
  Autor/Bearbeiter i. d. R. trotzdem gesetzt; bei Massenkopien über
  PowerShell/Migrationstools kann es abweichen (dann steht dort ggf. ein
  Dienstkonto).
- **Ausfallsicherheit:** Fällt ein Tageslauf aus (z. B. Wartungsfenster),
  schließt `varSeit = addDays(utcNow(), -1)` eine Lücke von mehr als 24h
  **nicht** automatisch. Um das robuster zu machen, kann `varSeit`
  stattdessen aus dem Zeitstempel des **letzten erfolgreichen Laufs**
  gelesen werden (z. B. über eine kleine Konfigurationsliste/-datei, in
  die der Flow nach jedem Lauf `utcNow()` schreibt). Für den
  Normalbetrieb ist die feste 24h-Differenz aber ausreichend.
- **Weitere Empfänger/Ordner:** Zusätzliche Empfänger einfach im Feld "An"
  per Semikolon ergänzen; ein zweiter zu überwachender Ordner am
  einfachsten durch Duplizieren des Flows und Anpassen von Schritt 3.

## Anpassungsmöglichkeiten

- **Nur bestimmte Dateitypen** (z. B. nur `.docx`/`.xlsx`): in Schritt 4/5
  zusätzlich `endsWith(item()?['{FilenameWithExtension}'], '.docx')` per
  `and(...)` ergänzen.
- **Kürzerer/längerer Turnus:** Wiederholung in Schritt 1 anpassen (z. B.
  stündlich) – `varSeit` in Schritt 2 dann entsprechend auf
  `addHours(utcNow(), -1)` ändern.
- **Teams-Nachricht statt/zusätzlich zur Mail:** Schritt 9 um eine
  Teams-Aktion ("Nachricht in einem Chat oder Kanal posten") ergänzen,
  gleicher Inhalt aus Schritt 8 wiederverwendbar.
