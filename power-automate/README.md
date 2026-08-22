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

## Fertige Import-Datei

Im Ordner liegt `Qualifizierung_Aenderungsmail.zip` – ein **"Package
(Legacy)"** mit exakt dem in dieser Anleitung beschriebenen Flow.

**Import:** https://make.powerautomate.com → **Meine Flows** → **Import**
→ **Import Package (Legacy)** → Datei hochladen → bei den beiden
Verbindungen (SharePoint, Office 365 Outlook) jeweils eine bestehende
Verbindung auswählen oder neu anmelden → **Import**.

> **Hinweis zur Entstehung:** Das Zip-Format ist von Microsoft nicht
> öffentlich dokumentiert. Diese Version basiert nicht mehr auf reinem
> Raten, sondern auf einem **echten, aus demselben Tenant exportierten
> Flow-Paket** (inkl. der tatsächlichen Bibliotheks-ID der SharePoint-Site
> `ICFLantusBB` sowie der realen `SendEmailV2`-Parametersyntax) – die
> Grundstruktur und die Connector-Aufrufe sind dadurch deutlich
> zuverlässiger als der erste Versuch. Trotzdem gilt: zwei Aktionen sind
> reine Best-Effort-Annahmen und nicht 1:1 aus einem echten Export
> bestätigt:
> - **"Dateien abrufen (nur Eigenschaften)"** – der Parameter für die
>   Ordner-Einschränkung (`id`) ist eine Annahme. Als Absicherung filtert
>   der Flow zusätzlich per "Array filtern" auf `{Path}` enthält
>   `/P-2024-05_Engineo/Qualifizierung` – falls die Ordner-Einschränkung
>   in der Connector-Aktion nicht greift und stattdessen die *gesamte*
>   Bibliothek zurückkommt, filtert dieser Schritt trotzdem korrekt auf
>   den richtigen Ordner nach.
> - **`nestingLimit`/`$top`** – ebenfalls Annahmen zu den erweiterten
>   Parametern für Unterordner-Rekursion bzw. Ergebnisobergrenze.
>
> Schlägt der Import trotzdem fehl oder eine dieser beiden Aktionen kommt
> fehlerhaft an: kurz Bescheid geben (am besten mit der genauen
> Fehlermeldung bzw. einem Screenshot der betroffenen Aktion in "Code
> view") – dann lässt sich das gezielt nachbessern. Alternativ bleibt die
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
