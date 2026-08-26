# ============================================================
# app.py
# Systembewertung-Editor - Web-Oberflaeche - 1.19
#
# Erzeugt NEUE Systembewertungen (V11) aus Daten der Master-Excel
# ("Datenbank"), aus einem hochgeladenen Fill-a-Masterform-Export
# (siehe masterform_import.py, Routen /masterform/...) oder von
# Grund auf, mit Zwischenspeicherung als Draft und Mehrbenutzer-
# faehiger Sperrung ueber draft_store.py.
#
# WICHTIG: laeuft LOKAL bei jeder/jedem Nutzer:in (kein zentraler
# Server) - siehe README.md, Abschnitt "Web-Editor". Drafts landen im
# gemeinsamen Ordner "Drafts" neben der Master-Excel, sind also fuer
# alle sichtbar/oeffenbar, unabhaengig davon, auf welchem PC sie
# angelegt wurden.
#
# WICHTIG (1.1): dieser Web-Editor schreibt NICHT (mehr) in die Master-
# Excel. Er erzeugt ausschliesslich das neue .docx zum Download. Die
# Master-Excel wird weiterhin ausschliesslich ueber den bestehenden,
# geprueften Import-Weg aktualisiert: das FINALE, tatsaechlich
# abgeschlossene/unterschriebene Dokument wird auf word_parser_main.py
# gezogen. Damit gibt es fuer die Excel-Befuellung nur EINEN Codepfad,
# und es landen keine Entwuerfe/Zwischenstaende in der Excel.
#
# Start: python app.py  (oeffnet automatisch den Browser)
# Benoetigt zusaetzlich zu den Anforderungen von word_parser_main.py:
#   pip install flask
# ============================================================

import json
import os
import re
import sys
import tempfile
import time
import webbrowser
from datetime import date
from threading import Thread, Timer

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "lib"))

from flask import Flask, flash, redirect, render_template, request, send_file, session, url_for

import db_reader
import draft_store
import masterform_import
import sysbew_common as common
import template_filler
from word_parser_v11 import VALIDATION_KATEGORIEN_V11

app = Flask(__name__)
# Nur lokal genutzt (127.0.0.1, kein oeffentlicher Zugriff) - deshalb
# reicht ein fester Schluessel, es stehen keine sicherheitskritischen
# Daten in der Session (nur der angezeigte Benutzername).
app.secret_key = "sysbew-editor-lokal"

# Wird unten in der Fusszeile jeder Seite angezeigt (siehe base.html) -
# damit sich nach einem "git pull" auf einen Blick pruefen laesst, ob
# der gerade laufende Prozess auch tatsaechlich neu gestartet wurde
# (Flask laedt Code-Aenderungen NICHT automatisch nach, debug=False).
APP_VERSION = "1.21"

@app.context_processor
def _globale_template_variablen():
    return {"app_version": APP_VERSION, "optionen_hinweise": OPTIONEN_HINWEISE}

PORT = 5151

# ============================================================
# Auto-Beenden: sobald niemand mehr die Webseite offen hat, beendet
# sich dieser Prozess selbst - das Terminal-Fenster schliesst sich
# dann automatisch (siehe Systembewertung_Editor_starten.bat, die bei
# sauberem Exit-Code das "Druecke eine Taste"-Pause ueberspringt).
#
# Funktionsweise: base.html schickt per JS alle paar Sekunden ein
# "Lebenszeichen" an /lebenszeichen, solange irgendeine Seite dieser
# App in einem Browser-Tab offen ist. ueberschreitet die Zeit seit dem
# letzten Lebenszeichen ein Zeitlimit, geht dieser Hintergrund-Thread
# davon aus, dass kein Tab mehr offen ist, und beendet den Prozess
# (os._exit - sofort, ohne Cleanup-Verzoegerung; unproblematisch, da
# nichts Ungespeichertes hier haengt: Drafts werden explizit per
# Speichern-Button persistiert, nicht automatisch beim Beenden).
#
# WICHTIG: reagiert NICHT auf einzelne Seitenwechsel innerhalb der App
# (z.B. von der Datenbank-Suche zum Editor) - die neu geladene Seite
# schickt sofort wieder ein Lebenszeichen, lange bevor das normale
# Zeitlimit ablaufen wuerde. Beim tatsaechlichen Schliessen eines Tabs
# (oder Browser-Fensters) schickt base.html zusaetzlich per
# navigator.sendBeacon() ein explizites "wird jetzt geschlossen"-
# Signal an /schliessen-signal, das die Wartezeit stark verkuerzt -
# dadurch reagiert das Beenden bei einem echten Tab-Schluss schnell
# (SCHLIESSEN_SIGNAL_TIMEOUT_SEK), waehrend ein laengeres, normales
# Zeitlimit (LEBENSZEICHEN_TIMEOUT_SEK) als Sicherheitsnetz dient -
# etwa wenn der PC in den Ruhezustand geht, die Verbindung abreisst,
# oder der Browser abstuerzt, ohne dass ein Schliessen-Signal ankam.
# Ein Tab, der nur in den Hintergrund/minimiert ist, schickt weiterhin
# regelmaessig Lebenszeichen (Browser drosseln Hintergrund-Timer zwar,
# aber nicht so stark, dass 20s ueberschritten wuerden) und beendet
# den Server dadurch NICHT versehentlich.
LEBENSZEICHEN_TIMEOUT_SEK = 20
SCHLIESSEN_SIGNAL_TIMEOUT_SEK = 3

_letztes_lebenszeichen = time.time()
_naechste_faelligkeit = None  # per /schliessen-signal gesetzte verkuerzte Frist

def _auto_beenden_watchdog():
    while True:
        time.sleep(1)
        frist = _naechste_faelligkeit or (_letztes_lebenszeichen + LEBENSZEICHEN_TIMEOUT_SEK)
        if time.time() > frist:
            print("\n  Keine offene Browser-Seite mehr erkannt - beende automatisch.")
            os._exit(0)

Thread(target=_auto_beenden_watchdog, daemon=True).start()

@app.route("/lebenszeichen", methods=["POST"])
def lebenszeichen():
    global _letztes_lebenszeichen, _naechste_faelligkeit
    _letztes_lebenszeichen = time.time()
    _naechste_faelligkeit = None
    return {"ok": True}

@app.route("/schliessen-signal", methods=["POST"])
def schliessen_signal():
    global _naechste_faelligkeit
    _naechste_faelligkeit = time.time() + SCHLIESSEN_SIGNAL_TIMEOUT_SEK
    return {"ok": True}

# Felder, die nicht als Eingabe im Formular auftauchen:
#   - automatisch berechnet: Systemtyp_CE, Erkannte_Version (siehe
#     info.html-Baustein direkt hinter "Dok. -Nr." - immer V11, solange
#     das nur eine hinterlegte Version ist), Testtiefe/Testtiefe-Matrix
#     (wird seit 1.1 automatisch aus GxP-Kritikalitaet + Software-
#     Kategorie berechnet, siehe template_filler.fill_testtiefe)
#   - Phenix: Phenix-Nummern gibt es laut Fachbereich nicht mehr
#   - DI EE-Anforderungen: das Ergebnis laesst sich nur ueber den
#     Entscheidungsbaum in Kapitel 5 korrekt ermitteln, den diese App
#     nicht abbildet - ein direktes Ankreuzen ohne den Baum waere
#     Raten und in einem GxP-Dokument nicht vertretbar
#   - "Steuerung erfolgt über?": hat keine eigene Frage im Template,
#     gehoert inhaltlich zur Prozessbeschreibung und wird dort direkt
#     mit erfasst statt als eigenes Feld abgefragt
#   - API/BE/Raum/SAP/Bearbeiter/PLSTA: reine Master-Excel-Spalten ohne
#     jede Rolle beim Erzeugen des Dokuments (per Code-Analyse
#     geprueft: template_filler.py liest sie nie; API ist zudem nur
#     eine aus "Betrieb" abgeleitete Teilzeichenkette fuers Filtern in
#     der Excel, "Betrieb" selbst wird geschrieben).
#     "DokNummerVorQualiPSO" ist NICHT mehr in dieser Liste - wird seit
#     1.7 im Formular als "Vorherige Doc-ID" angezeigt und beim Start
#     aus einem Datenbank-Eintrag automatisch mit der alten Dok.-Nr./
#     Version befuellt (siehe _neues_dokument_aus_db_zeile).
#   - "Version": ist beim LESEN eines bestehenden Dokuments die dort
#     bereits vorhandene, hoechste Version - fuer ein NEUES Dokument
#     irrelevant (das startet immer bei "Version_Historie" = 1.0).
#   - "SW-Version / Typ:"/"SW-Name:": beim Lesen per Regex aus dem
#     Besonderheiten-Text abgeleitet, haben aber keine eigene Zelle im
#     Template - bei Bedarf einfach direkt im Feld "Besonderheiten"
#     mit angeben.
#   - "BemerkungX": inhaltlich redundant, da Prozessbeschreibung/Daten/
#     Audit Trail/Parameter bereits weiter oben unter ihrem eigenen
#     Namen abgefragt werden - deshalb hier keine doppelte Abfrage.
#     Bei Start aus einem Datenbank-Eintrag wird ein evtl. vorhandener
#     BemerkungX-Wert in _neues_dokument_aus_db_zeile() bereits VOR dem
#     Anzeigen des Formulars in das jeweilige Hauptfeld gemergt (nicht
#     erst beim Erzeugen des Dokuments) - dadurch bleibt es sichtbar/
#     bearbeitbar statt unbemerkt im Hintergrund an das Hauptfeld
#     angehaengt zu werden. Der Transfer der Werte zurueck in die
#     "BemerkungX"-Spalten der Master-Excel passiert weiterhin
#     automatisch beim Einlesen des fertigen, unterschriebenen
#     Dokuments (word_parser_v8/10/11.py machen das schon seit jeher
#     so) - unabhaengig davon.
#   - "Version_Historie": wird nicht mehr getrennt abgefragt, sondern
#     zusammen mit "Dok. -Nr." in einem gemeinsamen Feld erfasst (siehe
#     DOK_NR_VERSION_FELD/_dok_nr_version_kombinieren/_auftrennen unten)
#     - beides landet inhaltlich zusammen ("Dok.-ID"), daher ein Feld
#     statt zwei getrennter.
#   - "KI-Reifegrad" (I-VI/N/A): genau wie "DI EE-Anforderungen" laesst
#     sich die konkrete Stufe nicht ohne die Detailfragen 9.2-9.5
#     (Autonomie-/Steuerungsdesign-Stufe) zuverlaessig abfragen. Ersetzt
#     durch die einfache Frage "Kommt KI zum Einsatz?" (Ja/Nein, siehe
#     _KATEGORIEN_ZUSATZ) - bei "Ja" bleibt die genaue Stufe wie bei
#     Kapitel 5 der manuellen Nacharbeit vorbehalten, bei "Nein" wird
#     automatisch "KINA" gesetzt.
SKIP_FELDER = {"Erkannte_Version", "Python ja/nein", "Systemtyp_CE",
               "Testtiefe", "Testtiefe-Matrix", "Phenix",
               "DI EE-Anforderungen", "Steuerung erfolgt über?",
               "KI-Reifegrad",
               "API", "BE", "Raum", "SAP",
               "Bearbeiter", "PLSTA", "Version",
               "SW-Version / Typ:", "SW-Name:", "Version_Historie",
               "Bemerkung1", "Bemerkung2", "Bemerkung3", "Bemerkung4"}

_TEXTAREA_FELDER = (
    set(template_filler._BESCHREIBUNG_ZEILEN.values())
    | {"Besonderheiten", "GxP_Produktqualitaet", "GxP_Patientensicherheit",
       "GxP_Datenintegritaet", "Kurzbeschreibung", "Historie", "Schnittstelle"}
    | set(common.ROLLEN_SPALTEN)
)

# Abteilungs-Felder (webapp-only, nicht Teil von EXCEL_COLUMNS): jede
# Rollenzeile im Deckblatt hat einen Abteilungs-Platzhalter im Label -
# bei Ersteller/SI-PL/TSO/BSO/BQR "(Site/Unit)", bei CSQ das im
# Template fest eingetragene "(FBC Quality Q&V CSV)" (wird nur
# ersetzt, wenn tatsaechlich ein CSQ_Abteilung-Wert angegeben wird -
# ohne Eingabe bleibt der Standardtext stehen). Werden im Formular
# direkt neben dem jeweiligen Namensfeld angezeigt (siehe
# baue_formular()) und von template_filler.fill_deckblatt_rollen() in
# das Label der jeweiligen Zeile eingesetzt.
ABTEILUNG_FELDER = {
    "Ersteller": "Ersteller_Abteilung",
    "SI/PL":     "SI_PL_Abteilung",
    "TSO":       "TSO_Abteilung",
    "BSO":       "BSO_Abteilung",
    "BQR":       "BQR_Abteilung",
    "CSQ":       "CSQ_Abteilung",
}

# Freundliche Beschriftung fuer Felder, deren interner Name (=
# Master-Excel-Spaltenname, muss unveraendert bleiben, siehe
# EXCEL_COLUMNS in sysbew_common.py) nicht 1:1 als Anzeigetext taugt:
# die 4 generischen "BemerkungX"-Spalten (feste Bedeutung laut
# Fachbereich innerhalb des Kapitels "Informationen und Bemerkungen")
# sowie zwei Namen, die noch ae/oe/ue statt ä/ö/ü aus der Excel-Zeit
# tragen (Excel-Spaltennamen selbst bleiben so - nur die Anzeige im
# Formular wird ausgeschrieben).
FELD_LABELS = {
    "Bemerkung1": "Bemerkung 1 (Prozessbeschreibung)",
    "Bemerkung2": "Bemerkung 2 (Daten)",
    "Bemerkung3": "Bemerkung 3 (Audit Trail)",
    "Bemerkung4": "Bemerkung 4 (Parameter)",
    "Gebaeude": "Gebäude",
    "UeberlagerteMLCS": "Überlagerte MLCS",
    "Dok. -Nr.": "Dok.-Nr. / Version",
    "DokNummerVorQualiPSO": "Vorherige Doc-ID",
}

# Hinweistexte, die im Editor unter dem jeweiligen Feld/der jeweiligen
# Kategorie angezeigt werden - wo moeglich woertlich aus dem Template
# uebernommen. PLSTA und "SW-Version / Typ:" stehen bewusst NICHT hier:
# beide sind reine Master-Excel-Spalten ohne Zelle im Template (siehe
# SKIP_FELDER oben) und werden im Formular gar nicht mehr angezeigt.
FELD_HINWEISE = {
    "SystemtypZugang_Begruendung": (
        "Begründung zur oben gewählten Zugangsbeschränkung - im Template selbst nur bei "
        "„N/A“ mit einem festen Beispieltext („mechanische Ausrüstung“) hinterlegt, wird hier "
        "aber für jede Auswahl angeboten. Wird als zusätzliche Zeile unter die Checkboxen "
        "angehängt."
    ),
    "CSQ_Abteilung": (
        "Nur ausfüllen, falls abweichend von der im Template standardmäßig eingetragenen "
        "„FBC Quality Q&V CSV“ - ohne Eingabe bleibt der Standardtext unverändert."
    ),
    "Dok. -Nr.": (
        "Dokumentennummer UND Version der NEUEN Systembewertung, zusammen in einem Feld - "
        "z. B. „QU-OPE-XXXXX / Version 1.0“. Nicht die Nummer/Version des Vorgänger-Dokuments "
        "(siehe dafür „Vorherige Doc-ID“) - bei einer aus der Datenbank gestarteten Systembewertung "
        "deshalb bewusst leer vorbelegt, nicht die alte Nummer. "
        "Ohne „/ Version ...“ eingegeben, wird automatisch „Version 1.0“ angenommen. "
        "ℹ️ Diese Systembewertung wird auf Basis der im System hinterlegten Vorlage V11 erstellt "
        "(\"Erkannte Version2\" in der Master-Excel) - aktuell die einzige hinterlegte Version, "
        "wird automatisch gesetzt."
    ),
    "DokNummerVorQualiPSO": (
        "Dok.-Nr./Version des VORGÄNGER-Dokuments (z. B. „QU-OPE-XXXXX / Version 1.0“), auf dessen "
        "Bewertung diese neue Systembewertung inhaltlich aufbaut. Bei Start aus der Datenbank "
        "automatisch mit der bisherigen Dok.-Nr./Version befüllt, sofern hier noch nichts stand."
    ),
    "Hyperlink": "Link auf das Dokument in QualiPSO (sobald dort abgelegt).",
    "MLCSID": "System-Identifier/CS-Inventarnummer gemäß QU-SOP-0052370. Keine MLCS erforderlich für S0 und Equipment ohne CS.",
    "UeberlagerteMLCS": "Übergeordnetes System: Systemname, MLCS-ID und ggf. Doc-ID der zugehörigen Systembewertung.",
    "Schnittstelle": (
        "Schnittstelle zu anderen/übergeordneten Systemen: Systemname, MLCS-ID und ggf. Doc-ID der "
        "Systembewertung - z. B. „Profibus DP zum überlagerten BDIS Lantus (MLCS-ID: 1194)“. Für den "
        "Schnittstellen-Typ unten dieselben Vorlagen wie bei „Schnittstellen mit PLS“ verfügbar."
    ),
    "DatenflussAbbildung": (
        "Einfache Darstellung der Datenflüsse - als Grafik direkt im erzeugten Word-Dokument ergänzen "
        "(diese App kann keine Grafiken einfügen), oder per Verweis auf das Feld „Schnittstellen mit PLS“."
    ),
    "Anlage": "Anlagenkennung - z. B. SAP- oder COMOS-Nummer (auch Equipment-Nr./QC-ID möglich).",
    "Schnittstellen mit PLS": (
        "Beschreibung der Datenschnittstelle(n) zu überlagerten Systemen (z. B. Schnittstellen-Typ + "
        "PNK-Nummer). Wurde diese Beschreibung bereits vollständig per Serienbrief befüllt, kann der "
        "Rest des Textes hier gelöscht werden."
    ),
    "AS/BDIS-Name": "Bei Equipment z. B. Laborwaage - zusätzlich den Namen des zugehörigen Systems bzw. die Produkt-Nummer angeben; bei System: Software-/Systemname.",
    "Kurzbeschreibung": "Wozu wird das System/Equipment/die Anlage eingesetzt?",
    "Betrieb": "Einsatzort: Site/Organisationseinheit.",
    "Gebaeude": "Gebäude-Kürzel, falls vorhanden.",
    "Hersteller": "Name Hersteller/Lieferant. Bei zugelassenen Lieferanten die QualiPSO-ID im Feld „Lieferantennummer“ ergänzen.",
    "SW-Hersteller": "Hersteller/Lieferant der Software (falls abweichend vom Equipment-Hersteller).",
    "Lieferantennummer": "QualiPSO-/QTP-Customer-ID des Lieferanten, falls vorhanden.",
    "Historie": "Grund der Erstellung/Änderung - siehe Textbaustein-Vorschläge unten.",
    "CCNr_Rahmen": (
        "Kurzangabe für Kapitel 1 (\"Grund der Systembewertung\") - NICHT der lange "
        "Text aus „Historie“, sondern nur CC-Nr. und der Rahmen der Erstellung, z. B. "
        "„CC-2024-01234 - Periodic Review“ oder „CC-2024-01234 - Requalifizierung“. "
        "Siehe Textbaustein-Vorschläge unten."
    ),
    "Besonderheiten": "Bei Gerätekategorien A, B und C bitte die Subkategorisierung (z. B. B1, C2) nach QU-SOP-0021736 begründen - siehe Textbaustein-Vorschläge unten.",
    "GxP_Produktqualitaet": (
        "Begründung der GxP-Risikoklassifizierung (Produktqualität) - Format je nach oben "
        "gewählter GxP-Kritikalität: „Major, da ...“ / „minor, da ...“ / „N/A, da ...“. "
        "Bei neuen Systemen ohne GxP-Bezug reicht z. B. „N/A, da nicht GMP-relevant.“ - siehe "
        "Textbaustein-Vorschlag unten. Wer die Begründung nicht in 3 einzelne Felder aufteilen "
        "möchte, kann die komplette Begründung hier eintragen und „Patientensicherheit“/"
        "„Datenintegrität“ leer lassen."
    ),
    "GxP_Patientensicherheit": (
        "Begründung der GxP-Risikoklassifizierung (Patientensicherheit) - gleiches Format wie "
        "bei „Produktqualität“. Kann leer bleiben, wenn die gesamte Begründung bereits im Feld "
        "„Produktqualität“ steht."
    ),
    "GxP_Datenintegritaet": (
        "Begründung der GxP-Risikoklassifizierung (Datenintegrität) - gleiches Format wie bei "
        "„Produktqualität“. Kann leer bleiben, wenn die gesamte Begründung bereits im Feld "
        "„Produktqualität“ steht."
    ),
}
# Funktion/Rolle je Person (aus den Bestätigungstexten der Deckblatt-
# Unterschriftentabelle übernommen) - als Hinweis unter dem jeweiligen
# Namensfeld, damit auf Anhieb klar ist, wer z. B. "TSO" ist.
ROLLEN_FUNKTIONEN = {
    "Ersteller": "Autor CSV/SME - bestätigt die inhaltliche Richtigkeit und Vollständigkeit der Systembewertung.",
    "SME": "Subject Matter Expert.",
    "SI/PL": "Systemintegrator/Projektleiter - hat bei der Erstellung der Systembewertung mitgewirkt.",
    "TSO": "Technical System Owner / Leiter Technik / Laborleiter - prüft die Systembewertung auf inhaltliche Richtigkeit und Vollständigkeit.",
    "BSO": "Business System Owner / Leiter Produktion / Herstellungsleiter / Laborleiter - bestätigt, dass die Systembewertung den aktuellen Anforderungen entspricht.",
    "BQR": "Business Quality Representative - bestätigt die Bewertung im Hinblick auf den Einfluss auf die Produktqualität.",
    "CSQ": "Computerized System Quality (FBC Quality Q&V CSV) - bestätigt die aktuellen GMP-Anforderungen und genehmigt Qualifizierung/Validierung.",
}
for _rolle in common.ROLLEN_SPALTEN:
    _funktion = ROLLEN_FUNKTIONEN.get(_rolle, "")
    FELD_HINWEISE.setdefault(
        _rolle,
        f"{_funktion} Bei mehreren Personen mit Zeilenumbruch trennen.".strip(),
    )
for _abteilung_feld in ABTEILUNG_FELDER.values():
    FELD_HINWEISE[_abteilung_feld] = 'Site/Organisationseinheit dieser Person - ersetzt den Platzhalter "(Site/Unit)" im Dokument.'

# Kurzerklaerung je EINZELNER Auswahl-Option (nicht ganze Kategorie) -
# wortnah aus den jeweiligen Kapiteln des Templates uebernommen, damit
# auch ohne Vorwissen erkennbar ist, was z.B. "GAMP 5, Kategorie 3"
# bedeutet. Erscheint als Tooltip (Maus draufhalten) an der jeweiligen
# Option. CS-Typ/Geraetekategorie ergeben sich im Template aus einem
# mehrstufigen Entscheidungsbaum (Kapitel 5) - hier bewusst nur eine
# verkuerzte Orientierung, keine vollstaendige Wiedergabe der Logik.
OPTIONEN_HINWEISE = {
    # GAMP5 Software-Kategorie (Kapitel 7)
    "KAT1": "Infrastruktur-/unterlagerte Software zur Verwaltung der Betriebsumgebung (z. B. Betriebssysteme, Datenbankmanager, Dienste) - selbst nicht konfigurierbar.",
    "KAT3": "Nicht konfigurierbare Software: Laufzeitparameter können eingegeben werden, aber keine Anpassung an den Geschäftsprozess (z. B. Firmware-basierte Systeme).",
    "KAT4": "Konfigurierbare Software, die der Anwender an den Geschäftsprozess anpassen kann, ohne den Quellcode zu ändern (z. B. LIMS, SCADA, ERP, PLS, HMI).",
    "KAT5": "Kundenspezifisch entwickelte/programmierte Software, passend zum jeweiligen Geschäftsprozess.",
    # ERES-Typ (Kapitel 6)
    "ERESTYP1": "Keine elektronischen Aufzeichnungen und keine elektronische Signatur (einfaches CS ohne E-Records).",
    "ERESTYP2": "Keine E-Records/Signatur im System selbst, weil diese über ein übergeordnetes System abgedeckt werden oder GxP-relevante Aufzeichnungen auf Papier erfolgen.",
    "ERESTYP3": "Elektronische Aufzeichnungen vorhanden, aber ohne elektronische Signatur.",
    "ERESTYP4": "Elektronische Aufzeichnungen MIT elektronischer Signatur zur Unterschrift von GxP-relevanten Aufzeichnungen.",
    # KI-Reifegrad (Kapitel 9)
    "KI1": "KI läuft parallel zum Produktionsprozess, ohne Einfluss auf einen GxP-relevanten Prozess.",
    "KI2": "Konventionelle Anwendung ohne Einsatz von Machine Learning.",
    "KI3": "Geschlossenes KI-gestütztes System (kein selbständiges Neutraining).",
    "KI4": "Autonom mit selbst auslösendem Neutraining - mit menschlicher Kontrolle der Updates.",
    "KI5": "Autonom mit selbst auslösendem Neutraining - ohne menschlichen Eingriff (nur Stichproben nach dem Betrieb).",
    "KI6": "Vollständig autonom, optimiert sich selbst anhand eines Ziels bzw. einer Rückkopplungsschleife.",
    # CS-Typ (Kapitel 5, verkürzt)
    "Systemtyp_CIS": "Computerized Information System: reine Software ohne steuerbare Elemente (z. B. SAP, EDS, PRODIS, MES).",
    "Subtyp_LCE": "Laboratory Computerized Equipment: System steuert ein/mehrere Analysegeräte.",
    "Subtyp_PCS": "Process Control System: Prozesskontrolle automatisierter Anlagen bzw. Prozesslenkung/-analyse.",
    "Subtyp_EE": "Electronic Equipment: eigenständiges Gerät mit Mess-/Kontrollfunktion bzw. Firmware/Software (kein CIS/PCS/LCE/S).",
    "VNAP_S0": "Spreadsheet/kleine Applikation, nur einfache Formeln - Verifizierung bei jedem Einsatz.",
    "VNAP_S1": "Spreadsheet/Applikation OHNE Datenspeicherung.",
    "VNAP_S2": "Spreadsheet/Applikation MIT Datenspeicherung.",
    # Gerätekategorie (Kapitel 5, verkürzt)
    "GKATA": "Mechanisches Equipment/Testhilfsmittel ohne Firmware/Software, bzw. Equipment ohne Mess-/Kontrollfunktion.",
    "GKATB": "Firmware-basiertes System, nicht bzw. eingeschränkt parametrierbar (siehe Kapitel 5.8 für B1/B2/B3).",
    "GKATC": "Software-basiertes, konfigurierbares/programmierbares System (siehe Kapitel 5.6-5.8 für C1/C2/C3).",
    # Periodic Review (Kapitel 2)
    "PR_SOP": "QU-SOP-0007359 „Periodic Review für GxP computerisierte Systeme“.",
    "PR_SOP2": (
        "QU-SOP-0028559 „Validierung und Lebenszyklus von GxP-Applikationen“ - "
        "anzuwenden bei Applikationen, also CS-Typ S1 oder S2."
    ),
    "PR-Zyklisch": (
        "QU-SOP-0072260 - regelmäßige Wiederholung der Qualifizierung/Validierung "
        "in festen Intervallen (kein Kreuz bei CIS und Spreadsheets)."
    ),
    # Globale CS-Klasse (Kapitel 3) - Wortlaut der Entscheidungsgrafik
    # aus dem Template selbst uebernommen.
    "KLASS_Global_1a": "„CS erfordert eine lokale Installation“: Nein.",
    "KLASS_Global_1b": "„CS erfordert eine lokale Installation“: Ja.",
    "KLASS_Global_2": "„CS erfordert nur die Anpassung von Stammdaten“.",
    "KLASS_Global_3": "„CS erfordert die Anpassung von Stammdaten und Funktionen“ (zusätzlich Kap. 8 relevant).",
    "KLASS_Global_NA": "System ist kein Globales CS (siehe Auswahl bei „Klassifizierung“ oben) - Kapitel 3 nicht relevant, weiter mit Kap. 4.",
}

KATEGORIE_HINWEISE = {
    "Klassifizierung": "Bei „Globales CS“ zusätzlich die Klasse (1a/1b/2/3) direkt darunter angeben - „Globales CS“ selbst muss dafür nicht extra angekreuzt werden.",
    "Globale CS-Klasse (Kapitel 3)": (
        "Nur relevant, wenn das System oben NICHT als „Lokales CS“, „Multi-Site-CS“ "
        "oder „Equipment ohne CS“ eingestuft wurde. Bei Auswahl einer Klasse wird "
        "„Globales CS“ automatisch mit übernommen (Kapitel 1 + Kapitel 3) - eine "
        "eigene Abfrage dafür entfällt, da sie sich aus der Klasse selbst ergibt. "
        "Trifft keine der 4 Klassen zu, bitte „N/A“ wählen (Kapitel 3 wird dann als "
        "nicht relevant markiert, weiter mit Kap. 4)."
    ),
    "Periodic Review": (
        "Zyklische Requalifizierung gemäß QU-SOP-0072260 (kein Kreuz bei CIS und "
        "Spreadsheets). PR und ReQ (zyklische Requalifizierung) sind nicht immer leicht "
        "zu unterscheiden: Nur PR - reines CS ohne qualifizierungspflichtige Hardware: "
        "ausschließlich PR nach QU-SOP-0007359. Nur ReQ - Equipment ohne GxP-relevanten "
        "CS-Anteil: ausschließlich zyklische Requalifizierung nach QU-SOP-0072260. "
        "Beide - qualifizierungspflichtige Hardware KOMBINIERT mit GxP-relevanter "
        "Software: sowohl ReQ als auch PR (ReQ bewertet die Hardware, PR den CS-Status; "
        "PR wird im Quality Review der ReQ referenziert - QU-SOP-0072260 Kap. 5.3.1; "
        "QU-MT-0009673 Kap. 2.6)."
    ),
    "ERES-Typ": "Diese Auswahl wird zusätzlich in Kapitel 6 des Dokuments übernommen.",
    "GAMP5 Software-Kategorie": "Diese Auswahl wird zusätzlich in Kapitel 7 des Dokuments übernommen.",
    "GxP-Kritikalität": "Diese Auswahl fließt zusammen mit der Software-Kategorie automatisch in die Testtiefe (Kapitel 2 + Kapitel 8) ein.",
}

# ERES-Typ 4 hat im Template eine eigene Unterfrage "Art der Signatur"
# (3 Checkboxen) - webapp-only, nicht Teil von EXCEL_COLUMNS. Nur
# relevant, wenn oben ERES-Typ 4 ausgewaehlt wurde, wird aber unabhaengig
# davon immer angezeigt (kein bedingtes Ein-/Ausblenden).
_KATEGORIEN_ZUSATZ = {
    "ERES-Typ 4 – Art der Signatur": {
        "optionen": [
            ("ERES4_SIG_ID_PW", "Identifikation und Passwort"),
            ("ERES4_SIG_BIOMETRISCH", "Biometrisch"),
            ("ERES4_SIG_TOKEN_PW", "Token und Passwort"),
        ],
        "mehrfachauswahl": True,
    },
    # Ersetzt "KI-Reifegrad" (siehe SKIP_FELDER) als einfache Ja/Nein-
    # Frage (entspricht Kapitel 9.1) - "Nein" setzt beim Erzeugen
    # automatisch KINA (siehe _dokument_erzeugen_und_senden), "Ja"
    # laesst die genaue Stufe (Kapitel 2 KI1-6 + Kapitel 9.2-9.5)
    # bewusst offen fuer die manuelle Nacharbeit.
    "Kommt KI zum Einsatz?": {
        "optionen": [
            ("KI_Einsatz_Ja", "Ja"),
            ("KI_Einsatz_Nein", "Nein"),
        ],
        "mehrfachauswahl": False,
    },
    # Ueberschreibt common.MEHRFACHAUSWAHL_KATEGORIEN["Klassifizierung"]
    # NUR fuer den Web-Editor (siehe _kategorien_lookup() - Zusatz wird
    # zuletzt angewendet): "Globales CS" und die Klasse 1a/1b/2/3 werden
    # hier NICHT mehr mitabgefragt, sondern per eigener, stufiger
    # Kategorie "Globale CS-Klasse (Kapitel 3)" direkt darunter erfasst
    # (siehe baue_formular). sysbew_common.MEHRFACHAUSWAHL_KATEGORIEN
    # selbst bleibt unveraendert, da dort weiterhin ALLE echten
    # Checkboxen (inkl. Globales CS/Klassen) fuer die Konsistenzpruefung
    # beim Einlesen fertiger Dokumente (word_parser_main.py) benoetigt
    # werden.
    "Klassifizierung": {
        "optionen": [
            ("KLASS_Lokal", "Lokales CS"),
            ("KLASS_Multisite", "Multi-Site-CS"),
            ("KLASS_Multisite_NurLokal", "nur lokal"),
            ("KLASS_Multisite_LokalGlobal", "lokal und global"),
            ("KLASS_OhneCS", "Equipment ohne CS"),
        ],
        "mehrfachauswahl": True,
    },
    # Kapitel 3 (Systemeinstufung Globales CS) ist Kapitel 1 stufig
    # nachgelagert: nur relevant, wenn das System ueberhaupt ein
    # Globales CS ist. "Globales CS" selbst wird dabei NICHT separat
    # abgefragt - es ergibt sich automatisch daraus, dass hier eine der
    # 4 Klassen gewaehlt wird (siehe _dokument_erzeugen_und_senden), da
    # eine Klasse ohne "Globales CS" laut Formular gar nicht moeglich
    # ist. Die Klassen selbst sind entweder/oder (Radiobuttons), "N/A"
    # deckt alle anderen Faelle ab (Lokales CS, Multi-Site-CS,
    # Equipment ohne CS) - Kapitel 3 ist dann schlicht nicht relevant.
    "Globale CS-Klasse (Kapitel 3)": {
        "optionen": [
            ("KLASS_Global_1a", "Klasse 1a"),
            ("KLASS_Global_1b", "Klasse 1b"),
            ("KLASS_Global_2", "Klasse 2"),
            ("KLASS_Global_3", "Klasse 3"),
            ("KLASS_Global_NA", "N/A - kein Globales CS"),
        ],
        "mehrfachauswahl": False,
    },
}

# Vorschlags-Textbausteine fuer Historie ("Grund der Erstellung") und
# Besonderheiten - werden per Knopf im Editor in das jeweilige Feld
# eingefuegt (siehe editor.html), koennen danach frei angepasst werden.
HISTORIE_VORLAGEN = [
    {
        "label": "Vorlage 1: Anpassung",
        "text": (
            "Anpassung mit inhaltlicher Überführung der gültigen Systemeinstufung "
            "AS/BDIS für dieses System (Dok.-Nr.: <Vorgänger-Dok.-Nr. / Version>) in "
            "die Systembewertung gemäß QU-MT-0001344 Version 11. Überführung in das "
            "Dokumenten-Managementsystem QualiPSO.\n\n"
            "Dabei erfolgte keine Änderung der gültigen Bewertungen. Ergänzende "
            "Angaben wurden im Kapitel „Informationen und Bemerkungen“ notiert."
        ),
    },
    {
        "label": "Vorlage 2: Neuerstellung (Papier → QualiPSO)",
        "text": (
            "Neuerstellung im Rahmen der Überführung der gültigen Systemeinstufung "
            "AS/BDIS für dieses System (bisher außerhalb von QualiPSO verwaltet, "
            "Dok.-Nr.: <Vorgänger-Dok.-Nr. / Version>) in die Systembewertung gemäß "
            "QU-MT-0001344 Version 11 und in das Dokumenten-Managementsystem "
            "QualiPSO.\n"
            "Dabei erfolgte keine Änderung der gültigen Bewertungen. Ergänzende "
            "Angaben wurden im Kapitel „Informationen und Bemerkungen“ notiert."
        ),
    },
]
BESONDERHEITEN_VORLAGEN = [
    {
        "label": "GxP minor → kein Periodic Review",
        "text": "Da das System mit GxP Kritikalität minor bewertet wurde, ist kein Periodic Review gemäß QU-SOP-0007359 nötig.",
    },
    {
        "label": "Gerätekategorie Bx (allgemein)",
        "text": "Es handelt sich gemäß QU-SOP-0021736 Kapitel 2.4.1 um Gerätekategorie Bx (gemäß Kapitel 5.8).",
    },
    {
        "label": "Vereinfachte Qualifizierung – B1",
        "text": "Das System wird vereinfacht qualifiziert, da es sich gemäß QU-SOP-0021736 um Gerätekategorie B1 handelt (nicht parametrierbares, Firmware-basiertes System).",
    },
    {
        "label": "Vereinfachte Qualifizierung – B3",
        "text": "Das System wird vereinfacht qualifiziert, da es sich gemäß QU-SOP-0021736 um Gerätekategorie B3 handelt (parametrierbares, jedoch nicht konfigurierbares Firmware-basiertes System).",
    },
    {
        "label": "Vereinfachte Qualifizierung – C1",
        "text": "Das System wird vereinfacht qualifiziert, da es sich gemäß QU-SOP-0021736 um Gerätekategorie C1 handelt (parametrierbares, jedoch nicht konfigurierbares Software-basiertes System).",
    },
    # Frueher als fester Anleitungstext im Leer-Template hinterlegt
    # (mehrere Absaetze, die bei jeder Systembewertung unveraendert
    # stehen blieben, statt individuell ausgewaehlt werden zu koennen) -
    # jetzt als eigene Textbaustein-Vorschlaege, analog zu den
    # uebrigen Vorlagen oben.
    {
        "label": "Erstkalibrierung (CE-EE)",
        "text": "Durch die Einstufung in CE-EE ist eine Erstkalibrierung und Funktionsprüfung des EEs durchzuführen. Dies erfolgt über ein projektspezifisches Formblatt.",
    },
    {
        "label": "Bestandssystem (keine neue Qualifizierung)",
        "text": "Es handelt sich um ein Bestandssystem, deshalb sind keine neuen Qualifizierungsaktivitäten notwendig.",
    },
    {
        "label": "Begründung - nicht geschäftskritisch",
        "text": "Begründung angeben, wenn es nicht geschäftskritisch ist.",
    },
    {
        "label": "LCE - Equipment-Anzahl beachten",
        "text": "LCE mit einem und mit mehreren angeschlossenen Equipments betrachten.",
    },
]

# Begruendung der GxP-Risikoklassifizierung (Produktqualitaet/
# Patientensicherheit/Datenintegritaet) - fuer neue Systeme ohne
# GxP-Bezug reicht meist diese Standardformulierung je Feld. Bei
# Critical/Major/minor nur der Formatanfang als Diktier-Starthilfe
# (siehe FELD_HINWEISE["GxP_Produktqualitaet"]) - der Rest der
# Begruendung ist fachlich zu individuell fuer eine feste Vorlage.
_GXP_BEGRUENDUNG_VORLAGEN_BASIS = [
    {"label": "Critical, da ...", "text": "Critical, da "},
    {"label": "Major, da ...", "text": "Major, da "},
    {"label": "minor, da ...", "text": "minor, da "},
    {"label": "N/A - nicht GMP-relevant", "text": "N/A, da nicht GMP-relevant."},
]
GXP_PRODUKTQUALITAET_VORLAGEN = list(_GXP_BEGRUENDUNG_VORLAGEN_BASIS)
GXP_PATIENTENSICHERHEIT_VORLAGEN = list(_GXP_BEGRUENDUNG_VORLAGEN_BASIS)
GXP_DATENINTEGRITAET_VORLAGEN = list(_GXP_BEGRUENDUNG_VORLAGEN_BASIS)

# Vorschlags-Textbausteine fuer die Felder aus Kapitel 2
# "Informationen und Bemerkungen" - je Feld mehrere typische
# Formulierungen. Ausgewaehlt UND priorisiert per Frequenzanalyse
# gegen die vollstaendige Master-Excel (737 Zeilen, Spalten
# Bemerkung1-4 = Prozessbeschreibung/Daten/Audit Trail (AT)/Parameter
# gemaess template_filler._BEMERKUNG_ZUORDNUNG, sowie die direkten
# Felder fuer Alarme/Chargenprotokoll/Benutzerverwaltung/
# Schnittstellen/Equipment/KI) - die jeweils haeufigsten wiederkehrenden
# Formulierungen je Feld sind hier hinterlegt. Werden per Knopf im
# Editor an das Feld angehaengt (siehe editor.html), danach frei
# anpassbar.
# WICHTIG: alle system-/dokumentspezifischen Angaben (Systemname,
# MLCS-ID, SOP-/Dok.-Nr., Fileserver usw.) sind als klar erkennbare
# Platzhalter "<...>" hinterlegt, NICHT als konkretes Beispiel (z.B.
# "PLS Lantus") - sonst wird beim Verwenden leicht uebersehen, dass der
# Wert noch durch den tatsaechlichen des jeweiligen Systems ersetzt
# werden muss (siehe README.md, Versionshistorie).
_VERWEIS_IST_ZUSTAND_VORLAGE = {
    "label": "Verweis Ist-Zustand (e-Records)",
    "text": (
        "Siehe Anlage „Aufnahme des Ist-Zustands“ für dieses System, zur Bewertung "
        "des Punktes „Elektronische Aufzeichnungen (e-Records)“."
    ),
}
PROZESSBESCHREIBUNG_VORLAGEN = [
    {"label": "PLS (Yokogawa/ABB, MIB)", "text": "Das Prozessleitsystem vom Typ Yokogawa Centum VP / ABB 800xA / ABB Freelance dient der Bedienung und Beobachtung der rezeptorientierten verfahrenstechnischen Herstellprozesse im Multi Insulin Betrieb (MIB) im Gebäude G650."},
    {"label": "PLS (frei, Yokogawa)", "text": "Der / Die Prozessleitsystem <Systemname> der Fa. Yokogawa dient <Zweck>."},
    {"label": "Filtertestgerät", "text": "Das Filtertestgerät <Bezeichnung> dient der Prüfung der Integrität diverser Flüssigkeits- und Gasfilter aus dem GxP-relevanten Umfeld der Wirkstoffproduktion. Hierbei handelt es sich um ein im Fermentationslabor fest installiertes Gerät / Hierbei handelt es sich um ein mobiles Gerät inkl. WIT-Trolley."},
    {"label": "SPS (Siemens S7)", "text": "Die Prozess-Automatisierung erfolgt über eine Speicherprogrammierbare Steuerung (SPS) vom Typ S7-300 / S7-400 / S7-1200 / S7-1500 der Firma Siemens, wobei die Bedienung über ein integriertes Touchpanel (HMI) durchgeführt wird."},
    {"label": "SPS für Prozesssteuerung (kurz)", "text": "Speicherprogrammierbare Steuerung (SPS) für die Prozesssteuerung."},
    {"label": "Durchreiche-Maschine", "text": "Die Bedienung der Durchreiche-Maschine erfolgt über zwei baugleiche Bedienpanels (Beladeseite / Entladeseite)."},
    {"label": "Überlagertes BDIS/PLS", "text": "Die Prozess-Automatisierung erfolgt über das <überlagerte BDIS/PLS, z. B. „BDIS Musterlinie“> (MLCS-ID: <ID>). In den Grundlagendokumenten wird die Anlage als Teilanlage <Nr./Bezeichnung> verwaltet."},
    {"label": "Firmware-System (Sartorius/Pall)", "text": "Die automatisierte Durchführung der Filtertests erfolgt über ein Firmware-basiertes System des Herstellers Sartorius / Pall inkl. Bedienpanel."},
    {"label": "Autarkes System (keine Schnittstelle)", "text": "Bezogen auf die Automatisierungstechnik handelt es sich um ein autarkes System ohne Datenschnittstellen zu überlagerten Systemen (z.B. PLS)."},
    {"label": "Rein mechanisch (kein CS)", "text": "Beim / Bei der Prozessleitsystem <Systemname> handelt es sich um rein mechanisches Equipment ohne Computergestütztes System."},
    {"label": "Mechanisches Equipment (an PLS angebunden)", "text": "Die Anlage wird als mechanisches Equipment eingestuft, weil das System ohne eigene Steuerung als Equipment an das überlagerte <PLS, z. B. „PLS Musterlinie“> (MLCS-ID: <ID>) angebunden ist."},
    {"label": "Bestandsaggregat (bereits qualifiziert)", "text": "Es handelt sich um ein bereits qualifiziertes Bestandsaggregat."},
    {"label": "Bedien-SOP (Verweis)", "text": "Die organisatorischen Prozesse inkl. Bedienung des / der <Anlage/System> werden über die <SOP-Nr.> (Bedien-SOP) spezifiziert."},
    _VERWEIS_IST_ZUSTAND_VORLAGE,
]
DATEN_VORLAGEN = [
    {"label": "Verweis auf Dokument", "text": "Der Umgang mit den GxP-relevanten Daten des / der <Anlage/System> ist im Dokument <Dok.-Nr.> spezifiziert."},
    {"label": "Keine dauerhafte Speicherung (→ BDIS)", "text": "Es werden keinerlei GxP-relevante Bewegungsdaten (Transactional Data) dauerhaft auf dem System gespeichert. Die Bewegungsdaten (z.B. Trendkurven) werden direkt an das überlagerte <BDIS/PLS, z. B. „BDIS Musterlinie“> (MLCS-ID: <ID>) übertragen und dort weiter verarbeitet."},
    {"label": "SPS überträgt Signale an PLS (Profibus DP)", "text": "Die SPS überführt via Profibus DP ausgewählte analoge / binäre Signale und Alarme / Meldungen an das bauseitig vorhandene <PLS, z. B. „PLS Musterlinie“> (MLCS-ID: <ID>) zur dauerhaften Speicherung."},
    {"label": "Digitales Maschinenlogbuch", "text": "Digitales Maschinenlogbuch & Chargendaten Speicherort: lokal und <Server, z. B. DASP-Server>."},
    {"label": "Audit Trail & Chargendaten (lokal + Server)", "text": "Audit Trail & Chargendaten Speicherort: lokal (Zugriffsgeschützt) und <Server, z. B. DASP-Server>."},
    {"label": "Systemsoftware nicht konfigurierbar (Messsystem)", "text": "Die Systemsoftware (z.B. Messsystemsteuerung, Methoden) ist nicht konfigurierbar. Messmethoden sind im Messsystem fest hinterlegt und auswählbar (nur änderbar über Hersteller)."},
    {"label": "Keine Speicherung, nur Papierausdruck", "text": "Es werden keine Daten in der Datei gespeichert. Nach abgeschlossener Berechnung erfolgt ein Ausdruck auf Papier."},
    {"label": "Ringspeicher (10 Tage)", "text": "Das System besitzt einen Ringspeicher, mit dem die Bewegungsdaten (Transactional Data) für 10 Tage lokal gespeichert werden. Dadurch können im Falle einer Störung am / an der <Anlage/System> über einen Ersatzprozess die Bewegungsdaten nachträglich übertragen werden."},
    {"label": "Transfer über USB-Medium", "text": "Da keine Datenschnittstellen vorhanden sind / genutzt werden, erfolgt der Datentransfer über ein freigegebenes USB-Medium gemäß QU-SOP-4711 / betrieblicher Anweisung <Dok.-Nr.>."},
    {"label": "Keine GxP-relevanten Daten", "text": "Das System erzeugt keine GxP-relevanten Daten."},
    {"label": "Archivierung gemäß SOP", "text": "Die Archivierung / Dearchivierung von Bewegungsdaten (Transactional Data) erfolgt über den etablierten Prozess gemäß QU-SOP-0051759 „Planung und Durchführung der Archivierung / Rückübertragung bei Automatisierungs- und Betriebsdatenerfassungssystemen“."},
    {"label": "Backup/Restore gemäß SOP", "text": "Die Konfigurationsdaten (Master Data) des Steuerungssystems können übergangsweise über den etablierten Backup / Restore-Prozess gemäß alter QU-SOP-0017807 „Planung und Durchführung des Backup / Restore von Automatisierungs- und Betriebsdaten-informationssystemen“ wieder hergestellt werden. Im Rahmen der zyklischen Requalifizierung soll das Backup/Restore-Konzept an die neue SOP QU-SOP-0077934 „Campus Frankfurt Backup & Restore“ angepasst werden."},
    _VERWEIS_IST_ZUSTAND_VORLAGE,
]
PARAMETER_VORLAGEN = [
    {"label": "Nur erhöhte Benutzerrechte (TM-Prozess)", "text": "Die Parametrierung der Produktionsprozesse ist ausschließlich für Personen mit erhöhten Benutzerrechten möglich (siehe dazu unter „Benutzerverwaltung“). Änderungen werden über den etablierten TM-Prozess gemäß QU-SOP-0041022 „Maßnahmen an technischen Einrichtungen (TM)“ abgebildet."},
    {"label": "Vorqualifizierte Schrittketten", "text": "Die Prozessierung der Anlagenprozesse erfolgt über vorqualifizierte Schrittketten / Methoden. Diese Schrittketten / Methoden sind fest parametriert und können nur mit administrativen Rechten angepasst werden, was über den etablierten TM-Prozess / CC-Prozess abgebildet wird."},
    {"label": "Rezeptdaten (lokal + Backup)", "text": "Rezeptdaten Speicherort: lokal (Zugriffsgeschützt) und Datensicherung (<Backup-System, z. B. ACRONIS>)."},
    {"label": "Gemäß Funktionsspezifikation PLS", "text": "Die Parametrierung der verschiedenen Softwareobjekte und -Funktionen erfolgt gemäß den Funktionsspezifikationen der PLS-Teilanlagen."},
    {"label": "Fest hinterlegt, nicht parametrierbar", "text": "Die Anlagenprozesse sind im System fest hinterlegt und können nicht parametriert werden."},
    {"label": "Spezifikation über Bedien-SOP", "text": "Die Spezifikation der Prozessparameter erfolgt über die Bedien-SOP <SOP-Nr.> / die Funktionsspezifikation mit der Dok.-Nr. <Dok.-Nr.>."},
    {"label": "Backup/Restore gemäß SOP", "text": "Im Falle einer schwerwiegenden Anlagenstörung, können die Parameter übergangsweise über den etablierten Backup / Restore-Prozess gemäß alter QU-SOP-0017807 „Planung und Durchführung des Backup / Restore von Automatisierungs- und Betriebsdaten-informationssystemen“ wieder hergestellt werden. Im Rahmen der zyklischen Requalifizierung soll das Backup/Restore-Konzept an die neue SOP QU-SOP-0077934 „Campus Frankfurt Backup & Restore“ angepasst werden."},
    {"label": "Q-Bericht (Referenz)", "text": "Q-Bericht: <Nr.> vom <Datum>."},
    _VERWEIS_IST_ZUSTAND_VORLAGE,
]
ALARME_VORLAGEN = [
    {"label": "Alarmliste im Rahmen der Qualifizierung", "text": "Alarme werden im Rahmen der Qualifizierung in einer Alarmliste definiert und eingestuft."},
    {"label": "Keine Alarme aufgezeichnet", "text": "Mit diesem System werden keine Alarme aufgezeichnet oder verarbeitet."},
    {"label": "Meldung über HMI/Statusampel", "text": "Die GxP-relevanten Alarme werden vom Automatisierungssystem (SPS / HMI) erzeugt und über das lokale Bedienpanel (HMI) gemeldet / Meldeleuchten signalisiert / die in Anlagennähe montierte Statusampel angezeigt."},
    {"label": "Bildung aus BDIS-Messwerten", "text": "Die GxP-relevanten Alarme werden anhand der an das <BDIS/PLS, z. B. „BDIS Musterlinie“> (MLCS-ID: <ID>) übertragenen Messwerte / Signale gebildet. Die Verwaltung und Speicherung der Alarmarchive erfolgt exklusiv über das PLS."},
    {"label": "Übertragung an BDIS (Profibus/Modbus)", "text": "Die GxP-relevanten Alarme werden vom System selbst erzeugt und über Profibus DP / Modbus an das <BDIS/PLS, z. B. „BDIS Musterlinie“> (MLCS-ID: <ID>) übertragen."},
    {"label": "Dokumentation im Chargenprotokoll", "text": "Die GxP-relevanten Alarme werden bei jedem Lauf direkt auf dem papierbasierten Chargenprotokoll dokumentiert."},
    {"label": "Spezifiziert in Bedien-SOP", "text": "Der Umgang mit den Alarmierungen wird in der Bedien-SOP <SOP-Nr.> der Anlage spezifiziert."},
    {"label": "Betriebsspezifisch gemäß SOP", "text": "Der Umgang mit den Alarmierungen wird betriebsspezifisch über die QU-SOP-0015470 „Handling von Alarmen und Meldungen“ festgelegt (Dok.-Nr. <Dok.-Nr.>)."},
]
CHARGENPROTOKOLL_VORLAGEN = [
    {"label": "Papierbasiert (HAW)", "text": "Es werden keine elektronischen Chargenprotokolle mit dem System erzeugt. Es ist ein papierbasierter Prozess über Herstellanweisungen (HAW) etabliert."},
    {"label": "Keine Chargenprotokolle", "text": "Das System erzeugt keine Chargenprotokolle."},
    {"label": "Audit Trail & Chargendaten (lokal + Server)", "text": "Audit Trail & Chargendaten Speicherort: lokal (Zugriffsgeschützt) und <Server, z. B. DASP-Server>."},
    {"label": "PDR über überlagertes BDIS", "text": "Das <PLS, z. B. „PLS Musterlinie“> stellt Chargendaten dem überlagerten <BDIS, z. B. „BDIS Musterlinie“> (MLCS-ID: <ID>) zur Verfügung, welches chargenbezogene Prozessdatenreports (PDR) erzeugt."},
]
AUDIT_TRAIL_VORLAGEN = [
    {"label": "Audit Trail vorhanden (Qualifizierung)", "text": "Die Anlage verfügt über einen Audit Trail. Das ATR-Konzept wird im Rahmen der Qualifizierung erarbeitet."},
    {"label": "Maschinenlog (Qualifizierung)", "text": "Die Anlage verfügt über einen Maschinenlog. Das ATR-Konzept wird bei Bedarf im Rahmen der Qualifizierung erarbeitet."},
    {"label": "Alarm-Handling über Funktionsspezifikation", "text": "Der Punkt „Handling von Alarmen und Meldungen (<SOP-Nr., z. B. FRA-SOP-03888>)“ aus dem Kapitel 6.0 „Maßnahmenfestlegung für AS/BDIS“ wird über die Funktionsspezifikation der Teilanlage (siehe Kapitel 2.0 „Kurzbeschreibung“) über das <PLS, z. B. „PLS Musterlinie“> (MLCS-ID: <ID>) erfüllt."},
    {"label": "Alarm-Handling über gekoppeltes PLS", "text": "Zum Punkt „Handling von Alarmen und Meldungen (<SOP-Nr., z. B. FRA-SOP-03888>)“: Die SPS ist an das PLS mit der MLCS-ID <ID> gekoppelt und überträgt Prozessdaten sowie Alarme und Meldungen. Die Klassifizierung und Alarmierung erfolgt somit über die Dokumente zum PLS."},
    {"label": "Archivierung in Papierform", "text": "Die Archivierung der relevanten Daten (z.B. Analysenergebnisse) erfolgt in Papierform."},
    {"label": "Kein GxP-Einfluss (System unrelevant)", "text": "Das System hat keinen Einfluss auf Produktqualität, Patientensicherheit oder Datenintegrität und ist dementsprechend GxP unrelevant."},
    {
        "label": "Anwender- + administrativer Audit Trail",
        "text": (
            "Das System bietet einen Audit-Trail, welcher alle Tätigkeiten der Anwender "
            "aufzeichnet, sowie einen administrativen Audit-Trail, welcher alle Tätigkeiten "
            "des Administrators aufzeichnet. Der Audit-Trail kann von niemandem gelöscht "
            "oder deaktiviert werden. Der Audit Trail Review wird gemäß <SOP-Nr.> durchgeführt."
        ),
    },
    {"label": "Verweis auf Einstufungsdokumente", "text": "Die Festlegungen zum Audit Trail / Audit Trail Review sind für das System in den Dokumenten <Dok.-Nr.> „Einstufung von computergestützten Systemen zu E-Record, systemgenerierter Audit Trail und Audit Trail Review“ und <Dok.-Nr.> „Festlegungen der Rahmenbedingungen zum Audit Trail Review“ und <Dok.-Nr.> „Festlegungen zum Audit-Trail-Review von Bewegungsdaten im <System>“ spezifiziert."},
    {"label": "Systemgenerierter AT (Robocopy-Auslagerung)", "text": "Das System besitzt am Bedienpanel (HMI) einen systemgenerierten Audit Trail, der nicht deaktiviert werden kann. Die dabei anfallenden Audit Trail Dateien werden über eine automatisierte Routine (Robocopy-Befehl) zyklisch auf den Fileserver <Fileserver-Name> ausgelagert."},
    {"label": "Kein AT (Review über Filtrationsprotokoll)", "text": "Das System besitzt keinen Audit Trail. Der Audit Trail Review erfolgt anhand des Filtrationsprotokolls, welches zu jeder Herstellanweisung (HAW) in Papierform dazu geheftet wird."},
    {"label": "Gemäß SOP QU-SOP-0038830", "text": "Die Audit Trail / Audit Trail Review Prozesse werden gemäß QU-SOP-0038830 „Vorgehensweise zur Definition prozess- und systemspezifischer Audit Trail Review Konzepte“ durchgeführt."},
]
BENUTZERVERWALTUNG_VORLAGEN = [
    {"label": "Personenbezogen gemäß SOP", "text": "Es wird eine personenbezogene Benutzerverwaltung gemäß QU-SOP-0020358 „GxP-Anforderungen an die Zugriffsverwaltung computergestützter Systeme“ implementiert."},
    {"label": "Domäne + lokale Notfallverwaltung", "text": "Die Benutzerverwaltung erfolgt über die Domäne. Eine lokale Benutzerverwaltung ist vorhanden und wird für Notfälle verwendet."},
    {"label": "Zentrale Benutzerverwaltung (Domäne)", "text": "Das System verfügt über eine zentrale Benutzerverwaltung, die durch die Administrativen Einheiten betreut wird. Das System verwendet die Domänen Accounts in der Applikation. Im Rahmen der Systemvalidierung wurde ein Berechtigungskonzept zur Konfiguration der Benutzerverwaltung etabliert (siehe <Dok.-Nr.>)."},
    {"label": "Zugriffsberechtigungen separat spezifiziert", "text": "Der Punkt „Zugangs-/Zugriffsberechtigungen (<SOP-Nr., z. B. FRA-SOP-03022>)“ aus Kapitel 6.0 „Maßnahmenfestlegung für AS/BDIS“ wird nicht über die SOP-Anlagen dokumentiert, sondern über eine separate Spezifikation, da es sich um Einzel-/Gruppen-Zugangsberechtigungen handelt."},
    {"label": "Verweis auf Benutzeranforderungen", "text": "Die Anforderungen zu Benutzerverwaltung, Systemsicherheit, Virenschutz und Patch-Management können detailliert den Benutzeranforderungen des Systems entnommen werden (Dok.-Nr. <Dok.-Nr.>)."},
    {"label": "Verknüpft mit Pharma-Domäne (I-/DE-Nummer)", "text": "Die lokale Benutzerverwaltung des Systems wird mit einem Anmeldeserver in der Pharma-Domäne von Sanofi verknüpft. Dadurch ist es möglich, sich mit dem Office-Account über die I-Nummer / DE-Nummer am System anzumelden."},
    {"label": "Lokale Notfallaccounts", "text": "Für den Fall, dass der Anmeldeserver nicht verfügbar ist, werden lokale Notfallaccounts installiert, so dass die Anlage jederzeit bedienbar bleibt."},
    {"label": "Gruppenpasswörter", "text": "Es wird eine auf Gruppenpasswörtern basierende Benutzerverwaltung implementiert."},
    {"label": "Technisch nicht möglich", "text": "Die Implementierung einer Benutzerverwaltung ist aus technischen Gründen nicht möglich."},
]
SCHNITTSTELLEN_PLS_VORLAGEN = [
    {"label": "USB organisatorisch verhindert", "text": "Die Nutzung der physikalischen Schnittstellen (USB) ist organisatorisch verhindert."},
    {"label": "Keine Schnittstellen (initiale Implementierung)", "text": "Im Rahmen der initialen Implementierung des Systems werden keine Schnittstellen zum Datenaustausch mit anderen Systemen etabliert. Die Equipments werden über vom Hersteller getestete Schnittstellen an das <System> angebunden."},
    {"label": "Profibus DP", "text": "Das System ist über eine Profibus DP Schnittstelle mit dem überlagerten <BDIS/PLS, z. B. „BDIS Musterlinie“> (MLCS-ID: <ID>) gekoppelt."},
    {"label": "Modbus TCP", "text": "Das System ist über eine Modbus TCP-Schnittstelle mit dem überlagerten <BDIS/PLS, z. B. „BDIS Musterlinie“> (MLCS-ID: <ID>) gekoppelt."},
    {"label": "OPC UA", "text": "Das System ist über eine OPC UA-Schnittstelle mit dem überlagerten <BDIS/PLS, z. B. „BDIS Musterlinie“> (MLCS-ID: <ID>) gekoppelt."},
    {"label": "Remote I/O", "text": "Das System ist über eine Remote I/O-Schnittstelle mit dem überlagerten <BDIS/PLS, z. B. „BDIS Musterlinie“> (MLCS-ID: <ID>) gekoppelt."},
    {"label": "Verweis auf Funktionsspezifikation", "text": "Die vollständige und detaillierte Darstellung aller Datenschnittstellen kann der Funktionsspezifikation des Systems entnommen werden (Dok.-Nr. <Dok.-Nr.>)."},
    {"label": "Keine dauerhafte Nutzung", "text": "Es werden keine Datenschnittstellen des Systems für dauerhaften Datenaustausch genutzt."},
    {"label": "Keine Datenschnittstellen", "text": "Das System besitzt keine Datenschnittstellen."},
]
ANGESCHLOSSENES_EQUIPMENT_VORLAGEN = [
    {"label": "Kein Equipment angeschlossen", "text": "Kein untergeordnetes Equipment vorhanden/angeschlossen."},
    {"label": "Barcode-Scanner (Honeywell)", "text": "Zur Erfassung der Filter-ID ist das Filtertestgerät mit einem Barcode-Scanner der Fa. Honeywell verbunden."},
    {"label": "Plattformwaage (Sartorius, Profibus DP)", "text": "Zur Erfassung des Gewichts der verarbeiteten Kleingebinde ist das System über Profibus DP mit einer Plattformwaage der Fa. Sartorius verbunden."},
    {"label": "Clamp-On-Durchflussmessung (Sartorius)", "text": "Die Clamp-On-Durchflussmessung ist über eine codierte Steckverbindung mit der dazugehörigen Auswerte-Einheit der Fa. Sartorius verbunden."},
]
SONSTIGES_VORLAGEN = [
    {
        "label": "Vorgängerdokument wird ungültig",
        "text": (
            "Mit Freigabe der vorliegenden Systembewertung wird die Systemeinstufung "
            "(Dok.-Nr.: <Vorgänger-Dok.-Nr. / Version>) ungültig. Die Bewertung des "
            "Equipments erfolgt neu gemäß den Vorgaben von QU-SOP-0021736 "
            "(Qualifizierung von Gebäuden, Einrichtungen und Ausrüstung) und "
            "QU-SOP-0049866 (Validierung computergestützter Systeme)."
        ),
    },
    {"label": "N/A (kein Vorgängerdokument)", "text": "N/A"},
]
KI_BEWERTUNG_VORLAGEN = [
    {"label": "Keine KI-Fähigkeiten", "text": "Das System besitzt gemäß Bewertung in Kapitel 9 keine KI-Fähigkeiten."},
]
DATENFLUSS_ABBILDUNG_VORLAGEN = [
    {"label": "Siehe Schnittstellen / Grafik einfügen", "text": "siehe unter „Schnittstellen“ (oder Grafik einfügen)"},
]

# Vereinheitlichte Zuordnung Feldname -> Vorlagenliste, damit editor.html
# die Textbaustein-Knopfleiste generisch (ein Codepfad fuer alle Felder)
# statt mit einem eigenen if/elif-Zweig je Feld rendern kann.
TEXTBAUSTEIN_VORLAGEN = {
    "Historie": HISTORIE_VORLAGEN,
    "Besonderheiten": BESONDERHEITEN_VORLAGEN,
    "GxP_Produktqualitaet": GXP_PRODUKTQUALITAET_VORLAGEN,
    "GxP_Patientensicherheit": GXP_PATIENTENSICHERHEIT_VORLAGEN,
    "GxP_Datenintegritaet": GXP_DATENINTEGRITAET_VORLAGEN,
    "Prozessbeschreibung": PROZESSBESCHREIBUNG_VORLAGEN,
    "Daten": DATEN_VORLAGEN,
    "Parameter": PARAMETER_VORLAGEN,
    "Alarme (GxP-relevant)": ALARME_VORLAGEN,
    "Chargenprotokoll": CHARGENPROTOKOLL_VORLAGEN,
    "Audit Trail (AT)": AUDIT_TRAIL_VORLAGEN,
    "Benutzer-verwaltung?": BENUTZERVERWALTUNG_VORLAGEN,
    "Schnittstellen mit PLS": SCHNITTSTELLEN_PLS_VORLAGEN,
    # "Schnittstelle" (Zusammenfassungstabelle Kapitel 2) - dieselben
    # Schnittstellen-Typ-Vorlagen wie "Schnittstellen mit PLS", da
    # beide Felder inhaltlich denselben Sachverhalt beschreiben (kurz
    # vs. ausfuehrlich) und dieselben Formulierungen passen.
    "Schnittstelle": SCHNITTSTELLEN_PLS_VORLAGEN,
    "Angeschlossenes Equipment": ANGESCHLOSSENES_EQUIPMENT_VORLAGEN,
    "Sonstiges": SONSTIGES_VORLAGEN,
    "KI Bewertung": KI_BEWERTUNG_VORLAGEN,
    "DatenflussAbbildung": DATENFLUSS_ABBILDUNG_VORLAGEN,
}

def _kategorien_lookup():
    lookup = {}
    for name, optionen in VALIDATION_KATEGORIEN_V11:
        if name == "Testtiefe-Matrix":
            continue
        lookup[name] = {"optionen": optionen, "mehrfachauswahl": False}
    for name, optionen in common.MEHRFACHAUSWAHL_KATEGORIEN:
        lookup[name] = {"optionen": optionen, "mehrfachauswahl": True}
    for name, optionen in common.ANZEIGE_ZUSATZ_KATEGORIEN:
        if name == "Testtiefe":
            continue
        # "Validierung/Qualifizierung nach SOP": QUAL und VAL sind kein
        # entweder/oder - laut Fachbereich koennen beide SOPs
        # gleichzeitig zutreffen, deshalb Mehrfachauswahl (Checkboxen)
        # statt Radiobuttons.
        mehrfachauswahl = name == "Validierung/Qualifizierung nach SOP"
        lookup[name] = {"optionen": optionen, "mehrfachauswahl": mehrfachauswahl}
    lookup.update(_KATEGORIEN_ZUSATZ)
    return lookup

def _dok_nr_version_kombinieren(data):
    """Baut den Anzeigewert des gemeinsamen "Dok. -Nr."-Feldes aus den
    getrennt gespeicherten Werten "Dok. -Nr." und "Version_Historie",
    z.B. "QU-OPE-XXXXX / Version 1.0"."""
    dok_nr = (data.get("Dok. -Nr.") or "").strip()
    version = (data.get("Version_Historie") or "").strip()
    if dok_nr and version:
        return f"{dok_nr} / Version {version}"
    return dok_nr or (f"Version {version}" if version else "")

def _dok_nr_version_auftrennen(text):
    """Gegenstueck zu _dok_nr_version_kombinieren(): liest den
    eingegebenen Text des gemeinsamen Feldes wieder in die getrennten
    Werte "Dok. -Nr."/"Version_Historie" ein - die Master-Excel-Spalte
    "Dok. -Nr." erwartet weiterhin nur die reine Nummer, ohne
    Versionszusatz. Ohne erkennbare "/ Version ..."-Angabe wird
    "Version_Historie" NICHT angetastet (bleibt beim bisherigen Wert,
    z.B. "1.0" bei einer neuen Systembewertung), damit ein einfaches
    Weglassen der Version keine bereits gesetzte Version loescht."""
    text = (text or "").strip()
    m = re.match(r"^(.*?)\s*/\s*Version\s*(.+?)\s*$", text, re.IGNORECASE)
    if m:
        return {"Dok. -Nr.": m.group(1).strip(), "Version_Historie": m.group(2).strip()}
    return {"Dok. -Nr.": text}

def _neues_dokument_aus_db_zeile(row):
    """Baut das Start-Dict fuer eine NEUE Systembewertung aus einer
    Datenbank-Zeile (fuer Weg 1 "Direkt erzeugen" UND Weg 2
    "Bearbeiten"): uebernimmt alle Werte der Zeile (Equipment, Rollen,
    GxP-Bewertung usw. bleiben ja unveraendert gueltig), setzt aber die
    Identitaet des NEUEN Dokuments selbst zurueck - Dok.-Nr./Version/
    Datum/Historie-Text gehoeren zum VORGAENGER-Dokument, nicht zum
    neuen. Ohne dieses Zuruecksetzen wuerde sonst versehentlich die
    alte Dok.-Nr. unveraendert weiterlaufen bzw. die alte Version in
    die Historie-Tabelle des neuen Dokuments geschrieben (siehe
    template_filler.fill_historie).

    Die alte Dok.-Nr./Version geht dabei nicht verloren: sie wandert
    automatisch in "DokNummerVorQualiPSO" ("Vorherige Doc-ID"), sofern
    dort noch nichts anderes steht.

    Ausserdem werden die 4 generischen "BemerkungX"-Spalten (falls in
    der Zeile befuellt) HIER bereits in ihr jeweiliges Hauptfeld
    (Prozessbeschreibung/Daten/Audit Trail (AT)/Parameter, siehe
    template_filler._BEMERKUNG_ZUORDNUNG) gemergt und aus `data`
    entfernt - dadurch gibt es im Editor nur noch EIN bearbeitbares
    Feld statt zwei getrennter Werte, die am Ende doppelt im
    generierten Dokument auftauchen wuerden (BemerkungX wird im
    Formular selbst nie angezeigt/abgefragt, siehe SKIP_FELDER - ohne
    dieses Mergen waere der Inhalt sonst schlicht verloren bzw. wuerde
    unbemerkt an das Hauptfeld angehaengt, ohne dass die/der
    Bearbeitende das im Editor sehen/anpassen koennte)."""
    data = {k: v for k, v in row.items() if k != "_zeile"}
    for bemerkung_feld, ziel_feld in template_filler._BEMERKUNG_ZUORDNUNG.items():
        zusatz = data.pop(bemerkung_feld, None)
        if zusatz:
            bestehend = data.get(ziel_feld) or ""
            data[ziel_feld] = f"{bestehend}\n{zusatz}" if bestehend else zusatz
    alte_dok_id = _dok_nr_version_kombinieren(data)
    if alte_dok_id and not data.get("DokNummerVorQualiPSO"):
        data["DokNummerVorQualiPSO"] = alte_dok_id
    data["Dok. -Nr."] = ""
    data["Version_Historie"] = "1.0"
    data["Datum"] = _heute()
    data["Historie"] = ""
    return data

def baue_formular(data):
    """Baut die Formularstruktur fuer editor.html aus den
    VORSCHAU_ABSCHNITTE (dieselbe Gliederung wie die
    Konsolen-Vorschau von word_parser_main.py - Deckblatt zuerst,
    dann Kapitel 1/2 usw.), ergaenzt um webapp-only Zusatzfelder
    (Abteilungen, Periodic-Review-Freitext, ERES-Typ-4-Signatur) und
    Hinweistexte."""
    kategorien = _kategorien_lookup()
    abschnitte = []
    for titel, felder in common.VORSCHAU_ABSCHNITTE:
        items = []
        for feld in felder:
            if feld in SKIP_FELDER:
                continue
            if feld in kategorien:
                info = kategorien[feld]
                ausgewaehlt = [f for f, _ in info["optionen"] if data.get(f) == "r"]
                items.append({
                    "art": "kategorie", "name": feld,
                    "mehrfachauswahl": info["mehrfachauswahl"],
                    "optionen": info["optionen"], "ausgewaehlt": ausgewaehlt,
                    "hinweis": KATEGORIE_HINWEISE.get(feld),
                })
                if feld == "Periodic Review":
                    items.append({
                        "art": "feld", "name": "PR_Andere_Text", "typ": "text", "breite": "kompakt",
                        "wert": data.get("PR_Andere_Text") or "",
                        "label": "Freie Angabe - Details",
                        "hinweis": 'Nur ausfüllen, wenn oben „andere/freie Angabe“ ausgewählt wurde.',
                    })
                if feld == "Systemtyp (Zugangsbeschränkung)":
                    items.append({
                        "art": "feld", "name": "SystemtypZugang_Begruendung", "typ": "text", "breite": "voll",
                        "wert": data.get("SystemtypZugang_Begruendung") or "",
                        "label": "Begründung Systemtyp/Zugangsbeschränkung (optional)",
                        "hinweis": FELD_HINWEISE.get("SystemtypZugang_Begruendung"),
                    })
                if feld == "Klassifizierung":
                    zusatz = _KATEGORIEN_ZUSATZ["Globale CS-Klasse (Kapitel 3)"]
                    items.append({
                        "art": "kategorie", "name": "Globale CS-Klasse (Kapitel 3)",
                        "mehrfachauswahl": zusatz["mehrfachauswahl"],
                        "optionen": zusatz["optionen"],
                        "ausgewaehlt": [f for f, _ in zusatz["optionen"] if data.get(f) == "r"],
                        "hinweis": KATEGORIE_HINWEISE.get("Globale CS-Klasse (Kapitel 3)"),
                    })
                if feld == "ERES-Typ":
                    zusatz = _KATEGORIEN_ZUSATZ["ERES-Typ 4 – Art der Signatur"]
                    items.append({
                        "art": "kategorie", "name": "ERES-Typ 4 – Art der Signatur",
                        "mehrfachauswahl": zusatz["mehrfachauswahl"],
                        "optionen": zusatz["optionen"],
                        "ausgewaehlt": [f for f, _ in zusatz["optionen"] if data.get(f) == "r"],
                        "hinweis": "Nur relevant, wenn oben ERES-Typ 4 ausgewählt wurde.",
                    })
                if feld == "GAMP5 Software-Kategorie":
                    zusatz = _KATEGORIEN_ZUSATZ["Kommt KI zum Einsatz?"]
                    items.append({
                        "art": "kategorie", "name": "Kommt KI zum Einsatz?",
                        "mehrfachauswahl": zusatz["mehrfachauswahl"],
                        "optionen": zusatz["optionen"],
                        "ausgewaehlt": [f for f, _ in zusatz["optionen"] if data.get(f) == "r"],
                        "hinweis": (
                            "Ersetzt die Detailfragen 9.2-9.5 (verbotene Praktiken, Autonomie-/"
                            "Steuerungsdesign-Stufe, Reifegrad) - deren genaue Einstufung lässt "
                            "sich ohne Entscheidungsbaum nicht zuverlässig abfragen und muss bei "
                            "„Ja“ nach dem Erzeugen manuell in Kapitel 2 (KI-Reifegrad) und "
                            "Kapitel 9 ergänzt werden (wie Kapitel 5)."
                        ),
                    })
                    items.append({
                        "art": "feld", "name": "KI_Einsatz_Begruendung", "typ": "textarea",
                        "breite": "voll", "rows": 3,
                        "wert": data.get("KI_Einsatz_Begruendung") or "",
                        "label": "Begründung (kein KI-Einsatz)",
                        "hinweis": (
                            "Nur relevant, wenn oben „Nein“ gewählt wurde - wird beim Erzeugen "
                            "an das Feld „KI Bewertung“ (Informationen und Bemerkungen) "
                            "angehängt."
                        ),
                    })
            else:
                # "Breite" der Feldgruppe im CSS-Raster: grosse Freitext-
                # Felder bekommen die volle Breite, kurze Werte (auch
                # die Rollen-Namen, obwohl technisch als Textarea
                # gerendert) stehen platzsparend nebeneinander.
                ist_textarea = feld in _TEXTAREA_FELDER
                breite = "voll" if (ist_textarea and feld not in common.ROLLEN_SPALTEN) else "kompakt"
                items.append({
                    "art": "feld", "name": feld,
                    "typ": "textarea" if ist_textarea else "text",
                    "breite": breite,
                    # Historie/Besonderheiten koennen ueber die Textbaustein-
                    # Vorschlaege laenger werden - mehr Platz von Anfang an.
                    # Rollen-Namen bleiben trotz Textarea (fuer Mehrfach-
                    # nennung mit "\n" getrennt, siehe extract_deckblatt_
                    # rollen) einzeilig wie die Abteilungsfelder - wer
                    # wirklich zwei Namen braucht, kann per Ziehgriff
                    # vergroessern.
                    "rows": (
                        8 if feld in ("Historie", "Besonderheiten") else
                        1 if feld in common.ROLLEN_SPALTEN else 3
                    ),
                    # "Dok. -Nr." zeigt/erfasst zusammen mit der Version in
                    # EINEM Feld (siehe _dok_nr_version_kombinieren) - das
                    # ist inhaltlich "eine Doc-ID" statt zwei getrennter
                    # Angaben.
                    "wert": _dok_nr_version_kombinieren(data) if feld == "Dok. -Nr." else (data.get(feld) or ""),
                    "label": FELD_LABELS.get(feld),
                    "hinweis": FELD_HINWEISE.get(feld),
                })
                abteilung_feld = ABTEILUNG_FELDER.get(feld)
                if abteilung_feld:
                    items.append({
                        "art": "feld", "name": abteilung_feld, "typ": "text", "breite": "kompakt",
                        "wert": data.get(abteilung_feld) or "",
                        "label": f"{feld} - Abteilung",
                        "hinweis": FELD_HINWEISE.get(abteilung_feld),
                    })
                if feld == "Historie":
                    # Kapitel 1 "Grund der Systembewertung" erwartet laut
                    # Template KEINE lange Freitext-Historie, sondern nur
                    # CC-Nr. + Rahmen der Erstellung (Revalidierung/
                    # Qualifizierung/Periodic Review usw.) - webapp-only,
                    # nicht Teil von EXCEL_COLUMNS. Direkt hinter
                    # "Historie" platziert, da beide denselben Anlass
                    # beschreiben, nur in unterschiedlicher Ausfuehrlichkeit
                    # (siehe fill_kapitel1() in template_filler.py).
                    items.append({
                        "art": "feld", "name": "CCNr_Rahmen", "typ": "text", "breite": "voll",
                        "wert": data.get("CCNr_Rahmen") or "",
                        "label": "CC-Nr. / Rahmen der Erstellung (Kapitel 1)",
                        "hinweis": FELD_HINWEISE.get("CCNr_Rahmen"),
                    })
                if feld == "Schnittstellen mit PLS":
                    # Template-Zeile 8 der Beschreibungstabelle
                    # ("Datenfluss / Abbildung:") - webapp-only, nicht
                    # Teil von EXCEL_COLUMNS, da diese Zeile primaer
                    # eine Grafik erwartet und in der Master-Excel bisher
                    # nie eine eigene Spalte hatte. Direkt hinter
                    # "Schnittstellen mit PLS" platziert, da beide
                    # inhaltlich zusammengehoeren (siehe Vorlage unten).
                    items.append({
                        "art": "feld", "name": "DatenflussAbbildung", "typ": "textarea",
                        "breite": "voll", "rows": 3,
                        "wert": data.get("DatenflussAbbildung") or "",
                        "label": "Datenfluss / Abbildung",
                        "hinweis": FELD_HINWEISE.get("DatenflussAbbildung"),
                    })
        if items:
            abschnitte.append((titel, items))
    return abschnitte

def formular_auswerten(form, bestehende_daten):
    """Liest die abgeschickten Formulardaten wieder in ein data-Dict
    (dieselbe Struktur wie EXCEL_COLUMNS)."""
    daten = dict(bestehende_daten)
    for name, info in _kategorien_lookup().items():
        feld_namen = [f for f, _ in info["optionen"]]
        if info["mehrfachauswahl"]:
            ausgewaehlt = set(form.getlist(f"kat__{name}"))
        else:
            einzel = form.get(f"kat__{name}", "")
            ausgewaehlt = {einzel} if einzel else set()
        for f in feld_namen:
            daten[f] = "r" if f in ausgewaehlt else "c"
    for feld in form:
        if feld.startswith("feld__"):
            name = feld[len("feld__"):]
            wert = form.get(feld, "").strip()
            if name == "Dok. -Nr.":
                # Kombiniertes Feld ("QU-OPE-XXXXX / Version 1.0") wieder
                # in "Dok. -Nr." (reine Nummer, so von der Master-Excel-
                # Spalte erwartet) und "Version_Historie" auftrennen.
                daten.update(_dok_nr_version_auftrennen(wert))
            else:
                daten[name] = wert
    return daten

def _heute():
    return date.today().strftime("%d.%m.%Y")

def _draft_titel(daten):
    """MLCS-ID vor den Systemnamen setzen (z.B. "MLCS-1193 - PLS
    Lantus"), damit man einen Draft in der Uebersicht sofort zuordnen
    kann - ausser es gibt (noch) keine MLCS-ID, dann nur der
    Systemname/die MLCS-ID allein."""
    name = (daten.get("AS/BDIS-Name") or "").strip()
    mlcs = (daten.get("MLCSID") or "").strip()
    if mlcs and name:
        return f"{mlcs} - {name}"
    return name or mlcs or "(ohne Titel)"

def _dateiname_vorschlag(data, vorschau=False):
    basis = data.get("Dok. -Nr.") or data.get("MLCSID") or data.get("AS/BDIS-Name") or "Systembewertung"
    basis = re.sub(r"[^A-Za-z0-9_\-]+", "_", basis).strip("_") or "Systembewertung"
    suffix = "_VORSCHAU" if vorschau else ""
    return f"{basis}_Systembewertung{suffix}.docx"

def _vorschau_hinweis_einfuegen(doc):
    """Fuegt ganz oben im Dokument einen deutlich sichtbaren Hinweis
    ein, dass es sich um eine Vorschau handelt - unmissverstaendlich
    von einem finalen (per "Fertigstellen" erzeugten) Dokument zu
    unterscheiden, selbst wenn Dateiname/Suffix beim Weiterleiten mal
    verloren geht."""
    from docx.shared import Pt, RGBColor
    hinweis = doc.paragraphs[0].insert_paragraph_before(
        "⚠️ VORSCHAU - NICHT FINAL - nicht als Abschlussdokument verwenden"
    )
    lauf = hinweis.runs[0]
    lauf.bold = True
    lauf.font.size = Pt(14)
    lauf.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

def _dokument_erzeugen_und_senden(data, vorschau=False):
    """Erzeugt das .docx und schickt es zum Download - schreibt NICHT
    in die Master-Excel (siehe Modul-Kommentar oben, Abschnitt 1.1).

    `vorschau=True` (siehe Route .../vorschau): erzeugt dasselbe
    Dokument, aber mit "_VORSCHAU" im Dateinamen UND einem deutlich
    sichtbaren Hinweis am Dokumentanfang - zum Pruefen des Ergebnisses,
    OHNE dass sich am Draft-Status etwas aendert (bleibt
    "in_bearbeitung", startet also NICHT die 30-Tage-Aufbewahrungsfrist
    fuer "fertig"-Drafts)."""
    data = dict(data)
    data.setdefault("Erkannte_Version", "V11")
    # "Kommt KI zum Einsatz?" (siehe _KATEGORIEN_ZUSATZ) -> Kapitel-2-
    # Checkbox + Begruendung: bei "Nein" wird KINA gesetzt (die genaue
    # Stufe I-VI wird im Editor nicht mehr abgefragt, siehe SKIP_FELDER)
    # und eine angegebene Begruendung an "KI Bewertung" angehaengt. Bei
    # "Ja" bleibt Kapitel 2 (KI1-6) bewusst leer fuer die manuelle
    # Nacharbeit - fill_kapitel9() erkennt "Ja" trotzdem direkt anhand
    # von KI_Einsatz_Ja fuer die Kapitel-9.1-Checkbox.
    if data.get("KI_Einsatz_Nein") == "r":
        data["KINA"] = "r"
        for _ki_feld in ("KI1", "KI2", "KI3", "KI4", "KI5", "KI6"):
            data[_ki_feld] = "c"
        begruendung = (data.get("KI_Einsatz_Begruendung") or "").strip()
        if begruendung:
            bestehend = data.get("KI Bewertung") or ""
            data["KI Bewertung"] = f"{bestehend}\n{begruendung}" if bestehend else begruendung
    # "Globale CS-Klasse (Kapitel 3)" (siehe _KATEGORIEN_ZUSATZ): eine
    # Klasse (1a/1b/2/3) impliziert laut Formular immer "Globales CS"
    # (Kapitel 1) - wird deshalb hier automatisch mitgesetzt, ohne dass
    # dafuer im Editor extra eine eigene Checkbox angekreuzt werden
    # muss. "KLASS_Global_NA" ist rein webapp-intern (keine Excel-
    # Spalte, nicht in EXCEL_COLUMNS) und wird hier bewusst nicht
    # weiterverarbeitet - "N/A" bedeutet lediglich, dass keine der 4
    # Klassen zutrifft (siehe fill_kapitel3: N/A ergibt sich dort
    # ohnehin automatisch daraus, dass KLASS_Global nicht "r" ist).
    if any(data.get(f) == "r" for f in ("KLASS_Global_1a", "KLASS_Global_1b", "KLASS_Global_2", "KLASS_Global_3")):
        data["KLASS_Global"] = "r"
    tmp_dir = tempfile.mkdtemp(prefix="sysbew_")
    dateiname = _dateiname_vorschlag(data, vorschau=vorschau)
    ausgabe_pfad = os.path.join(tmp_dir, dateiname)
    doc = template_filler.fill_template(data)
    if vorschau:
        _vorschau_hinweis_einfuegen(doc)
    doc.save(ausgabe_pfad)
    return send_file(ausgabe_pfad, as_attachment=True, download_name=dateiname)

# ============================================================
# Benutzer-Kennung (nur fuer Sperren/Protokoll - keine echte
# Authentifizierung, dafuer gibt es hier keinen Bedarf: die App
# laeuft lokal, jede:r startet ihre/seine eigene Instanz)
# ============================================================
@app.before_request
def _benutzer_pruefen():
    if request.endpoint in ("name_setzen", "static", "lebenszeichen", "schliessen_signal"):
        return None
    if not session.get("user"):
        return redirect(url_for("name_setzen", weiter=request.path))
    return None

@app.route("/name", methods=["GET", "POST"])
def name_setzen():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        if name:
            session["user"] = name
            session.permanent = True
            return redirect(request.args.get("weiter") or url_for("index"))
        flash("Bitte einen Namen eingeben.")
    return render_template("name.html")

# ============================================================
# Startseite: Datenbank filtern (Weg 1 + Einstieg Weg 2)
# ============================================================
@app.route("/")
def start():
    """Startseite: fragt zuerst, WAS gemacht werden soll, bevor die
    eigentlichen Seiten (Datenbank-Suche bzw. Draft-Übersicht) kommen -
    vorher landete man direkt auf der Datenbank-Suche."""
    return render_template("start.html", anzahl_offene_drafts=draft_store.anzahl_offene_drafts())

@app.route("/datenbank")
def index():
    suchtext = request.args.get("q", "")
    fehler = None
    treffer = []
    try:
        _, rows = db_reader.read_master_rows()
        treffer = db_reader.filter_rows(rows, suchtext)
    except Exception as e:
        fehler = str(e)
    return render_template(
        "index.html", rows=treffer[:200], gesamt=len(treffer), q=suchtext,
        fehler=fehler, spalten=db_reader.UEBERSICHT_SPALTEN,
    )

@app.route("/db/<int:zeile>/uebernehmen")
def db_uebernehmen(zeile):
    """Weg 1: Daten direkt, ohne Editor, in eine neue Systembewertung
    uebertragen."""
    _, rows = db_reader.read_master_rows()
    row = db_reader.get_row_by_index(rows, zeile)
    if not row:
        flash("Zeile nicht gefunden.")
        return redirect(url_for("index"))
    data = _neues_dokument_aus_db_zeile(row)
    return _dokument_erzeugen_und_senden(data)

@app.route("/db/<int:zeile>/bearbeiten")
def db_bearbeiten(zeile):
    """Weg 2, Einstieg aus der Datenbank: Editor mit den Daten der
    Zeile vorausgefuellt oeffnen."""
    _, rows = db_reader.read_master_rows()
    row = db_reader.get_row_by_index(rows, zeile)
    if not row:
        flash("Zeile nicht gefunden.")
        return redirect(url_for("index"))
    data = _neues_dokument_aus_db_zeile(row)
    draft_id = draft_store.create_draft(
        data, session["user"], titel=_draft_titel(data), quelle_zeile=zeile,
    )
    return redirect(url_for("editor", draft_id=draft_id))

@app.route("/masterform", methods=["GET", "POST"])
def masterform():
    """Dritte Startquelle fuer eine neue Systembewertung: eine von der
    anderen Seite ("Fill-a-Masterform") als Excel bereitgestellte,
    bereits dekodierte Zeile importieren (siehe masterform_import.py).

    Bewusst kein dauerhaft konfigurierter Netzwerkpfad wie bei der
    Master-Excel - die Datei wird bei Bedarf hochgeladen und nur fuer
    diese eine Anfrage gelesen, nirgends gespeichert. Die Zeilen
    werden deshalb pro Treffer als verstecktes Formularfeld (JSON) in
    die Ergebnisliste eingebettet, damit "Bearbeiten" die Datei nicht
    erneut braucht (siehe masterform_bearbeiten())."""
    zeilen = None
    fehler = None
    suchtext = ""
    if request.method == "POST":
        datei = request.files.get("datei")
        suchtext = request.form.get("q", "")
        if not datei or not datei.filename:
            flash("Bitte zuerst eine Datei auswaehlen.")
        else:
            try:
                _, rows = masterform_import.lese_masterform_export(datei)
                zeilen = masterform_import.filter_rows(rows, suchtext)
            except Exception as e:
                fehler = str(e)
    return render_template(
        "masterform.html", zeilen=zeilen, fehler=fehler, q=suchtext,
    )

@app.route("/masterform/bearbeiten", methods=["POST"])
def masterform_bearbeiten():
    """Legt aus einer im Formular mitgeschickten Fill-a-Masterform-
    Zeile (siehe masterform.html, verstecktes Feld "zeile_json") einen
    neuen Draft an - analog zu db_bearbeiten(), nur mit
    masterform_import.zeile_zu_sysbew_daten() statt einer echten
    ML-Zeile als Quelle. Alle nicht zweifelsfrei uebernommenen Werte
    werden vorher als Hinweis angezeigt (siehe dortiger Modul-
    Kommentar) - der Draft ist trotzdem sofort im Editor pruefbar."""
    roh = request.form.get("zeile_json", "")
    try:
        row = json.loads(roh)
    except (json.JSONDecodeError, TypeError):
        flash("Zeile konnte nicht gelesen werden - bitte Datei erneut hochladen.")
        return redirect(url_for("masterform"))
    daten, hinweise = masterform_import.zeile_zu_sysbew_daten(row)
    for hinweis in hinweise:
        flash(f"⚠️ {hinweis}")
    data = _neues_dokument_aus_db_zeile(daten)
    draft_id = draft_store.create_draft(
        data, session["user"], titel=_draft_titel(data),
        quelle_zeile=f"masterform:{row.get('mlcs_id') or row.get('_zeile')}",
    )
    return redirect(url_for("editor", draft_id=draft_id))

@app.route("/editor/neu")
def editor_neu():
    """Weg 2, Einstieg ohne Datenbank: leerer Editor."""
    data = {"Version_Historie": "1.0", "Datum": _heute()}
    draft_id = draft_store.create_draft(data, session["user"], titel="Neue Systembewertung")
    return redirect(url_for("editor", draft_id=draft_id))

# ============================================================
# Editor (Weg 2)
# ============================================================
@app.route("/editor/<draft_id>")
def editor(draft_id):
    draft = draft_store.load_draft(draft_id)
    if not draft:
        flash("Draft nicht gefunden (evtl. geloescht).")
        return redirect(url_for("drafts_uebersicht"))
    if not draft_store.acquire_lock(draft_id, session["user"]):
        return render_template("gesperrt.html", draft=draft, sperre=draft_store.lock_status(draft_id))
    formular = baue_formular(draft["daten"])
    return render_template(
        "editor.html", draft=draft, formular=formular,
        textbaustein_vorlagen=TEXTBAUSTEIN_VORLAGEN,
    )

@app.route("/editor/<draft_id>/speichern", methods=["POST"])
def editor_speichern(draft_id):
    draft = draft_store.load_draft(draft_id)
    if not draft:
        flash("Draft nicht gefunden.")
        return redirect(url_for("drafts_uebersicht"))
    daten = formular_auswerten(request.form, draft["daten"])
    draft_store.save_draft(draft_id, daten, session["user"], titel=_draft_titel(daten))
    draft_store.acquire_lock(draft_id, session["user"])
    flash("Zwischengespeichert.")
    return redirect(url_for("editor", draft_id=draft_id))

@app.route("/editor/<draft_id>/vorschau", methods=["POST"])
def editor_vorschau(draft_id):
    """Erzeugt ein Vorschau-Dokument (siehe _dokument_erzeugen_und_senden)
    aus dem aktuellen Formularstand - speichert den Draft dabei wie
    "Zwischenspeichern" ab (nichts geht verloren), aendert aber NICHT
    den Status auf "fertig": der Draft bleibt "in_bearbeitung", die
    30-Tage-Aufbewahrungsfrist fuer "fertig"-Drafts startet also nicht."""
    draft = draft_store.load_draft(draft_id)
    if not draft:
        flash("Draft nicht gefunden.")
        return redirect(url_for("drafts_uebersicht"))
    daten = formular_auswerten(request.form, draft["daten"])
    draft_store.save_draft(draft_id, daten, session["user"], titel=_draft_titel(daten))
    draft_store.acquire_lock(draft_id, session["user"])
    return _dokument_erzeugen_und_senden(daten, vorschau=True)

@app.route("/editor/<draft_id>/heartbeat", methods=["POST"])
def editor_heartbeat(draft_id):
    ok = draft_store.heartbeat(draft_id, session.get("user", ""))
    return {"ok": ok}

@app.route("/editor/<draft_id>/fertigstellen", methods=["POST"])
def editor_fertigstellen(draft_id):
    draft = draft_store.load_draft(draft_id)
    if not draft:
        flash("Draft nicht gefunden.")
        return redirect(url_for("drafts_uebersicht"))
    daten = formular_auswerten(request.form, draft["daten"])
    daten["Systemtyp_CE"] = common.berechne_systemtyp_ce(daten)
    draft_store.save_draft(
        draft_id, daten, session["user"], status="fertig", titel=_draft_titel(daten),
    )
    draft_store.release_lock(draft_id, session["user"])
    return _dokument_erzeugen_und_senden(daten)

@app.route("/editor/<draft_id>/freigeben", methods=["POST"])
def editor_freigeben(draft_id):
    draft_store.release_lock(draft_id, session["user"])
    flash("Sperre freigegeben.")
    return redirect(url_for("drafts_uebersicht"))

# ============================================================
# Draft-Uebersicht (Punkt 3: Zwischenstaende, mehrbenutzerfaehig)
# ============================================================
@app.route("/drafts")
def drafts_uebersicht():
    return render_template("drafts.html", drafts=draft_store.list_drafts(), user=session.get("user"))

@app.route("/drafts/<draft_id>/loeschen", methods=["POST"])
def draft_loeschen(draft_id):
    draft_store.delete_draft(draft_id)
    flash("Draft geloescht.")
    return redirect(url_for("drafts_uebersicht"))

@app.route("/drafts/<draft_id>/herunterladen")
def draft_herunterladen(draft_id):
    """Fuer bereits 'fertig' gestellte Drafts: Dokument erneut
    erzeugen und herunterladen, ohne den Editor erneut zu oeffnen."""
    draft = draft_store.load_draft(draft_id)
    if not draft:
        flash("Draft nicht gefunden.")
        return redirect(url_for("drafts_uebersicht"))
    return _dokument_erzeugen_und_senden(draft["daten"])

if __name__ == "__main__":
    Timer(1.0, lambda: webbrowser.open(f"http://127.0.0.1:{PORT}/")).start()
    app.run(host="127.0.0.1", port=PORT, debug=False)
