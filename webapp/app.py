# ============================================================
# app.py
# Systembewertung-Editor - Web-Oberflaeche - 1.2
#
# Erzeugt NEUE Systembewertungen (V11) aus Daten der Master-Excel
# ("Datenbank") oder von Grund auf, mit Zwischenspeicherung als
# Draft und Mehrbenutzer-faehiger Sperrung ueber draft_store.py.
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

import os
import re
import sys
import tempfile
import webbrowser
from datetime import date
from threading import Timer

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "lib"))

from flask import Flask, flash, redirect, render_template, request, send_file, session, url_for

import db_reader
import draft_store
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
APP_VERSION = "1.2"

@app.context_processor
def _globale_template_variablen():
    return {"app_version": APP_VERSION, "optionen_hinweise": OPTIONEN_HINWEISE}

PORT = 5151

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
#   - API/BE/Raum/SAP/DokNummerVorQualiPSO/Bearbeiter/PLSTA: reine
#     Master-Excel-Spalten ohne jede Rolle beim Erzeugen des Dokuments
#     (per Code-Analyse geprueft: template_filler.py liest sie nie;
#     API ist zudem nur eine aus "Betrieb" abgeleitete Teilzeichenkette
#     fuers Filtern in der Excel, "Betrieb" selbst wird geschrieben).
#   - "Version": ist beim LESEN eines bestehenden Dokuments die dort
#     bereits vorhandene, hoechste Version - fuer ein NEUES Dokument
#     irrelevant (das startet immer bei "Version_Historie" = 1.0).
#   - "SW-Version / Typ:"/"SW-Name:": beim Lesen per Regex aus dem
#     Besonderheiten-Text abgeleitet, haben aber keine eigene Zelle im
#     Template - bei Bedarf einfach direkt im Feld "Besonderheiten"
#     mit angeben.
#   - "BemerkungX": inhaltlich redundant, da Prozessbeschreibung/Daten/
#     Audit Trail/Parameter bereits weiter oben unter ihrem eigenen
#     Namen abgefragt werden. Der Transfer dieser Werte in die
#     "BemerkungX"-Spalten der Master-Excel passiert automatisch beim
#     Einlesen des fertigen Dokuments (word_parser_v8/10/11.py machen
#     das schon seit jeher so) - hier also keine doppelte Abfrage.
SKIP_FELDER = {"Erkannte_Version", "Python ja/nein", "Systemtyp_CE",
               "Testtiefe", "Testtiefe-Matrix", "Phenix",
               "DI EE-Anforderungen", "Steuerung erfolgt über?",
               "API", "BE", "Raum", "SAP", "DokNummerVorQualiPSO",
               "Bearbeiter", "PLSTA", "Version",
               "SW-Version / Typ:", "SW-Name:",
               "Bemerkung1", "Bemerkung2", "Bemerkung3", "Bemerkung4"}

_TEXTAREA_FELDER = (
    set(template_filler._BESCHREIBUNG_ZEILEN.values())
    | {"Besonderheiten", "GxP_Produktqualitaet", "GxP_Patientensicherheit",
       "GxP_Datenintegritaet", "Kurzbeschreibung", "Historie"}
    | set(common.ROLLEN_SPALTEN)
)

# Abteilungs-Felder (webapp-only, nicht Teil von EXCEL_COLUMNS): jede
# Rollenzeile im Deckblatt hat einen "(Site/Unit)"-Platzhalter (ausser
# CSQ - dort ist die Abteilung im Template fest vorgegeben). Werden im
# Formular direkt neben dem jeweiligen Namensfeld angezeigt (siehe
# baue_formular()) und von template_filler.fill_deckblatt_rollen() in
# das Label der jeweiligen Zeile eingesetzt.
ABTEILUNG_FELDER = {
    "Ersteller": "Ersteller_Abteilung",
    "SI/PL":     "SI_PL_Abteilung",
    "TSO":       "TSO_Abteilung",
    "BSO":       "BSO_Abteilung",
    "BQR":       "BQR_Abteilung",
}

# Freundliche Beschriftung fuer die 4 generischen "BemerkungX"-Spalten
# der Master-Excel - haben laut Fachbereich eine feste Bedeutung
# innerhalb des Kapitels "Informationen und Bemerkungen".
BEMERKUNG_LABELS = {
    "Bemerkung1": "Bemerkung 1 (Prozessbeschreibung)",
    "Bemerkung2": "Bemerkung 2 (Daten)",
    "Bemerkung3": "Bemerkung 3 (Audit Trail)",
    "Bemerkung4": "Bemerkung 4 (Parameter)",
}

# Hinweistexte, die im Editor unter dem jeweiligen Feld/der jeweiligen
# Kategorie angezeigt werden - wo moeglich woertlich aus dem Template
# uebernommen. Fuer PLSTA und "SW-Version / Typ:" (Bedeutung von "VV")
# gibt es aktuell KEINE gesicherte Definition in Template oder Skripten
# - das wird hier bewusst ehrlich ausgewiesen statt geraten.
FELD_HINWEISE = {
    "Dok. -Nr.": (
        "Dokumentennummer der NEUEN Systembewertung (z. B. QU-OPE-XXXXX) - nicht die des "
        "Vorgänger-Dokuments. ℹ️ Diese Systembewertung wird auf Basis der im System hinterlegten "
        "Vorlage V11 erstellt (\"Erkannte Version2\" in der Master-Excel) - aktuell die einzige "
        "hinterlegte Version, wird automatisch gesetzt."
    ),
    "Hyperlink": "Link auf das Dokument in QualiPSO (sobald dort abgelegt).",
    "MLCSID": "System-Identifier/CS-Inventarnummer gemäß QU-SOP-0052370. Keine MLCS erforderlich für S0 und Equipment ohne CS.",
    "UeberlagerteMLCS": "Übergeordnetes System: Systemname, MLCS-ID und ggf. Doc-ID der zugehörigen Systembewertung.",
    "Schnittstelle": "Schnittstelle zu anderen/übergeordneten Systemen: Systemname, MLCS-ID und ggf. Doc-ID der Systembewertung.",
    "Anlage": "Anlagen-IDs/Equipment-Nr./QC-ID.",
    "PLSTA": "⚠️ Bedeutung nicht dokumentiert (in keinem der Vorgänger-Skripte befüllt worden) - bitte klären oder leer lassen.",
    "AS/BDIS-Name": "Bei Equipment z. B. Laborwaage; bei System: Software-/Systemname.",
    "Kurzbeschreibung": "Wozu wird das System/Equipment/die Anlage eingesetzt?",
    "Betrieb": "Einsatzort: Site/Organisationseinheit.",
    "Gebaeude": "Gebäude-Kürzel, falls vorhanden.",
    "Hersteller": "Name Hersteller/Lieferant. Bei zugelassenen Lieferanten die QualiPSO-ID im Feld „Lieferantennummer“ ergänzen.",
    "SW-Hersteller": "Hersteller/Lieferant der Software (falls abweichend vom Equipment-Hersteller).",
    "SW-Version / Typ:": "⚠️ Bedeutung von „VV“ und den weiteren dort üblichen Optionen ist uns nicht dokumentiert - bitte bei Bedarf klären. Frei ausfüllen.",
    "DokNummerVorQualiPSO": "Dokumentennummer des Vorgänger-Dokuments vor Überführung nach QualiPSO, falls vorhanden.",
    "Lieferantennummer": "QualiPSO-/QTP-Customer-ID des Lieferanten, falls vorhanden.",
    "Historie": "Grund der Erstellung/Änderung - siehe Textbaustein-Vorschläge unten.",
    "Besonderheiten": "Bei Gerätekategorien A, B und C bitte die Subkategorisierung (z. B. B1, C2) nach QU-SOP-0021736 begründen - siehe Textbaustein-Vorschläge unten.",
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
}

KATEGORIE_HINWEISE = {
    "Klassifizierung": "Diese Auswahl wird zusätzlich in Kapitel 3 des Dokuments übernommen (Systemeinstufung Globales CS).",
    "Periodic Review": "Zyklische Requalifizierung gemäß QU-SOP-0072260 (kein Kreuz bei CIS und Spreadsheets).",
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
}

# Vorschlags-Textbausteine fuer Historie ("Grund der Erstellung") und
# Besonderheiten - werden per Knopf im Editor in das jeweilige Feld
# eingefuegt (siehe editor.html), koennen danach frei angepasst werden.
HISTORIE_VORLAGEN = [
    {
        "label": "Vorlage 1: Anpassung",
        "text": (
            "Anpassung mit inhaltlicher Überführung der gültigen Systemeinstufung "
            "AS/BDIS für dieses System (Dok.-Nr.: FRA-BERI-G-011453 / Version 2) in "
            "die Systembewertung gemäß QU-MT-0001344 Version 11. Überführung in das "
            "Dokumenten-Managementsystem QualiPSO.\n\n"
            "Dabei erfolgte keine Änderung der gültigen Bewertungen. Ergänzende "
            "Angaben wurden im Kapitel „Informationen und Bemerkungen“ notiert."
        ),
    },
    {
        "label": "Vorlage 2: Neuerstellung (Papier → QualiPSO)",
        "text": (
            "Neuerstellung mit inhaltlicher Überführung der gültigen Systemeinstufung "
            "AS/BDIS für das PLS Lantus (Dok.-Nr.: L_QUA_1143 / Version 4 / "
            "MLCS-ID: 1193) in die Systembewertung gemäß FRA-FORM-001283 Version 4.0.\n"
            "Dabei erfolgte keine Änderung der gültigen Bewertungen. Ergänzende "
            "Angaben wurden im Kapitel „Informationen und Bemerkungen“ notiert.\n"
            "Anpassungen auf Basis der FRA-QU-MT-0001344 (FRA-FORM-001283) Version "
            "11.0 anlässlich der Qualifizierung für ..."
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
]

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
        lookup[name] = {"optionen": optionen, "mehrfachauswahl": False}
    lookup.update(_KATEGORIEN_ZUSATZ)
    return lookup

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
                    # "— keine Angabe —" waere hier redundant, wenn die
                    # Kategorie ohnehin schon eine eigene N/A-Option hat.
                    "hat_na": any(label.strip().upper() == "N/A" for _, label in info["optionen"]),
                })
                if feld == "Periodic Review":
                    items.append({
                        "art": "feld", "name": "PR_Andere_Text", "typ": "text", "breite": "kompakt",
                        "wert": data.get("PR_Andere_Text") or "",
                        "label": "Freie Angabe - Details",
                        "hinweis": 'Nur ausfüllen, wenn oben „andere/freie Angabe“ ausgewählt wurde.',
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
                    "rows": 8 if feld in ("Historie", "Besonderheiten") else 3,
                    "wert": data.get(feld) or "",
                    "label": BEMERKUNG_LABELS.get(feld),
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
            daten[feld[len("feld__"):]] = form.get(feld, "").strip()
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

def _dateiname_vorschlag(data):
    basis = data.get("Dok. -Nr.") or data.get("MLCSID") or data.get("AS/BDIS-Name") or "Systembewertung"
    basis = re.sub(r"[^A-Za-z0-9_\-]+", "_", basis).strip("_") or "Systembewertung"
    return f"{basis}_Systembewertung.docx"

def _dokument_erzeugen_und_senden(data):
    """Erzeugt das .docx und schickt es zum Download - schreibt NICHT
    in die Master-Excel (siehe Modul-Kommentar oben, Abschnitt 1.1)."""
    data = dict(data)
    data.setdefault("Erkannte_Version", "V11")
    tmp_dir = tempfile.mkdtemp(prefix="sysbew_")
    dateiname = _dateiname_vorschlag(data)
    ausgabe_pfad = os.path.join(tmp_dir, dateiname)
    template_filler.fill_template(data, output_path=ausgabe_pfad)
    return send_file(ausgabe_pfad, as_attachment=True, download_name=dateiname)

# ============================================================
# Benutzer-Kennung (nur fuer Sperren/Protokoll - keine echte
# Authentifizierung, dafuer gibt es hier keinen Bedarf: die App
# laeuft lokal, jede:r startet ihre/seine eigene Instanz)
# ============================================================
@app.before_request
def _benutzer_pruefen():
    if request.endpoint in ("name_setzen", "static"):
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
    data = {k: v for k, v in row.items() if k != "_zeile"}
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
    data = {k: v for k, v in row.items() if k != "_zeile"}
    draft_id = draft_store.create_draft(
        data, session["user"], titel=_draft_titel(data), quelle_zeile=zeile,
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
        historie_vorlagen=HISTORIE_VORLAGEN,
        besonderheiten_vorlagen=BESONDERHEITEN_VORLAGEN,
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
