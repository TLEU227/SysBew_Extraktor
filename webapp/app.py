# ============================================================
# app.py
# Systembewertung-Editor - Web-Oberflaeche - 1.0
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

PORT = 5151

# Felder, die nicht als Eingabe im Formular auftauchen, weil sie
# automatisch berechnet werden (Systemtyp_CE, Erkannte_Version) oder
# weil das Filler-Modul sie (noch) nicht ins Dokument schreiben kann
# (Testtiefe/Testtiefe-Matrix - siehe README.md "Bekannte
# Einschraenkungen": haengt an Kapitel 8, das nicht automatisch
# befuellt wird).
SKIP_FELDER = {"Erkannte_Version", "Python ja/nein", "Systemtyp_CE",
               "Testtiefe", "Testtiefe-Matrix"}

_TEXTAREA_FELDER = (
    set(template_filler._BESCHREIBUNG_ZEILEN.values())
    | {"Besonderheiten", "GxP_Produktqualitaet", "GxP_Patientensicherheit",
       "GxP_Datenintegritaet", "Kurzbeschreibung", "Historie"}
    | set(common.ROLLEN_SPALTEN)
)

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
    return lookup

def baue_formular(data):
    """Baut die Formularstruktur fuer editor.html aus den
    VORSCHAU_ABSCHNITTE (dieselbe Gliederung wie die
    Konsolen-Vorschau von word_parser_main.py - Deckblatt zuerst,
    dann Kapitel 1/2 usw.)."""
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
                })
            else:
                items.append({
                    "art": "feld", "name": feld,
                    "typ": "textarea" if feld in _TEXTAREA_FELDER else "text",
                    "wert": data.get(feld) or "",
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

def _dateiname_vorschlag(data):
    basis = data.get("Dok. -Nr.") or data.get("MLCSID") or data.get("AS/BDIS-Name") or "Systembewertung"
    basis = re.sub(r"[^A-Za-z0-9_\-]+", "_", basis).strip("_") or "Systembewertung"
    return f"{basis}_Systembewertung.docx"

def _dokument_erzeugen_und_senden(data, excel_eintragen=False):
    data = dict(data)
    data.setdefault("Erkannte_Version", "V11")
    tmp_dir = tempfile.mkdtemp(prefix="sysbew_")
    dateiname = _dateiname_vorschlag(data)
    ausgabe_pfad = os.path.join(tmp_dir, dateiname)
    template_filler.fill_template(data, output_path=ausgabe_pfad)

    if excel_eintragen:
        try:
            data_mit_flag = dict(data)
            data_mit_flag["Python ja/nein"] = "ja"
            ziel_zeile = common.write_to_master_excel(data_mit_flag, ausgabe_pfad)
            if ziel_zeile:
                flash(f"Auch in Master-Excel eingetragen (Zeile {ziel_zeile}).")
            else:
                flash("Dokument erzeugt, Eintrag in Master-Excel ist aber fehlgeschlagen (siehe Konsole).")
        except Exception as e:
            flash(f"Dokument erzeugt, aber Eintrag in Master-Excel fehlgeschlagen: {e}")

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
    draft_id = draft_store.create_draft(data, session["user"], quelle_zeile=zeile)
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
    return render_template("editor.html", draft=draft, formular=formular)

@app.route("/editor/<draft_id>/speichern", methods=["POST"])
def editor_speichern(draft_id):
    draft = draft_store.load_draft(draft_id)
    if not draft:
        flash("Draft nicht gefunden.")
        return redirect(url_for("drafts_uebersicht"))
    daten = formular_auswerten(request.form, draft["daten"])
    draft_store.save_draft(
        draft_id, daten, session["user"],
        titel=daten.get("AS/BDIS-Name") or daten.get("MLCSID"),
    )
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
    draft_store.save_draft(draft_id, daten, session["user"], status="fertig")
    draft_store.release_lock(draft_id, session["user"])
    excel_eintragen = request.form.get("excel_eintragen") == "on"
    return _dokument_erzeugen_und_senden(daten, excel_eintragen=excel_eintragen)

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
