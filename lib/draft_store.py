# ============================================================
# draft_store.py
# Zwischenspeicherung von Bearbeitungsstaenden (Drafts) - 1.1
#
# Drafts liegen als einzelne JSON-Dateien in einem gemeinsamen Ordner
# neben der Master-Excel (also im selben, bereits per OneDrive/
# SharePoint synchronisierten Netzlaufwerk-Ordner - kein zusaetzlicher
# Server/Freigabe noetig). Dadurch kann JEDE:R Nutzer:in JEDEN Draft
# wieder oeffnen, auch an einem anderen PC - es gibt keine feste
# Anzahl "Slots"; es werden so viele Draft-Dateien angelegt, wie
# gerade gebraucht werden.
#
# Gleichzeitige Bearbeitung mehrerer Nutzer:innen wird ueber eine
# Sperr-Datei (<draft_id>.lock) je Draft geregelt, NICHT ueber einen
# festen Server-Prozess:
#   - Beim Oeffnen zum Bearbeiten wird ein Lock mit Benutzername +
#     Zeitstempel geschrieben.
#   - Waehrend der Bearbeitung schickt der Browser periodisch einen
#     Heartbeat (siehe webapp/app.py), der den Zeitstempel erneuert.
#   - Ohne Heartbeat laeuft der Lock nach LOCK_TIMEOUT_SEKUNDEN
#     automatisch ab (z.B. Browser-Tab geschlossen, PC abgestuerzt) -
#     der Draft wird dann automatisch wieder frei, OHNE dass jemand
#     ihn manuell entsperren muss.
#   - Ist ein Draft durch jemand anderen gesperrt, bekommt die
#     Startseite/Draft-Uebersicht das klar angezeigt (wer, seit wann) -
#     es gibt also NIE eine stille Blockade ohne Erklaerung.
# ============================================================

import glob
import json
import os
import tempfile
import uuid
from datetime import datetime, timedelta

from sysbew_common import MASTER_EXCEL_PFAD

DRAFTS_ORDNER = os.path.join(os.path.dirname(MASTER_EXCEL_PFAD), "Drafts")

LOCK_TIMEOUT_SEKUNDEN = 15 * 60  # 15 Minuten ohne Heartbeat -> Lock verfaellt

# Ein "fertig" gestellter Draft wird noch AUFBEWAHRUNG_FERTIG_TAGE lang
# aufbewahrt (z.B. um das Dokument ueber "Erneut herunterladen" noch
# einmal zu bekommen), danach automatisch geloescht - die eigentlichen
# Daten leben ab dem Fertigstellen ohnehin im erzeugten SB-Dokument,
# der Draft ist ab da nur noch eine Sicherheitskopie fuer den
# Uebergang, keine dauerhafte Ablage. Betrifft NUR Drafts mit Status
# "fertig" - Drafts "in_bearbeitung" werden nie automatisch geloescht,
# egal wie alt (koennten ja noch aktiv bearbeitet werden).
AUFBEWAHRUNG_FERTIG_TAGE = 30

def _sicherstellen_ordner():
    os.makedirs(DRAFTS_ORDNER, exist_ok=True)

def _draft_pfad(draft_id):
    return os.path.join(DRAFTS_ORDNER, f"{draft_id}.json")

def _lock_pfad(draft_id):
    return os.path.join(DRAFTS_ORDNER, f"{draft_id}.lock")

def _jetzt():
    return datetime.now().isoformat(timespec="seconds")

def _atomar_schreiben(pfad, inhalt_dict):
    """Schreibt JSON atomar (erst temp-Datei, dann os.replace), damit
    ein gleichzeitiger Lesezugriff nie eine halb geschriebene Datei
    sieht - wichtig bei mehreren Nutzer:innen auf demselben
    Netzlaufwerk-Ordner."""
    _sicherstellen_ordner()
    fd, temp_pfad = tempfile.mkstemp(dir=DRAFTS_ORDNER, suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(inhalt_dict, f, ensure_ascii=False, indent=2)
        os.replace(temp_pfad, pfad)
    except Exception:
        if os.path.exists(temp_pfad):
            os.remove(temp_pfad)
        raise

# ============================================================
# Draft CRUD
# ============================================================
def neue_draft_id():
    return uuid.uuid4().hex[:12]

def create_draft(daten, user, titel=None, quelle_zeile=None):
    """Legt einen neuen Draft an und gibt seine ID zurueck."""
    draft_id = neue_draft_id()
    jetzt = _jetzt()
    inhalt = {
        "id": draft_id,
        "titel": titel or daten.get("AS/BDIS-Name") or daten.get("MLCSID") or "(ohne Titel)",
        "status": "in_bearbeitung",
        "erstellt_von": user,
        "erstellt_am": jetzt,
        "geaendert_von": user,
        "geaendert_am": jetzt,
        "quelle_zeile": quelle_zeile,
        "daten": daten,
    }
    _atomar_schreiben(_draft_pfad(draft_id), inhalt)
    return draft_id

def load_draft(draft_id):
    pfad = _draft_pfad(draft_id)
    if not os.path.exists(pfad):
        return None
    with open(pfad, encoding="utf-8") as f:
        return json.load(f)

def save_draft(draft_id, daten, user, titel=None, status=None):
    """Aktualisiert einen bestehenden Draft. Legt ihn an, falls er
    (z.B. durch Race Condition) nicht mehr existiert."""
    inhalt = load_draft(draft_id) or {
        "id": draft_id,
        "status": "in_bearbeitung",
        "erstellt_von": user,
        "erstellt_am": _jetzt(),
        "quelle_zeile": None,
    }
    inhalt["daten"] = daten
    inhalt["geaendert_von"] = user
    inhalt["geaendert_am"] = _jetzt()
    if titel:
        inhalt["titel"] = titel
    if status:
        inhalt["status"] = status
        if status == "fertig":
            # Startet/erneuert die Aufbewahrungsfrist (siehe
            # AUFBEWAHRUNG_FERTIG_TAGE) - wird bei jedem (erneuten)
            # Fertigstellen aktualisiert, nicht nur beim ersten Mal.
            inhalt["fertiggestellt_am"] = _jetzt()
    _atomar_schreiben(_draft_pfad(draft_id), inhalt)
    return inhalt

def delete_draft(draft_id):
    for pfad in (_draft_pfad(draft_id), _lock_pfad(draft_id)):
        if os.path.exists(pfad):
            os.remove(pfad)

def _alte_fertige_drafts_aufraeumen():
    """Loescht Drafts, die seit mehr als AUFBEWAHRUNG_FERTIG_TAGE Tagen
    den Status "fertig" haben. Wird bei jedem list_drafts()-Aufruf
    (also bei jedem Besuch der Draft-Uebersicht/Startseite) beilaeufig
    ausgefuehrt - kein separater Hintergrund-Prozess/Scheduler noetig,
    da die App ohnehin nur laeuft, waehrend jemand sie im Browser
    offen hat (siehe Auto-Beenden in webapp/app.py)."""
    grenze = datetime.now() - timedelta(days=AUFBEWAHRUNG_FERTIG_TAGE)
    for pfad in glob.glob(os.path.join(DRAFTS_ORDNER, "*.json")):
        draft_id = os.path.splitext(os.path.basename(pfad))[0]
        try:
            with open(pfad, encoding="utf-8") as f:
                inhalt = json.load(f)
        except (json.JSONDecodeError, OSError):
            continue
        if inhalt.get("status") != "fertig":
            continue
        fertig_am = inhalt.get("fertiggestellt_am")
        if not fertig_am:
            # Aeltere, bereits vor diesem Feature fertiggestellte
            # Drafts haben noch keinen Zeitstempel - werden nicht
            # geloescht, sondern beim naechsten (erneuten)
            # Fertigstellen erstmals einen bekommen. Kein rueckwirkendes
            # Raten des Zeitpunkts.
            continue
        try:
            zeitpunkt = datetime.fromisoformat(fertig_am)
        except ValueError:
            continue
        if zeitpunkt < grenze:
            delete_draft(draft_id)

def list_drafts():
    """Alle Drafts mit aktuellem Sperrstatus, neueste Aenderung
    zuerst. Raeumt beilaeufig lange fertige Drafts auf (siehe
    _alte_fertige_drafts_aufraeumen)."""
    _sicherstellen_ordner()
    _alte_fertige_drafts_aufraeumen()
    ergebnis = []
    for pfad in glob.glob(os.path.join(DRAFTS_ORDNER, "*.json")):
        draft_id = os.path.splitext(os.path.basename(pfad))[0]
        try:
            with open(pfad, encoding="utf-8") as f:
                inhalt = json.load(f)
        except (json.JSONDecodeError, OSError):
            continue
        eintrag = {k: v for k, v in inhalt.items() if k != "daten"}
        eintrag["sperre"] = lock_status(draft_id)
        ergebnis.append(eintrag)
    ergebnis.sort(key=lambda e: e.get("geaendert_am", ""), reverse=True)
    return ergebnis

def anzahl_offene_drafts():
    """Anzahl der Drafts, die noch "in_bearbeitung" sind (nicht
    "fertig") - fuer einen kurzen Hinweis auf der Startseite, ohne dass
    diese jeden Draft komplett auflisten muss."""
    return sum(1 for d in list_drafts() if d.get("status") != "fertig")

# ============================================================
# Sperren (Locking)
# ============================================================
def lock_status(draft_id):
    """None, wenn frei (kein Lock oder abgelaufen), sonst
    {"user":..., "seit":..., "heartbeat":...}."""
    pfad = _lock_pfad(draft_id)
    if not os.path.exists(pfad):
        return None
    try:
        with open(pfad, encoding="utf-8") as f:
            lock = json.load(f)
    except (json.JSONDecodeError, OSError):
        return None
    heartbeat = datetime.fromisoformat(lock["heartbeat"])
    if datetime.now() - heartbeat > timedelta(seconds=LOCK_TIMEOUT_SEKUNDEN):
        # Abgelaufen (z.B. Browser-Tab geschlossen, PC abgestuerzt) ->
        # als frei behandeln und gleich aufraeumen.
        os.remove(pfad)
        return None
    return lock

def acquire_lock(draft_id, user):
    """Versucht, den Draft fuer `user` zu sperren. True = gesperrt
    (neu oder bereits von `user` selbst gehalten), False = von
    jemand anderem gesperrt (siehe lock_status() fuer Details)."""
    bestehend = lock_status(draft_id)
    if bestehend and bestehend["user"] != user:
        return False
    jetzt = _jetzt()
    lock = {"user": user, "seit": bestehend["seit"] if bestehend else jetzt, "heartbeat": jetzt}
    _atomar_schreiben(_lock_pfad(draft_id), lock)
    return True

def heartbeat(draft_id, user):
    """Erneuert den Lock, falls er `user` gehoert. Gibt False zurueck,
    wenn der Draft inzwischen von jemand anderem gesperrt wurde
    (Browser sollte dann eine Warnung anzeigen)."""
    return acquire_lock(draft_id, user)

def release_lock(draft_id, user):
    bestehend = lock_status(draft_id)
    if bestehend and bestehend["user"] == user:
        os.remove(_lock_pfad(draft_id))
        return True
    return False
