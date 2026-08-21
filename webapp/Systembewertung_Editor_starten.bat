@echo off
REM Startet den Systembewertung-Editor lokal und oeffnet den Browser.
REM Doppelklick reicht - kein Server, keine Installation ausser den
REM einmalig noetigen Python-Paketen (siehe requirements.txt/README.md).
cd /d "%~dp0"
python app.py
REM Fenster schliesst sich automatisch, wenn app.py sich selbst sauber
REM beendet hat (Auto-Beenden, sobald kein Browser-Tab mehr offen ist -
REM Exit-Code 0). Bei einem Fehler/Absturz (Exit-Code ungleich 0)
REM bleibt das Fenster offen, damit die Fehlermeldung lesbar ist.
if errorlevel 1 pause
