@echo off
REM Startet den Systembewertung-Editor lokal und oeffnet den Browser.
REM Doppelklick reicht - kein Server, keine Installation ausser den
REM einmalig noetigen Python-Paketen (siehe requirements.txt/README.md).
cd /d "%~dp0"
python app.py
pause
