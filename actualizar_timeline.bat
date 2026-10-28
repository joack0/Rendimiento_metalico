@echo off
rem Regenera dashboard\timeline.html con todos los CSV de datos\interrupciones y lo abre.
cd /d "%~dp0"
python scripts\preparar_timeline.py
if errorlevel 1 pause & exit /b 1
start "" "dashboard\timeline.html"
