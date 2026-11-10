@echo off
rem Toma una captura ahora mismo y abre el dashboard.
cd /d "%~dp0"
python rastreador.py
start "" "dashboard.html"
