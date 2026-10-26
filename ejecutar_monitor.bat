@echo off
REM Revisa las normas de hoy en El Peruano, genera el Word y lo envía.
REM Para otro rango:  ejecutar_monitor.bat --dias 7
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" monitor_normas.py %*
