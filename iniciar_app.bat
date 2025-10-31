@echo off
cd /d "%~dp0"
echo Iniciando Auto-Ficha-OPE...
start "" "%~dp0Python313\pythonw.exe" "%~dp0updater.py"
exit
