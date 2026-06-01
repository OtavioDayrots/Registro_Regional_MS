@echo off
start /B pythonw "RegistroLocalidade.py"

start /B pythonw "RegistroRegional.py"

schtasks /create /tn "Python" /tr "%~dp0executar_registro.bat" 