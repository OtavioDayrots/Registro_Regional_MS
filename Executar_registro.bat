@echo off
cd /d "%~dp0"

python "RegistroLocalidade.py"
if errorlevel 1 goto :end

python "RegistroRegional.py"
if errorlevel 1 goto :end

rem Adiciona e commita apenas se houver alterações
git add .
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "Atualiza arquivos Excel de registro"
    git push origin main
) else (
    echo Nada para commitar.
)

rem Mantém o agendamento existente
schtasks /create /tn "Python" /tr "%~dp0executar_registro.bat"

:end 