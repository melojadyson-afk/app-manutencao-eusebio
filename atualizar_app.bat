@echo off
chcp 65001 >nul
cd /d "C:\Users\mello\OneDrive\Documentos\GitHub\app-manutencao-eusebio"

set "PLANILHA_PATH=C:\Users\mello\ELIS\SpO-BR-Management-Eusébio - Manutenção\26 - PCM\Salomão\01 - Jadyson\01- App Manutenção\Planilha Integrada de Indicadores APP Manutenção.xlsx"

echo Caminho configurado:
echo %PLANILHA_PATH%
echo.

echo Rodando extracao...
python scripts\extract.py
if errorlevel 1 (
    echo Falha na extracao. Abortando.
    exit /b 1
)

echo Publicando no GitHub...
git add docs/index.html
git commit -m "Atualizacao automatica local"
git push

echo Concluido.