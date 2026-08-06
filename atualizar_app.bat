@echo off
chcp 65001 >nul
cd /d "C:\Users\mello\OneDrive\Documentos\GitHub\app-manutencao-eusebio"

echo Procurando a planilha...
for /f "delims=" %%F in ('powershell -NoProfile -Command "Get-ChildItem -Path 'C:\Users\mello\ELIS' -Filter *.xlsx -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.FullName -like '*App Manuten*' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName"') do set "PLANILHA_PATH=%%F"

if "%PLANILHA_PATH%"=="" (
    echo Nao encontrei nenhuma planilha com "App Manuten" no caminho, dentro de C:\Users\mello\ELIS
    echo Abortando.
    exit /b 1
)

echo Planilha encontrada:
echo %PLANILHA_PATH%
echo.

echo Rodando extracao...
python scripts\extract.py
if errorlevel 1 (
    echo Falha na extracao. Abortando.
    exit /b 1
)

echo Publicando no GitHub...
git add .
git commit -m "Atualizacao automatica local"
git push

echo Concluido.