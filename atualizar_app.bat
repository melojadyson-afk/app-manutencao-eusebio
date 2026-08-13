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
for /f "delims=" %%D in ('powershell -NoProfile -Command "(Get-Item '%PLANILHA_PATH%').LastWriteTime.ToString('dd/MM/yyyy HH:mm')"') do set "PLANILHA_DATA=%%D"
echo Ultima modificacao: %PLANILHA_DATA%
echo.
echo ATENCAO: confira se essa data bate com a ultima vez que voce salvou a planilha no SharePoint.
echo Se estiver desatualizada (por exemplo, por causa de internet instavel durante a sincronizacao do OneDrive),
echo feche este terminal, espere sincronizar e rode de novo.
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