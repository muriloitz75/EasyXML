@echo off
echo Iniciando EasyXML - Processador de Notas Fiscais...
echo.

REM Verifica se o executável existe
if exist "dist\EasyXML\EasyXML.exe" (
    start "" "dist\EasyXML\EasyXML.exe"
) else (
    echo Erro: Executável não encontrado.
    echo Verifique se o arquivo "dist\EasyXML\EasyXML.exe" existe.
    echo.
    pause
)
