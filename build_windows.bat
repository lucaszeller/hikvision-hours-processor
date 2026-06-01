@echo off
setlocal

if exist dist\hikvision-hours-processor rmdir /S /Q dist\hikvision-hours-processor
if exist dist\hikvision-hours-processor.exe del /Q dist\hikvision-hours-processor.exe

python -m PyInstaller --clean --noconfirm hikvision_hours_processor.spec
if errorlevel 1 (
    echo Error generando el ejecutable.
    exit /b 1
)

if not exist dist mkdir dist

if exist date.xlsx copy /Y date.xlsx dist\date.xlsx >nul
if exist info.xlsx copy /Y info.xlsx dist\info.xlsx >nul
if exist feriados_nacionales_argentina_2026.xlsx copy /Y feriados_nacionales_argentina_2026.xlsx dist\feriados_nacionales_argentina_2026.xlsx >nul
if exist logo.png copy /Y logo.png dist\logo.png >nul
if exist app.ico copy /Y app.ico dist\app.ico >nul
if exist MANUAL_USUARIO_APP_DATE.xlsx.docx copy /Y MANUAL_USUARIO_APP_DATE.xlsx.docx dist\MANUAL_USUARIO_APP_DATE.xlsx.docx >nul

if exist dist\hikvision-hours-processor-portable.zip del /Q dist\hikvision-hours-processor-portable.zip
powershell -NoProfile -ExecutionPolicy Bypass -Command "$items = @('dist\hikvision-hours-processor.exe','dist\date.xlsx','dist\info.xlsx','dist\feriados_nacionales_argentina_2026.xlsx','dist\logo.png','dist\app.ico','dist\MANUAL_USUARIO_APP_DATE.xlsx.docx') | Where-Object { Test-Path -LiteralPath $_ }; Compress-Archive -Path $items -DestinationPath 'dist\hikvision-hours-processor-portable.zip' -Force"

echo.
echo Ejecutable generado:
echo   dist\hikvision-hours-processor.exe
echo   dist\hikvision-hours-processor-portable.zip
echo.
echo Para usarlo en otra PC, copiar la carpeta dist completa.

endlocal
