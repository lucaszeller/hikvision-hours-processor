@echo off
python -m PyInstaller --clean hikvision_hours_processor.spec
echo Ejecutable generado en .\dist\hikvision-hours-processor\
