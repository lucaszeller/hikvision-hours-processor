@echo off
setlocal
cd /d "%~dp0"

if exist "venv\Scripts\pythonw.exe" (
    start "" "venv\Scripts\pythonw.exe" "main.py"
    goto :eof
)

if exist "venv\Scripts\python.exe" (
    start "" "venv\Scripts\python.exe" "main.py"
) else (
    where pythonw >nul 2>nul
    if not errorlevel 1 (
        start "" pythonw "main.py"
    ) else (
        start "" python "main.py"
    )
)
