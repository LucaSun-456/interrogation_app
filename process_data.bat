@echo off
echo ===================================================
echo Processing experiment data...
echo ===================================================

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" "scripts\process_data.py"
) else (
    python "scripts\process_data.py"
)

echo.
pause
