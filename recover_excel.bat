@echo off
setlocal EnableDelayedExpansion

cd /d "%~dp0"
if errorlevel 1 (
    echo [ERROR] Cannot cd to: %~dp0
    goto done
)

title Excel Recovery

echo ========================================
echo   Recover experiment_data.xlsx
echo ========================================
echo.
echo Project: %CD%
echo.

set "ROOT=%~dp0"
set "PY="

if exist "%ROOT%.venv\Scripts\python.exe" (
    set "PY=%ROOT%.venv\Scripts\python.exe"
    echo [OK] Using venv Python
) else (
    where python >nul 2>&1
    if not errorlevel 1 (
        set "PY=python"
        echo [OK] Using system Python
    )
)

if not defined PY (
    echo [ERROR] Python not found. Run start.bat first or install Python 3.
    goto done
)

if not exist "%ROOT%scripts\recover_excel.py" (
    echo [ERROR] Missing scripts\recover_excel.py
    goto done
)

"%PY%" -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installing openpyxl...
    "%PY%" -m pip install openpyxl
    if errorlevel 1 (
        echo [ERROR] pip install openpyxl failed
        goto done
    )
)

if not exist "%ROOT%data" mkdir "%ROOT%data"

set "INPUT="
set "OUTPUT=%ROOT%data\experiment_data_recovered.xlsx"

if not "%~1"=="" (
    set "INPUT=%~1"
    goto recover
)

set "NEWEST="
for /f "delims=" %%F in ('dir /b /o-d "%ROOT%data\experiment_data.xlsx.corrupt.*" 2^>nul') do (
    if not defined NEWEST set "NEWEST=%ROOT%data\%%F"
)
if not defined NEWEST (
    for /f "delims=" %%F in ('dir /b /o-d "%ROOT%data\*.corrupt.*" 2^>nul') do (
        if not defined NEWEST set "NEWEST=%ROOT%data\%%F"
    )
)

if defined NEWEST (
    echo Found backup:
    echo   !NEWEST!
    echo.
    set /p "USE_AUTO=Use this file [Y/n]: "
    if /i "!USE_AUTO!"=="" set "USE_AUTO=Y"
    if /i "!USE_AUTO!"=="Y" (
        set "INPUT=!NEWEST!"
        goto recover
    )
)

echo.
echo Put corrupt file in data\ folder, or type path below.
echo Example: data\experiment_data.xlsx.corrupt.20260611_174451
echo.
set /p "INPUT=Corrupt file path: "
if "!INPUT!"=="" (
    echo Cancelled.
    goto done
)
if not exist "!INPUT!" (
    if exist "%ROOT%!INPUT!" set "INPUT=%ROOT%!INPUT!"
)

:recover
if not exist "!INPUT!" (
    echo [ERROR] File not found: !INPUT!
    goto done
)

if exist "!OUTPUT!" (
    set "OUTPUT=%ROOT%data\experiment_data_recovered_%RANDOM%.xlsx"
)

echo.
echo Recovering...
echo   In:  !INPUT!
echo   Out: !OUTPUT!
echo.

"%PY%" "%ROOT%scripts\recover_excel.py" "!INPUT!" -o "!OUTPUT!"
if errorlevel 1 (
    echo.
    echo [FAILED] Try Excel Open and Repair, or 7-Zip extract.
    goto done
)

echo.
echo [OK] Saved to:
echo   !OUTPUT!
echo.
echo Open in Excel to verify, then replace server data\experiment_data.xlsx
echo.
set /p "OPENIT=Open in Explorer [Y/n]: "
if /i "!OPENIT!"=="" set "OPENIT=Y"
if /i "!OPENIT!"=="Y" explorer /select,"!OUTPUT!"

:done
echo.
echo ========================================
pause
endlocal
