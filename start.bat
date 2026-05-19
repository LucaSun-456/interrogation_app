@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

title Interrogation Experiment

echo ========================================
echo   Interrogation Experiment System
echo ========================================
echo.

:: Check Python
where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found!
    echo Please install Python 3 and add it to PATH.
    echo Download: https://www.python.org/downloads/
    goto :error
)

:: Check/create venv
if not exist ".venv\Scripts\python.exe" (
    echo [1/3] Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to create venv
        goto :error
    )
) else (
    echo [1/3] Virtual environment already exists
)

:: Install dependencies
echo [2/3] Installing dependencies...
.venv\Scripts\python.exe -m pip install -q -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Dependency install failed
    goto :error
)

:: DeepSeek API Key (pre-configured)
set "DEEPSEEK_API_KEY=sk-7a8a601d58b64900a6cfe3f2a0110e7c"

:: Kill old process on port 5000
echo.
echo [Check] Checking if port 5000 is available...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5000 " ^| findstr LISTENING 2^>nul') do (
    echo [WARN] Port 5000 is in use by PID %%a, killing...
    taskkill /F /PID %%a >nul 2>&1
    if errorlevel 1 (
        echo [WARN] Could not kill old process, please close it manually
    ) else (
        echo [OK] Old process terminated
    )
    timeout /t 1 /nobreak >nul
)

:: Start server
echo.
echo [3/3] Starting server...
echo.

:: Try to get local IP
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4" ^| findstr /v "127.0.0.1" ^| findstr /v "fe80" ^| findstr /v "::" 2^>nul') do (
    set "TRYIP=%%a"
    set TRYIP=!TRYIP: =!
    if not "!TRYIP!"=="" set "LOCAL_IP=!TRYIP!" & goto :start_server
)
:: Fallback: try Chinese Windows format
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4" 2^>nul') do (
    set "TRYIP=%%a"
    set TRYIP=!TRYIP: =!
    if not "!TRYIP!"=="127.0.0.1" set "LOCAL_IP=!TRYIP!" & goto :start_server
)
:: Final fallback
set "LOCAL_IP=127.0.0.1"

:start_server
echo ========================================
echo   Access URLs:
echo.
echo   Local:       http://localhost:5000
echo   Management:  http://localhost:5000/manage
echo   LAN:         http://%LOCAL_IP%:5000
echo.
echo   Press Ctrl+C to stop the server
echo ========================================
echo.

:: Open browser
start "" http://localhost:5000
.venv\Scripts\python.exe app.py

echo.
echo Server stopped. Close this window or press any key to restart.
pause
exit /b 0

:error
echo.
echo Press any key to exit...
pause >nul
exit /b 1
