@echo off
title PEPS BOM Tool - Live Server
color 0A
cd /d "D:\Hari JR. DATA\Development\Bom Tool"

REM ── STARTUP: kill ONLY a zombie on port 5020 (no response within 3 s) ──────
REM A healthy server gets left alone — we will detect it in the loop (exit 99)
REM and wait without crashing.  This protects any other servers on other ports.
for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":5020 .*LISTENING"') do (
    powershell -NoProfile -Command ^
      "try{$null=(Invoke-WebRequest 'https://127.0.0.1:5020/api/health' -TimeoutSec 3 -SkipCertificateCheck -UseBasicParsing -EA Stop);exit 1}catch{exit 0}" >nul 2>&1
    if errorlevel 1 (
        echo [%date% %time%] Healthy server on port 5020 ^(PID %%p^) — leaving it running.
        echo [%date% %time%] Healthy server on port 5020 ^(PID %%p^) — leaving it running. >> server_loop.log
    ) else (
        echo [%date% %time%] Killing zombie PID %%p on port 5020...
        echo [%date% %time%] Killing zombie PID %%p on port 5020... >> server_loop.log
        taskkill /F /PID %%p >nul 2>&1
        timeout /t 3 /nobreak >nul
    )
)

:loop
echo.
echo ============================================================
echo  [%date% %time%] Starting BOM Tool server...
echo ============================================================
echo [%date% %time%] Starting BOM Tool server... >> server_loop.log

".venv\Scripts\python.exe" app_v2_1.py
set EXIT_CODE=%errorlevel%

REM Exit 99 = port already held by a healthy BOM Tool instance (not a crash).
REM Wait 60 s and retry — if that instance dies, we take over automatically.
if %EXIT_CODE%==99 (
    echo [%date% %time%] Port 5020 busy -- another instance running. Waiting 60 s... >> server_loop.log
    echo [%date% %time%] Port 5020 busy -- another instance running. Waiting 60 s...
    timeout /t 60 /nobreak >nul
    goto loop
)

echo.
echo [%date% %time%] Server exited (code %EXIT_CODE%). Restarting in 5 seconds...
echo [%date% %time%] Server exited (code %EXIT_CODE%). Restarting in 5 seconds... >> server_loop.log
timeout /t 5 /nobreak >nul
goto loop
