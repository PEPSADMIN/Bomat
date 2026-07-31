@echo off
title PEPS BOM Tool - Live Server
color 0A
cd /d "D:\Hari JR. DATA\Development\Bom Tool"

REM ── STARTUP GUARD ────────────────────────────────────────────────────────
REM If a healthy BOM server already answers on port 5020, skip straight to
REM the monitor loop — do NOT kill it.  This protects an existing instance
REM started by Task Scheduler (or a previous loop) from being torn down.
powershell -NoProfile -Command "try{$null=Invoke-WebRequest 'https://127.0.0.1:5020/api/health' -TimeoutSec 3 -SkipCertificateCheck -UseBasicParsing -EA Stop;exit 0}catch{exit 1}" >nul 2>&1
if %errorlevel%==0 (
    echo [%date% %time%] Healthy server found on port 5020 -- skipping kill, entering monitor loop.
    echo [%date% %time%] Healthy server found on port 5020 -- skipping kill, entering monitor loop. >> server_loop.log
    goto loop
)

REM No healthy server -- kill any zombie holding the port, then start fresh.
for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":5020.*LISTENING"') do (
    echo [%date% %time%] Killing zombie PID %%p on port 5020...
    echo [%date% %time%] Killing zombie PID %%p on port 5020... >> server_loop.log
    taskkill /F /PID %%p >nul 2>&1
)
timeout /t 3 /nobreak >nul

:loop
echo.
echo ============================================================
echo  [%date% %time%] Starting BOM Tool server...
echo ============================================================
echo [%date% %time%] Starting BOM Tool server... >> server_loop.log

".venv\Scripts\python.exe" app_v2_1.py
set EXIT_CODE=%errorlevel%

REM Exit 99 = port 5020 is held by another healthy BOM Tool process (not a crash).
REM Wait 60 s and retry so we take over automatically if that process ever dies.
if %EXIT_CODE%==99 (
    echo [%date% %time%] Port 5020 busy ^(another instance running^). Waiting 60 s before retry...
    echo [%date% %time%] Port 5020 busy -- another instance running. Waiting 60 s. >> server_loop.log
    timeout /t 60 /nobreak >nul
    goto loop
)

echo.
echo [%date% %time%] Server exited ^(code %EXIT_CODE%^). Restarting in 5 seconds...
echo [%date% %time%] Server exited (code %EXIT_CODE%). Restarting in 5 seconds... >> server_loop.log
timeout /t 5 /nobreak >nul
goto loop
