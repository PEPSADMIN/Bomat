@echo off
title PEPS BOM Tool - Live Server
color 0A
cd /d "D:\Hari JR. DATA\Development\Bom Tool"

REM ── SINGLE-INSTANCE GUARD ──────────────────────────────────────────────────
REM Prevent a SECOND run_forever.bat loop from ever starting. Two loops spawn
REM two competing servers on port 5020, which crash each other (exit code -1)
REM in a restart loop and make the tool appear "not loading". Exit immediately
REM if another cmd.exe is already running this script. The launching cmd (this
REM script's own parent) is excluded so we don't false-positive on ourselves.
powershell -NoProfile -Command "$me=[System.Diagnostics.Process]::GetCurrentProcess().Id; $pp=(Get-CimInstance Win32_Process -Filter \"ProcessId=$me\").ParentProcessId; $others=@(Get-CimInstance Win32_Process | Where-Object{$_.Name -eq 'cmd.exe' -and $_.CommandLine -like '*run_forever.bat*' -and $_.ProcessId -ne $me -and $_.ProcessId -ne $pp}); if($others.Count -gt 0){exit 1}else{exit 0}" < NUL >nul 2>&1
if %errorlevel%==1 (
    echo [%date% %time%] Another run_forever.bat loop is already running -- refusing to start a second. Exiting.
    echo [%date% %time%] Another run_forever.bat loop is already running -- refusing to start a second. Exiting. >> server_loop.log
    exit /b 0
)

REM ── STARTUP GUARD ────────────────────────────────────────────────────────
REM If a healthy BOM server already answers on port 5020, this instance is
REM redundant — exit immediately instead of "entering the monitor loop".
REM :loop always launches python.exe unconditionally regardless of how it's
REM reached, so "skip to the loop" used to mean "retry-and-back-off every
REM 60s forever" (see the EXIT_CODE==99 branch below), not actually just
REM watching. That silently left a second, permanent, harmless-but-noisy
REM run_forever.bat process behind every time this guard was hit — the
REM watchdog task (every 5 min) already restarts the loop if it ever dies,
REM so no standby loop is needed here for that.
powershell -NoProfile -Command "try{$null=Invoke-WebRequest 'http://127.0.0.1:5020/api/health' -TimeoutSec 3 -UseBasicParsing -EA Stop;exit 0}catch{exit 1}" < NUL >nul 2>&1
if %errorlevel%==0 (
    echo [%date% %time%] Healthy server found on port 5020 -- another instance is already serving. Exiting.
    echo [%date% %time%] Healthy server found on port 5020 -- another instance is already serving. Exiting. >> server_loop.log
    exit /b 0
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
