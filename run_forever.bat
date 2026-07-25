@echo off
title PEPS BOM Tool - Live Server
color 0A
cd /d "D:\Hari JR. DATA\Development\Bom Tool"

REM Only one instance should run. If port 5020 is already listening, exit.
netstat -ano | findstr ":5020.*LISTENING" >nul 2>&1
if %errorlevel% == 0 (
    echo [%date% %time%] Port 5020 already in use -- another instance is running. Exiting.
    echo [%date% %time%] Duplicate run_forever exited - port already in use. >> server_loop.log
    exit /b 0
)

:loop
echo.
echo ============================================================
echo  [%date% %time%] Starting BOM Tool server...
echo ============================================================
echo [%date% %time%] Starting BOM Tool server... >> server_loop.log

"D:\Hari JR. DATA\Development\Bom Tool\.venv\Scripts\python.exe" app_v2_1.py

echo.
echo [%date% %time%] Server exited (code %errorlevel%). Restarting in 5 seconds...
echo [%date% %time%] Server exited (code %errorlevel%). Restarting in 5 seconds... >> server_loop.log
timeout /t 5 /nobreak >nul
goto loop
