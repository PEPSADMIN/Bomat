@echo off
title PEPS BOM Tool - Live Server
color 0A
cd /d "D:\Hari JR. DATA\Development\Bom Tool"

REM Kill any zombie process holding port 5020 before starting
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

"D:\Hari JR. DATA\Development\Bom Tool\.venv\Scripts\python.exe" app_v2_1.py

echo.
echo [%date% %time%] Server exited (code %errorlevel%). Restarting in 5 seconds...
echo [%date% %time%] Server exited (code %errorlevel%). Restarting in 5 seconds... >> server_loop.log
timeout /t 5 /nobreak >nul
goto loop
