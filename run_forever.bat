@echo off
REM Keeps the PEPS BOM Tool server running 24/7 without manual intervention.
REM If the server process exits for any reason (crash, kill, error), it is
REM relaunched automatically after a short delay - forever, no retry limit.
REM Logs each start/stop to server_loop.log for after-the-fact diagnosis.

cd /d "D:\Hari JR. DATA\Development\Bom Tool"

:loop
echo [%date% %time%] Starting BOM Tool server... >> server_loop.log
"D:\Hari JR. DATA\Development\Bom Tool\.venv\Scripts\python.exe" app_v2_1.py >> server_loop.log 2>&1
echo [%date% %time%] Server exited (code %errorlevel%). Restarting in 5 seconds... >> server_loop.log
timeout /t 5 /nobreak >nul
goto loop
