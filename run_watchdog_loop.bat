@echo off
:: Runs the watchdog every 45 minutes in a silent background loop.
:: Added to registry startup so it runs alongside run_forever.bat.
:loop
timeout /t 2700 /nobreak >nul
powershell.exe -ExecutionPolicy Bypass -NonInteractive -WindowStyle Hidden -File "D:\Hari JR. DATA\Development\Bom Tool\watchdog_bom.ps1"
goto loop
