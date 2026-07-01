@echo off
:: ─────────────────────────────────────────────────────────────────
:: PEPS BOM Tool — Auto-Start Setup
:: RIGHT-CLICK -> "Run as Administrator"  (one time only)
:: ─────────────────────────────────────────────────────────────────

echo Setting up PEPS BOM Tool auto-start on boot...

schtasks /create /tn "PEPS BOM Tool" /f ^
  /sc ONSTART ^
  /tr "\"D:\Hari JR. DATA\Development\Bom Tool\run_forever.bat\"" ^
  /ru SYSTEM ^
  /rl HIGHEST

if %errorlevel% equ 0 (
    echo.
    echo SUCCESS — Server will auto-start every time the PC is switched on,
    echo and auto-restart itself within seconds if it ever crashes.
) else (
    echo.
    echo FAILED — Make sure you right-clicked and chose "Run as Administrator".
)

echo.
echo Setting up Watchdog (every 45 minutes)...

schtasks /create /tn "PEPS BOM Watchdog" /f ^
  /sc MINUTE /mo 45 ^
  /tr "\"D:\Hari JR. DATA\Development\Bom Tool\run_watchdog.bat\"" ^
  /ru SYSTEM ^
  /rl HIGHEST

if %errorlevel% equ 0 (
    echo SUCCESS — Watchdog will check server health every 45 minutes.
) else (
    echo FAILED — Watchdog setup failed.
)

echo.
echo Done. You can close this window.
pause
