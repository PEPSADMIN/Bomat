@echo off
echo ============================================================
echo PEPS BOM AUTOMATION TOOL v2.1 - Starting Server
echo ============================================================
echo.
echo Starting Flask development server...
echo Access at: http://localhost:5005
echo.
echo Press Ctrl+C to stop the server
echo ============================================================
echo.

"%~dp0.venv\Scripts\python.exe" app_v2_1.py

pause
