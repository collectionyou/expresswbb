@echo off
setlocal
set "ROOT=%~dp0"

powershell -ExecutionPolicy Bypass -File "%ROOT%stop_order_analysis_v1.ps1" >nul 2>nul
powershell -ExecutionPolicy Bypass -File "%ROOT%run_order_analysis_v1.ps1" -Background
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:8791"

endlocal
