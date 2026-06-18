@echo off
setlocal
set "ROOT=%~dp0"
powershell -ExecutionPolicy Bypass -File "%ROOT%stop_order_analysis_v1.ps1"
endlocal
