@echo off
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File "release\setup-portable-bridge.ps1"
pause
