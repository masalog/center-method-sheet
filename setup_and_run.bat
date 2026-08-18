@echo off
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File insert_image.ps1
pause
