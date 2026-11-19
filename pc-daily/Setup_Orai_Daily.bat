@echo off
title Orai Fieldwork App - daily Billed/Unbilled update setup
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Orai_Daily_Update.ps1" -Setup
echo.
pause
