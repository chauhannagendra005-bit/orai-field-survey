@echo off
title Orai Fieldwork App - send today's Billed/Unbilled update now
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Orai_Daily_Update.ps1" -Force
echo.
pause
