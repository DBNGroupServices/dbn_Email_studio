@echo off
rem Finds and fixes the cause of "Sorry, we can't load the add-in".
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Diagnose-EmailStudio.ps1" -Repair
pause
