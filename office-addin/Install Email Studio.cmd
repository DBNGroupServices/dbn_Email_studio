@echo off
rem Installs or upgrades DBN Email Studio for Outlook Classic (this Windows user only, no admin needed).
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-EmailStudio.ps1" %*
pause
