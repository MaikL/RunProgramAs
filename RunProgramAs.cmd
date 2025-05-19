@echo off
set SCRIPT_DIR=%~dp0
start /min "" powershell.exe -ExecutionPolicy Bypass -File "%SCRIPT_DIR%RunProgramAs.ps1"
