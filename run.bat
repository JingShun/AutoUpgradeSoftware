@echo off
setlocal

:: 批次檔所在目錄
set "SCRIPT_DIR=%~dp0"

:: 呼叫 PowerShell 執行 .ps1 並繞過執行政策
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%upgrade7z.ps1"
