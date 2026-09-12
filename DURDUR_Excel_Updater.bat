@echo off
title Excel SQL Server IP Guncelleyici - Sunucuyu Durdur
cd /d "%~dp0"

echo =================================================================
echo  Excel Bulk Updater Sunucusu Kapatiliyor...
echo =================================================================
echo.

"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -Command "try { $res = Invoke-RestMethod -Uri 'http://127.0.0.1:3005/api/shutdown' -Method POST -TimeoutSec 3 -ErrorAction Stop; Write-Host 'Sunucu basariyla kapatildi.' -ForegroundColor Green } catch { Write-Host 'Sunucu zaten calismiyor veya port 3005 aktif degil.' -ForegroundColor Yellow }"

echo.
timeout /t 3 >nul
