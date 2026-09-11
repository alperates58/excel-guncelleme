@echo off
title Excel SQL Server IP ve Baglanti Guncelleyici
cd /d "%~dp0"

echo =================================================================
echo  Excel SQL Server IP ve Baglanti Dizesi Toplu Guncelleyici
echo =================================================================
echo.
echo Sunucu baslatiliyor ve tarayici aciliyor... Lutfen bekleyin.
echo (Kapatmak istediginizde bu pencereyi kapatabilirsiniz.)
echo.

"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1" -Port 3005

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo =================================================================
    echo HATA: Sunucu baslatilamadi veya beklenmeyen bir hata ile sonlandi!
    echo Hata Kodu: %ERRORLEVEL%
    echo =================================================================
    pause
)

