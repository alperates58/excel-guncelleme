@echo off
title Excel SQL Server IP ve Baglanti Guncelleyici
cd /d "%~dp0"

echo =================================================================
echo  Excel SQL Server IP ve Baglanti Dizesi Toplu Guncelleyici
echo =================================================================
echo.
echo Sunucu baslatiliyor ve Google Chrome aciliyor... Lutfen bekleyin.
echo (Kapatmak istediginizde bu pencereyi kapatabilirsiniz.)
echo.

if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (
    start "" "%ProgramFiles%\Google\Chrome\Application\chrome.exe" "http://127.0.0.1:3005/"
) else if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" (
    start "" "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" "http://127.0.0.1:3005/"
) else if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" (
    start "" "%LocalAppData%\Google\Chrome\Application\chrome.exe" "http://127.0.0.1:3005/"
) else (
    start "" "http://127.0.0.1:3005/"
)

"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1" -Port 3005

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo =================================================================
    echo HATA: Sunucu baslatilamadi veya beklenmeyen bir hata ile sonlandi!
    echo Hata Kodu: %ERRORLEVEL%
    echo =================================================================
    pause
)

