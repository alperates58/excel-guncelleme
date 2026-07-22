@echo off
chcp 65001 > nul
title Excel SQL Connection & IP Updater

echo =================================================================
echo  Excel SQL Server IP ve Baglanti Dizesi Toplu Guncelleyici
echo =================================================================
echo.
echo Sunucu baslatiliyor... Lutfen bekleyin.
echo.

start "" powershell.exe -ExecutionPolicy Bypass -File "%~dp0server.ps1" -Port 3005

timeout /t 2 /nobreak > nul
start http://localhost:3005/

echo.
echo Uygulama tarayicinizda acildi: http://localhost:3005/
echo Bu pencereyi kapatmayin. Kapatmak icin Ctrl+C basabilirsiniz.
echo.
