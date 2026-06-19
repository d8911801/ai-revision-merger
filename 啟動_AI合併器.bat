@echo off
chcp 65001 >nul
title AI 學術修訂合併器

REM 切換到專案目錄
cd /d "%~dp0"

REM 啟動 Python GUI
"C:\Users\d8911801\.workbuddy\binaries\python\versions\3.13.12\python.exe" main.py

pause
