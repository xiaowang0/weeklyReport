@echo off
chcp 65001 >nul
echo 启动周报系统...
cd /d "%~dp0"
python app.py
pause