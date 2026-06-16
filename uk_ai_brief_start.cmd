@echo off
chcp 65001 > nul
cd /d "%~dp0"

python ".\代码\uk_ai_brief.py" %*

if "%~1"=="" pause
