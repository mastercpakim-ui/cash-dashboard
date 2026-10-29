@echo off
rem Daily dashboard build (called by Windows Task Scheduler). Appends to build.log next to this file.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
echo ==== %date% %time% >> build.log
python -X utf8 build_dashboard.py >> build.log 2>&1
echo exit code %errorlevel% >> build.log
