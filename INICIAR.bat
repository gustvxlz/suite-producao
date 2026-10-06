@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    py suite_producao.py
) else (
    python suite_producao.py
)
if errorlevel 1 pause
