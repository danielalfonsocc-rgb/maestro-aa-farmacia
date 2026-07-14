@echo off
title AUTO SSASUR - Maestro AA Farmacia
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

echo.
echo  ============================================================
echo   AUTO SSASUR  -  Maestro AA Farmacia
echo  ============================================================
echo.

py --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado.
    pause & exit /b 1
)

py -c "import playwright" >nul 2>&1
if errorlevel 1 (
    echo  Instalando Playwright...
    py -m pip install playwright --quiet
)

py -c "from playwright.sync_api import sync_playwright; p=sync_playwright().start(); p.chromium.executable_path; p.stop()" >nul 2>&1
if errorlevel 1 (
    echo  Instalando navegador Chromium...
    py -m playwright install chromium
)

py -c "import pandas, openpyxl, numpy" >nul 2>&1
if errorlevel 1 (
    echo  Instalando dependencias...
    py -m pip install -r requirements.txt --quiet
)

echo.
echo  Iniciando automatizacion...
echo.

py AUTO_SSASUR.py %* 2>> auto_ssasur_error.log

echo.
echo  ============================================================
if errorlevel 1 (
    echo   [ERROR] Revisa auto_ssasur_error.log
    echo.
    type auto_ssasur_error.log
) else (
    echo   Completado.
)
echo  ============================================================
echo.
pause
