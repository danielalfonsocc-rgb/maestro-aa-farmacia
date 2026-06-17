@echo off
title AUTO SSASUR - Maestro AA Farmacia
cd /d "%~dp0"

echo.
echo  ============================================================
echo   AUTO SSASUR  -  Maestro AA Farmacia
echo  ============================================================
echo.

:: ── Python ─────────────────────────────────────────────────────────────────
py --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado. Ejecuta INSTALAR.bat primero.
    pause & exit /b 1
)

:: ── Playwright ─────────────────────────────────────────────────────────────
py -c "import playwright" >nul 2>&1
if errorlevel 1 (
    echo  Instalando Playwright...
    py -m pip install playwright --quiet
    if errorlevel 1 (
        echo  [ERROR] No se pudo instalar playwright.
        pause & exit /b 1
    )
)

:: ── Chromium ───────────────────────────────────────────────────────────────
py -c "from playwright.sync_api import sync_playwright; p=sync_playwright().start(); p.chromium.executable_path; p.stop()" >nul 2>&1
if errorlevel 1 (
    echo  Instalando navegador Chromium...
    py -m playwright install chromium
)

:: ── Dependencias maestro ───────────────────────────────────────────────────
py -c "import pandas, openpyxl, numpy" >nul 2>&1
if errorlevel 1 (
    echo  Instalando dependencias...
    py -m pip install -r requirements.txt --quiet
)

echo.
echo  Iniciando automatizacion...
echo.

py AUTO_SSASUR.py

echo.
