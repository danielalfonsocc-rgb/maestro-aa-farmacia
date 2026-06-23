@echo off
title Gestion Territorial - Farmacia Hospital de Pitrufquen
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

echo.
echo  ============================================================
echo   GESTION TERRITORIAL  -  Solo GT (sin stock ni maestro)
echo   Rango: ayer y hoy  ^|  Clasifica y genera planillas
echo   Para corrida completa usar AUTO_SSASUR.bat
echo  ============================================================
echo.

:: ── Python ──────────────────────────────────────────────────────────
py --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado. Ejecuta INSTALAR.bat primero.
    pause & exit /b 1
)

:: ── Playwright ──────────────────────────────────────────────────────
py -c "import playwright" >nul 2>&1
if errorlevel 1 (
    echo  Instalando Playwright...
    py -m pip install playwright --quiet
    if errorlevel 1 (
        echo  [ERROR] No se pudo instalar playwright.
        pause & exit /b 1
    )
)

py -c "from playwright.sync_api import sync_playwright; p=sync_playwright().start(); p.chromium.executable_path; p.stop()" >nul 2>&1
if errorlevel 1 (
    echo  Instalando navegador Chromium...
    py -m playwright install chromium
)

:: ── Dependencias ────────────────────────────────────────────────────
py -c "import openpyxl, PIL" >nul 2>&1
if errorlevel 1 (
    echo  Instalando dependencias...
    py -m pip install -r requirements.txt --quiet
)

echo.
echo  Iniciando... (debes logarte en SSASUR cuando abra el navegador)
echo.

py AUTO_SSASUR.py --gt

echo.
if %errorlevel% == 0 (
    echo  Copiando planillas al Escritorio...
    py publicar_escritorio.py --gt
    echo  ============================================================
    echo   LISTO
    echo   Planillas en: out_gt\
    echo   Copia en el Escritorio: Farmacia AA\2 - Gestion Territorial
    echo  ============================================================
) else (
    echo  [AVISO] Revisa los mensajes arriba.
)
echo.
pause
