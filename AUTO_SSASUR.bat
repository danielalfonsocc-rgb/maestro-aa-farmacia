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

rem stdout se guarda en auto_ssasur_stdout.log (ademas de mostrarse en pantalla)
rem porque los [ERROR]/[AVISO] del paso GT son print() normales, no excepciones,
rem y antes se perdian sin quedar en ningun lado si la corrida "fallaba en silencio".
rem auto_ssasur_error.log se trunca en CADA corrida (2> en vez de 2>>): antes
rem quedaba acumulando errores de corridas anteriores para siempre, y el
rem "type" de abajo mostraba mezclados bugs ya resueltos de hace semanas con
rem el error real de la corrida actual (detectado 30-07-2026).
rem --clozapina va siempre (31-08-2026): antes era opt-in y solo lo activaba
rem CLOZAPINA.bat por separado; ahora la corrida diaria (tarea programada)
rem saca tambien el consolidado de hemogramas sin necesidad de un acceso
rem directo aparte. CLOZAPINA.bat fue eliminado por quedar redundante.
powershell -NoProfile -ExecutionPolicy Bypass -Command "& { py AUTO_SSASUR.py --clozapina %* 2> 'auto_ssasur_error.log' | Tee-Object -FilePath 'auto_ssasur_stdout.log'; exit $LASTEXITCODE }"

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
