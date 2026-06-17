@echo off
title Instalando Playwright para AUTO SSASUR
cd /d "%~dp0"

echo.
echo  ============================================================
echo   Instalando Playwright  (necesario para AUTO_SSASUR)
echo  ============================================================
echo.

:: Verificar Python
py --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado.
    echo  Descarga Python desde https://python.org e instala marcando
    echo  la opcion "Add Python to PATH".
    pause & exit /b 1
)

echo  Python encontrado:
py --version
echo.

:: Instalar Playwright
echo  Instalando playwright...
py -m pip install playwright
if errorlevel 1 (
    echo  [ERROR] Fallo la instalacion de playwright.
    pause & exit /b 1
)
echo.

:: Instalar Chromium
echo  Instalando navegador Chromium (puede tardar 2-3 minutos)...
py -m playwright install chromium
if errorlevel 1 (
    echo  [ERROR] Fallo la instalacion de Chromium.
    pause & exit /b 1
)
echo.

echo  ============================================================
echo   LISTO - Playwright instalado correctamente
echo   Ahora puedes usar AUTO_SSASUR.bat
echo  ============================================================
echo.
pause
