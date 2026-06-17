@echo off
title Instalador - Maestro AA Farmacia AT Abierta
cd /d "%~dp0"

echo  ============================================================
echo   INSTALADOR - Maestro AA
echo   Farmacia AT Abierta - Pitrufquen
echo   Solo necesitas ejecutar esto UNA VEZ por computador.
echo  ============================================================
echo.

:: ── Verificar Python ──────────────────────────────────────────────
py --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no esta instalado.
    echo.
    echo  Por favor instala Python desde:
    echo  https://www.python.org/downloads/
    echo.
    echo  IMPORTANTE: Durante la instalacion marca la casilla
    echo  "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('py --version 2^>^&1') do echo  Python detectado: %%v
echo.

:: ── Instalar dependencias ─────────────────────────────────────────
echo  Instalando dependencias (puede tomar 1-2 minutos)...
echo.

py -m pip install --upgrade pip --quiet

py -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo.
    echo  [ERROR] No se pudieron instalar las dependencias.
    echo  Verifica que tienes conexion a internet.
    pause
    exit /b 1
)

echo.
echo  ============================================================
echo   INSTALACION COMPLETADA
echo.
echo   Ahora puedes usar:
echo.
echo   > ACTUALIZAR_DATOS.bat
echo     Para generar el consolidado con tus archivos de SSASUR
echo.
echo   > ABRIR_APP.bat
echo     Para abrir el dashboard interactivo
echo  ============================================================
echo.
pause
