@echo off
title Servidor App AA - Red Local
cd /d "%~dp0"

echo  ============================================================
echo   SERVIDOR APP PEDIDOS AA - RED LOCAL
echo   Farmacia AT Abierta - Pitrufquen
echo  ============================================================
echo.

:: ── Python ─────────────────────────────────────────────────────────
py --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado. Ejecuta INSTALAR.bat primero.
    pause & exit /b 1
)

:: ── Dependencias ───────────────────────────────────────────────────
py -c "import streamlit, openpyxl, reportlab, rapidfuzz" >nul 2>&1
if errorlevel 1 (
    echo  Instalando dependencias...
    py -m pip install -r requirements.txt --quiet
)

:: ── Archivo maestro ────────────────────────────────────────────────
if not exist "Consolidado_AA_MAESTRO.xlsx" (
    echo  [ERROR] No se encuentra Consolidado_AA_MAESTRO.xlsx
    echo  Ejecuta primero ACTUALIZAR_DATOS.bat
    pause & exit /b 1
)

:: ── Obtener IP local ───────────────────────────────────────────────
echo  Detectando IP de este computador...
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4"') do (
    set IP=%%a
    goto :found
)
:found
set IP=%IP: =%

echo.
echo  ============================================================
echo   SERVIDOR INICIADO
echo.
echo   Este computador:  http://%IP%:8501
echo.
echo   Los demas computadores deben abrir su navegador
echo   y escribir exactamente esa direccion.
echo   (no necesitan Python ni instalar nada)
echo.
echo   Para cerrar el servidor: Ctrl+C en esta ventana
echo   o simplemente cierra esta ventana.
echo  ============================================================
echo.

py -m streamlit run app_pedidos.py ^
    --server.address 0.0.0.0 ^
    --server.port 8501 ^
    --server.headless true ^
    --browser.gatherUsageStats false

if errorlevel 1 (
    echo.
    echo  [ERROR] El servidor cerro inesperadamente.
    pause
)
