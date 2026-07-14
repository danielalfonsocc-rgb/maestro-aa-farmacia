@echo off
title Maestro AA - Centro de Operaciones - Farmacia AT Abierta
cd /d "%~dp0"

echo  ============================================================
echo   MAESTRO AA - Centro de Operaciones
echo   Pedidos AA . Pedidos Fusionados . Centinela . Recetas Cheque . GT
echo  ============================================================
echo.

:: -- Python ----------------------------------------------------------
py --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado.
    echo  Ejecuta INSTALAR.bat primero.
    pause & exit /b 1
)

:: -- Dependencias ------------------------------------------------------
py -c "import streamlit, openpyxl, reportlab, rapidfuzz, pandas" >nul 2>&1
if errorlevel 1 (
    echo  Instalando dependencias...
    py -m pip install -r requirements.txt --quiet
)

:: -- Archivo maestro -----------------------------------------------------
if not exist "Consolidado_AA_MAESTRO.xlsx" (
    echo  [AVISO] No se encuentra Consolidado_AA_MAESTRO.xlsx
    echo  El modulo Pedidos AA no cargara hasta ejecutar ACTUALIZAR_DATOS.bat
    echo.
)

echo  Iniciando app...
echo.
echo  Se abrira en el navegador en unos segundos.
echo  URL: http://localhost:8502
echo.
echo  Para cerrar: Ctrl+C en esta ventana
echo  ============================================================
echo.

start "" cmd /c "ping -n 5 127.0.0.1 >nul && start http://localhost:8502"

py -m streamlit run app_maestro.py --server.headless true --browser.gatherUsageStats false --server.port 8502

if errorlevel 1 (
    echo.
    echo  [ERROR] La app cerro inesperadamente.
    pause
)
