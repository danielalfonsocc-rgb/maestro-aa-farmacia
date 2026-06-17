@echo off
title App Pedidos AA - Farmacia AT Abierta
cd /d "%~dp0"

echo  ============================================================
echo   APP PEDIDOS AA - Farmacia AT Abierta
echo  ============================================================
echo.

:: ── Python ────────────────────────────────────────────────────────
py --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado.
    echo  Ejecuta INSTALAR.bat primero.
    pause & exit /b 1
)

:: ── Dependencias ─────────────────────────────────────────────────
py -c "import streamlit, openpyxl, reportlab, rapidfuzz" >nul 2>&1
if errorlevel 1 (
    echo  Instalando dependencias...
    py -m pip install -r requirements.txt --quiet
)

:: ── Archivo maestro ───────────────────────────────────────────────
if not exist "Consolidado_AA_MAESTRO.xlsx" (
    echo  [ERROR] No se encuentra Consolidado_AA_MAESTRO.xlsx
    echo  Ejecuta primero ACTUALIZAR_DATOS.bat
    pause & exit /b 1
)

echo  Iniciando app...
echo.
echo  Se abrira en el navegador en unos segundos.
echo  URL: http://localhost:8501
echo.
echo  Para cerrar: Ctrl+C en esta ventana
echo  ============================================================
echo.

start "" cmd /c "ping -n 5 127.0.0.1 >nul && start http://localhost:8501"

py -m streamlit run app_pedidos.py --server.headless true --browser.gatherUsageStats false --server.port 8501

if errorlevel 1 (
    echo.
    echo  [ERROR] La app cerro inesperadamente.
    pause
)
