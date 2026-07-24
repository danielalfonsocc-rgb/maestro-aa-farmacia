@echo off
title Conteo de Controlados - Farmacia Hospital Pitrufquen
cd /d "%~dp0"

echo  ============================================================
echo   CONTEO DE CONTROLADOS
echo   Pauta de supervision de stock - AT Cerrada . AT Abierta . Urgencia
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
py -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo  Instalando dependencias...
    py -m pip install -r requirements.txt --quiet
)

echo  Iniciando app de conteo...
echo.
echo  Se abrira en el navegador en unos segundos.
echo  URL: http://localhost:8503
echo.
echo  El stock del dia anterior se carga solo tras correr AUTO_SSASUR.
echo  Para cerrar: Ctrl+C en esta ventana
echo  ============================================================
echo.

start "" cmd /c "ping -n 5 127.0.0.1 >nul && start http://localhost:8503"

py -m streamlit run conteo_controlados_app.py --server.headless true --browser.gatherUsageStats false --server.port 8503

if errorlevel 1 (
    echo.
    echo  [ERROR] La app cerro inesperadamente.
    pause
)
