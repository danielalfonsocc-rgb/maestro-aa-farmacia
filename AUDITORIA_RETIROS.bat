@echo off
title Auditoria de Retiros por Medicamento - Farmacia Hospital Pitrufquen
cd /d "%~dp0"

echo  ============================================================
echo   AUDITORIA DE RETIROS POR MEDICAMENTO
echo   Pacientes distintos y N de retiros en un periodo dado
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

echo  Iniciando app de auditoria...
echo.
echo  Se abrira en el navegador en unos segundos.
echo  URL: http://localhost:8505
echo.
echo  La primera busqueda carga TODO el historico de recetas disponible
echo  (repo + carpetas de respaldo) y puede tardar ~1 minuto.
echo  Para cerrar: Ctrl+C en esta ventana
echo  ============================================================
echo.

start "" cmd /c "ping -n 5 127.0.0.1 >nul && start http://localhost:8505"

py -m streamlit run app_auditoria_retiros.py --server.headless true --browser.gatherUsageStats false --server.port 8505

if errorlevel 1 (
    echo.
    echo  [ERROR] La app cerro inesperadamente.
    pause
)
