@echo off
title Servidor Conteo Controlados - Red Local
cd /d "%~dp0"

echo  ============================================================
echo   SERVIDOR CONTEO DE CONTROLADOS - RED LOCAL
echo   Farmacia Hospital Pitrufquen
echo  ============================================================
echo.

:: -- Python ----------------------------------------------------------
py --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado. Ejecuta INSTALAR.bat primero.
    pause & exit /b 1
)

:: -- Dependencias ----------------------------------------------------
py -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo  Instalando dependencias...
    py -m pip install -r requirements.txt --quiet
)

:: -- Obtener IP local ------------------------------------------------
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
echo   Este computador:  http://%IP%:8503
echo.
echo   Los demas computadores de la MISMA red del hospital
echo   deben abrir su navegador y escribir exactamente esa
echo   direccion. (no necesitan Python ni instalar nada)
echo.
echo   Nota: cada computador guarda su propio historial. Para
echo   juntar conteos usa "Exportar JSON" y la pestana Consolidado.
echo.
echo   Para cerrar el servidor: Ctrl+C o cierra esta ventana.
echo  ============================================================
echo.

py -m streamlit run conteo_controlados_app.py ^
    --server.address 0.0.0.0 ^
    --server.port 8503 ^
    --server.headless true ^
    --browser.gatherUsageStats false

if errorlevel 1 (
    echo.
    echo  [ERROR] El servidor cerro inesperadamente.
    pause
)
