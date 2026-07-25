@echo off
title Conteo de Controlados - Servidor Online (varias computadoras)
cd /d "%~dp0"

echo  ============================================================
echo   CONTEO DE CONTROLADOS - SERVIDOR ONLINE
echo   Varias computadoras llenando el MISMO conteo a la vez
echo  ============================================================
echo.

:: -- Python ----------------------------------------------------------
py --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado. Ejecuta INSTALAR.bat primero.
    pause & exit /b 1
)

py servidor_conteo.py --port 8504

if errorlevel 1 (
    echo.
    echo  [ERROR] El servidor cerro inesperadamente.
    pause
)
