@echo off
title Registro ISP Recetas Cheque - Farmacia AT Abierta
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

echo.
echo  ============================================================
echo   REGISTRO ISP RECETAS CHEQUE  -  Farmacia AT Abierta
echo   Agrega los folios cheque NUEVOS al formulario ISP del mes
echo   (usa la ultima sabana descargada; no toca filas existentes)
echo  ============================================================
echo.

:: -- Python ----------------------------------------------------------
py --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado. Ejecuta INSTALAR.bat primero.
    pause & exit /b 1
)

:: -- Dependencias ----------------------------------------------------
py -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo  Instalando dependencias...
    py -m pip install pandas openpyxl --quiet
)

py recetas_cheque.py

echo.
echo  Copiando formulario al Escritorio...
py publicar_escritorio.py --rch

echo.
pause
