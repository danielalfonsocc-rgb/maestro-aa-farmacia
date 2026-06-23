@echo off
title Maestro AA - Consolidado Operacional
cd /d "%~dp0"

echo  ============================================================
echo   MAESTRO AA - Farmacia AT Abierta
echo  ============================================================
echo.

:: ── Python ────────────────────────────────────────────────────────
py --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado.
    echo  Ejecuta INSTALAR.bat primero.
    pause & exit /b 1
)

:: ── Dependencias minimas ──────────────────────────────────────────
py -c "import pandas, openpyxl, numpy" >nul 2>&1
if errorlevel 1 (
    echo  Instalando dependencias...
    py -m pip install -r requirements.txt --quiet
)

:: ── Archivos de entrada ───────────────────────────────────────────
if not exist "reporte_de_stock_*.xlsx" (
    echo  [AVISO] No se encontro reporte_de_stock_*.xlsx
    echo  Descarga el stock desde SSASUR y copialo en esta carpeta.
    echo.
)
if not exist "informe_completo_recetas*.csv" (
    echo  [AVISO] No se encontro informe_completo_recetas*.csv
    echo  Descarga las recetas desde SSASUR y copialas en esta carpeta.
    echo.
)

echo  Ejecutando Maestro...
echo.
py maestro_aa.py

if %errorlevel% == 0 (
    echo.
    echo  Copiando resultados al Escritorio...
    py publicar_escritorio.py --app
    echo.
    echo  ============================================================
    echo   LISTO
    echo   - Consolidado_AA_MAESTRO.xlsx
    echo   - Resumen_Pedidos_AA.xlsx
    echo   - Copia en el Escritorio: Farmacia AA\1 - App Pedidos
    echo  ============================================================
) else (
    echo.
    echo  [ERROR] Revisa los mensajes arriba.
)
echo.
pause
