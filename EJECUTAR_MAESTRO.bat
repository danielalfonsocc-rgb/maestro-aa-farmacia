@echo off
title Maestro AA - Consolidado Operacional
cd /d "%~dp0"

echo.
echo  ============================================================
echo   MAESTRO AA - Consolidado Farmacia AT Abierta
echo   Universo activo: 378 medicamentos AA
echo   (Ultima actualizacion exclusiones: 28/05/2026)
echo  ============================================================
echo.
echo  Antes de continuar:
echo.
echo    1. Copia el reporte_de_stock_*.xlsx mas reciente a esta carpeta
echo.
echo    2. Copia los archivos informe_completo_recetas*.csv
echo       Puedes incluir historicos de meses anteriores:
echo       mas archivos = mejor estadistica de consumo
echo       El periodo se detecta automaticamente desde la
echo       primera hasta la ultima fecha de entrega real.
echo.
echo    3. Cierra el Excel si esta abierto
echo.
echo    4. Presiona cualquier tecla para continuar...
echo.
pause >nul

echo  Ejecutando...
echo.

py maestro_aa.py

echo.
if %errorlevel% == 0 (
    echo  ============================================================
    echo   LISTO - Se generaron 2 archivos:
    echo.
    echo   1. Consolidado_AA_MAESTRO.xlsx
    echo      Analisis completo (14 hojas)
    echo.
    echo   2. Resumen_Pedidos_AA.xlsx
    echo      Faltantes + Pedidos Farmacia y Bodega AA
    echo      ajustados por tendencia semanal
    echo  ============================================================
) else (
    echo  ============================================================
    echo   ERROR - Algo fallo. Revisa el mensaje de arriba.
    echo  ============================================================
)
echo.
pause