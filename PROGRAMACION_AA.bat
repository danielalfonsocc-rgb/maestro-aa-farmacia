@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo  Programacion AA — Ciclo Bodega AA vs Reporte SSASUR
echo ============================================================
echo.
echo  Uso normal: al reiniciar el ciclo Bodega AA - Bodega Farmacos
echo  (cada 2 semanas), genera la planilla para imprimir y contar.
echo.
echo  Antes: descarga en SSASUR - Reportes - Reporte de consumo por
echo  centro de costo - Centro de Costo = FARMACIA - Generar XLS
echo.
echo  Opciones:
echo    --forzar              generar aunque hoy no sea inicio de ciclo
echo    --aplicar-conteo J    aplicar el conteo (JSON) tras escanear
echo.

py programacion_aa.py %*

echo.
pause
