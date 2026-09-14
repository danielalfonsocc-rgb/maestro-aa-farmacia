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
echo  El reporte de consumo por centro de costo (FARMACIA) lo baja
echo  AUTO_SSASUR todos los dias; corre AUTO_SSASUR antes si hoy no corrio.
echo.
echo  Opciones:
echo    --forzar              generar aunque hoy no sea inicio de ciclo
echo    --aplicar-conteo J    aplicar el conteo (JSON) tras escanear
echo.

py programacion_aa.py %*

echo.
pause
