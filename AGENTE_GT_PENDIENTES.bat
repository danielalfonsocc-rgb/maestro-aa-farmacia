@echo off
title Agente GT Pendientes - Farmacia Hospital de Pitrufquen
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

echo.
echo  ============================================================
echo   AGENTE GT PENDIENTES
echo   Clasifica medicamentos pendientes por urgencia (Haiku IA)
echo   Requiere: gt_enriquecido.json en out_gt\
echo  ============================================================
echo.

if not exist "out_gt\gt_enriquecido.json" (
    echo  [ERROR] No se encontro out_gt\gt_enriquecido.json
    echo.
    echo  Ejecuta primero GT.bat para descargar y procesar el reporte.
    echo.
    pause & exit /b 1
)

py agente_gt_pendientes.py

echo.
if %errorlevel% == 0 (
    echo  ============================================================
    echo   LISTO: out_gt\Pendientes_GT_Prioridad.xlsx
    echo  ============================================================
) else (
    echo  [AVISO] Revisa los mensajes arriba.
)
echo.
pause
