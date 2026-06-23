@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo  Agente IA — Duplicados del dia (Farmacia AT Abierta)
echo ============================================================
echo.

:: Fecha de hoy por defecto.  Para otra fecha: --fecha 2026-06-20
:: Para cambiar la ventana: --ventana 60
:: Para modelo mas potente: --modelo claude-sonnet-4-6

py agente_duplicados.py --ventana 90 %*

echo.
pause
