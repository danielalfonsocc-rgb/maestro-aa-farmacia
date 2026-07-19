@echo off
chcp 65001 >nul
title Auditoría de Cantidad Recetada vs. Posología

echo ============================================================
echo   AUDITORÍA DE CANTIDAD RECETADA VS. POSOLOGÍA
echo   Farmacia AT Abierta - Hospital de Pitrufquén
echo ============================================================
echo.
echo  Opciones:
echo    1. Mes actual                    (recomendado)
echo    2. Elegir otro mes (AAAA-MM)
echo    3. Salir
echo.
set /p OPC="  Elige [1-3]: "

if "%OPC%"=="1" goto ACTUAL
if "%OPC%"=="2" goto OTRO_MES
if "%OPC%"=="3" goto FIN
goto ACTUAL

:ACTUAL
echo.
echo  [*] Ejecutando auditoria del mes actual...
py auditoria_cantidad_posologia.py
goto FIN

:OTRO_MES
echo.
set /p MES="  Mes a analizar (formato AAAA-MM, ej. 2026-06): "
echo.
echo  [*] Ejecutando auditoria de %MES%...
py auditoria_cantidad_posologia.py --mes %MES%
goto FIN

:FIN
echo.
pause
