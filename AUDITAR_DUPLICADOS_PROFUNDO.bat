@echo off
chcp 65001 >nul
title Auditoría Profunda de Prescripciones Duplicadas

echo ============================================================
echo   AUDITORÍA PROFUNDA DE PRESCRIPCIONES DUPLICADAS
echo   Farmacia AT Abierta - Hospital de Pitrufquén
echo ============================================================
echo.
echo  Opciones:
echo    1. Auditoria COMPLETA con propuestas IA  (recomendado)
echo    2. Auditoria rapida SIN IA               (sin llamadas a Claude)
echo    3. Auditoria rapida CON IA               (solo CSV mas reciente)
echo    4. Salir
echo.
set /p OPC="  Elige [1-4]: "

if "%OPC%"=="1" goto COMPLETO
if "%OPC%"=="2" goto SIN_IA
if "%OPC%"=="3" goto RAPIDO
if "%OPC%"=="4" goto FIN
goto COMPLETO

:COMPLETO
echo.
echo  [*] Ejecutando auditoria completa con propuestas IA...
py auditoria_duplicados_profunda.py
goto FIN

:SIN_IA
echo.
echo  [*] Ejecutando auditoria sin llamadas IA...
py auditoria_duplicados_profunda.py --sin-ia
goto FIN

:RAPIDO
echo.
echo  [*] Ejecutando auditoria rapida (ultimo CSV) con IA...
py auditoria_duplicados_profunda.py --rapido
goto FIN

:FIN
echo.
pause
