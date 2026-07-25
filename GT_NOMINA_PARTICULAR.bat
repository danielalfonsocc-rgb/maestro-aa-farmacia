@echo off
title Gestion Territorial - Nomina particular - Farmacia Hospital de Pitrufquen
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

:MENU
cls
echo.
echo  ============================================================
echo   GESTION TERRITORIAL - Generar Nomina de Envio particular
echo  ============================================================
echo.
echo   Para recetas que NO salen con el flujo automatico de
echo   AUTO_SSASUR.bat / GT.bat (sin destino en el sistema, o que
echo   se quedaron atras en la corrida automatica).
echo.
echo   1. Registrar recetas a mano (yo escribo los datos)
echo   2. Leer un Excel de Modalidad de Despacho ya descargado
echo      (detecta solas las que no alcanzaron a salir)
echo   3. Salir
echo.
set /p OPCION="Elige una opcion (1/2/3): "

if "%OPCION%"=="1" goto MANUAL
if "%OPCION%"=="2" goto DESDE_EXCEL
if "%OPCION%"=="3" exit /b 0
echo  Opcion invalida.
pause
goto MENU

:MANUAL
py --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado. Ejecuta INSTALAR.bat primero.
    pause & exit /b 1
)

if not exist "_gt_manual_plantilla.csv" (
    echo  [ERROR] No encuentro _gt_manual_plantilla.csv en esta carpeta.
    pause & exit /b 1
)

for /f %%F in ('py -c "import datetime; print(datetime.date.today().isoformat())"') do set FECHA_HOY=%%F
set LOTE=_gt_manual_lote_%FECHA_HOY%.csv
if not exist "%LOTE%" copy "_gt_manual_plantilla.csv" "%LOTE%" >nul

echo.
echo  Se va a abrir %LOTE% en Excel.
echo  Llena una fila por receta: receta,paciente,rut,destino,periodo,especialidad,n_presc,telefono,pendiente
echo  Guarda el archivo y CIERRA Excel antes de continuar.
echo.
pause
start "" "%LOTE%"
echo.
echo  Cuando hayas guardado y cerrado el Excel, presiona una tecla para continuar...
pause >nul

echo.
echo  Vista previa (no se guarda nada todavia):
echo  ------------------------------------------------------------
py agregar_gt_manual.py --csv "%LOTE%" --dry-run
echo  ------------------------------------------------------------
echo.
set /p CONFIRMA="Se ve bien? Aplicar y generar la(s) Nomina(s) de Envio? (S/N): "
if /i not "%CONFIRMA%"=="S" (
    echo  Cancelado. El archivo %LOTE% queda guardado por si quieres reintentar.
    pause
    goto MENU
)
py agregar_gt_manual.py --csv "%LOTE%"
echo.
echo  Subiendo a Google Drive...
py publicar_drive.py --solo-gt
echo.
echo  ============================================================
echo   LISTO
echo  ============================================================
pause
goto MENU

:DESDE_EXCEL
py --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado. Ejecuta INSTALAR.bat primero.
    pause & exit /b 1
)

echo.
echo  Arrastra aqui el archivo reporteGestionTerritorial_*.xlsx
echo  (el Informe Modalidad de Despacho que baja AUTO_SSASUR.bat, en
echo  ..\04_Farmacia_Gestion_Territorial\) y presiona Enter:
echo.
set /p RUTA_EXCEL="Ruta del Excel: "
set RUTA_EXCEL=%RUTA_EXCEL:"=%

if not exist "%RUTA_EXCEL%" (
    echo  [ERROR] No encuentro ese archivo.
    pause
    goto MENU
)

echo.
echo  Revisando cuales recetas del informe todavia NO estan en el maestro...
echo  ------------------------------------------------------------
py agregar_gt_manual.py --gt-excel "%RUTA_EXCEL%" --dry-run
echo  ------------------------------------------------------------
echo.
set /p CONFIRMA="Se ve bien? Aplicar y generar la(s) Nomina(s) de Envio? (S/N): "
if /i not "%CONFIRMA%"=="S" (
    echo  Cancelado.
    pause
    goto MENU
)
py agregar_gt_manual.py --gt-excel "%RUTA_EXCEL%"
echo.
echo  Subiendo a Google Drive...
py publicar_drive.py --solo-gt
echo.
echo  ============================================================
echo   LISTO
echo  ============================================================
pause
goto MENU
