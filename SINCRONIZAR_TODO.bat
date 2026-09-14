@echo off
title Sincronizar Todo - Maestro AA
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

set NORCH=
set NOPAUSE=
for %%A in (%*) do (
    if /i "%%A"=="--no-rch"   set NORCH=1
    if /i "%%A"=="--no-pause" set NOPAUSE=1
)

echo.
echo  ============================================================
echo   SINCRONIZAR TODO  -  Maestro AA Farmacia
echo  ============================================================
echo   Publica lo YA generado (Escritorio + Drive), SIN descargar
echo   de SSASUR ni recalcular (excepto el registro ISP de Recetas
echo   Cheque, que se re-chequea contra la sabana ya descargada por
echo   exigencia de la autoridad sanitaria). Para actualizar datos
echo   desde SSASUR usa AUTO_SSASUR en vez de esto.
echo  ============================================================
echo.

rem El paso "Publicando en GitHub" (PUBLICAR_DATOS.bat) se elimino el
rem 14-09-2026: hacia un commit diario "Datos ..." con el Consolidado y el
rem Resumen para la app Streamlit, que ya no existe. El codigo se sube a
rem GitHub con commits normales.

echo  [1/3] Copiando resultados al Escritorio...
py publicar_escritorio.py

echo.
if not exist "token_drive.json" (
    echo  [2/3] Google Drive: sin token - ejecuta 'py publicar_drive.py --setup' para activar.
) else (
    echo  [2/3] Subiendo a Google Drive...
    py publicar_drive.py
)

echo.
rem recetas_cheque.py corre aca UNA sola vez por corrida (antes AUTO_SSASUR
rem tambien lo corria como PASO 5d, justo antes). Corre aunque no haya token
rem de Drive: el registro ISP es obligacion legal y es local; solo la subida
rem a Drive necesita el token.
if defined NORCH (
    echo  [3/3] --no-rch: omito Recetas Cheque ISP.
) else (
    echo  [3/3] Registro ISP Recetas Cheque - actualizando con la sabana ya descargada...
    py recetas_cheque.py --no-pause
    if not exist "token_drive.json" (
        echo        Sin token de Drive - no se sube a Drive.
    ) else (
        echo        Subiendo Recetas Cheque ISP a Drive - CONFIDENCIAL, RUT pacientes...
        echo        Excepcion autorizada por el usuario 2026-06-30, confirmada
        echo        automatica 2026-07-15 - ver subir_recetas_cheque_drive.py
        py subir_recetas_cheque_drive.py
    )
)

echo.
echo  ============================================================
echo   Completado.
echo  ============================================================
echo.
if not defined NOPAUSE pause
