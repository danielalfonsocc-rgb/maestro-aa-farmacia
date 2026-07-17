@echo off
title Sincronizar Todo - Maestro AA
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

set NOGIT=
set NORCH=
set NOPAUSE=
for %%A in (%*) do (
    if /i "%%A"=="--no-git"   set NOGIT=1
    if /i "%%A"=="--no-rch"   set NORCH=1
    if /i "%%A"=="--no-pause" set NOPAUSE=1
)

echo.
echo  ============================================================
echo   SINCRONIZAR TODO  -  Maestro AA Farmacia
echo  ============================================================
echo   Publica lo YA generado (Escritorio + GitHub + Drive), SIN
echo   descargar de SSASUR ni recalcular (excepto el registro ISP
echo   de Recetas Cheque, que se re-chequea contra la sabana ya
echo   descargada por exigencia de la autoridad sanitaria). Para
echo   actualizar datos desde SSASUR usa AUTO_SSASUR en vez de esto.
echo  ============================================================
echo.

echo  [1/4] Copiando resultados al Escritorio...
py publicar_escritorio.py

echo.
if defined NOGIT (
    echo  [2/4] --no-git: omito la publicacion en GitHub.
) else if not exist ".git" (
    echo  [2/4] Git no configurado - ejecuta CONFIGURAR_GITHUB.bat primero.
) else (
    echo  [2/4] Publicando en GitHub...
    call "%~dp0PUBLICAR_DATOS.bat"
)

echo.
if not exist "token_drive.json" (
    echo  [3/4] Google Drive: sin token - ejecuta 'py publicar_drive.py --setup' para activar.
) else (
    echo  [3/4] Subiendo a Google Drive...
    py publicar_drive.py
)

echo.
if defined NORCH (
    echo  [4/4] --no-rch: omito Recetas Cheque ISP.
) else if not exist "token_drive.json" (
    echo  [4/4] Recetas Cheque ISP: sin token de Drive - omitido.
) else (
    echo  [4/4] Registro ISP Recetas Cheque - actualizando con la sabana ya descargada...
    py recetas_cheque.py --no-pause
    echo        Subiendo Recetas Cheque ISP a Drive - CONFIDENCIAL, RUT pacientes...
    echo        Excepcion autorizada por el usuario 2026-06-30, confirmada
    echo        automatica 2026-07-15 - ver subir_recetas_cheque_drive.py
    py subir_recetas_cheque_drive.py
)

echo.
echo  ============================================================
echo   Completado.
echo  ============================================================
echo.
if not defined NOPAUSE pause
