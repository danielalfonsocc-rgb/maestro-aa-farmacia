@echo off
setlocal enabledelayedexpansion
title Publicar Datos - Maestro AA
cd /d "%~dp0"

if not exist ".git" goto :nogit

for /f "delims=" %%i in ('powershell -NoProfile -Command "Get-Date -Format 'dd-MM-yyyy HH:mm'"') do set FECHA_HORA=%%i
set MSG=Datos %FECHA_HORA%

rem Lista fija de archivos que este script publica automaticamente. Se
rem comitea/pushea SOLO esta lista, via pathspec en "git commit -- <lista>"
rem -- nunca "todo lo que haya en staging". Si hay otro cambio de codigo a
rem medio comitear en el indice (ej. una sesion de Claude preparando un
rem commit manual con mensaje descriptivo), se deja intacto en staging para
rem ese commit, en vez de tragarselo con el mensaje generico "Datos ...".
rem (Se solto asi el 21-07-2026: un commit manual en curso quedo con este
rem mensaje generico por la condicion de carrera. Ver memoria del proyecto.)
set ARCHIVOS=
for %%F in (
    Consolidado_AA_MAESTRO.xlsx Resumen_Pedidos_AA.xlsx
    app_pedidos.py maestro_aa.py aa_colors.py AUTO_SSASUR.py auditoria_prescripcion.py
    requirements.txt .gitignore
    CONFIGURAR_GITHUB.bat PUBLICAR_DATOS.bat INSTALAR_TAREA_PROGRAMADA.bat TAREA_PROGRAMADA_AUTO.xml
    feedback.json auditoria_prescripcion.json
) do (
    if exist "%%F" set ARCHIVOS=!ARCHIVOS! "%%F"
)

if "!ARCHIVOS!"=="" (
    echo   Sin archivos de la lista presentes en el repo.
    goto :eof
)

echo   Subiendo datos a GitHub...
git add -- !ARCHIVOS! 2>nul

git diff --cached --quiet -- !ARCHIVOS!
if errorlevel 1 (
    git commit -m "%MSG%" -- !ARCHIVOS!
    git push
    if errorlevel 1 (
        echo   [AVISO] Push fallo. Ejecuta CONFIGURAR_GITHUB.bat para reconfigurar.
    ) else (
        echo   OK Datos publicados: %MSG%
    )
) else (
    echo   Sin cambios que publicar.
)
goto :eof

:nogit
echo   [AVISO] Git no configurado. Ejecuta CONFIGURAR_GITHUB.bat primero.