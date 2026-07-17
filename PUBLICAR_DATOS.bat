@echo off
title Publicar Datos - Maestro AA
cd /d "%~dp0"

if not exist ".git" goto :nogit

for /f "delims=" %%i in ('powershell -NoProfile -Command "Get-Date -Format 'dd-MM-yyyy HH:mm'"') do set FECHA_HORA=%%i
set MSG=Datos %FECHA_HORA%

echo   Subiendo datos a GitHub...
git add Consolidado_AA_MAESTRO.xlsx Resumen_Pedidos_AA.xlsx 2>nul
git add app_pedidos.py maestro_aa.py aa_colors.py AUTO_SSASUR.py auditoria_prescripcion.py 2>nul
git add requirements.txt .gitignore 2>nul
git add CONFIGURAR_GITHUB.bat PUBLICAR_DATOS.bat INSTALAR_TAREA_PROGRAMADA.bat TAREA_PROGRAMADA_AUTO.xml 2>nul
git add feedback.json auditoria_prescripcion.json 2>nul

git diff --cached --quiet
if errorlevel 1 (
    git commit -m "%MSG%"
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