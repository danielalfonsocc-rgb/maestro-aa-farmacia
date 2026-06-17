@echo off
title Publicar Datos - Maestro AA
cd /d "%~dp0"

if not exist ".git" goto :nogit

for /f "tokens=1-3 delims=/ " %%a in ("%DATE%") do (
    set DIA=%%a
    set MES=%%b
    set ANO=%%c
)
for /f "tokens=1-2 delims=: " %%a in ("%TIME%") do (
    set HORA=%%a
    set MIN=%%b
)
set MSG=Datos %DIA%/%MES%/%ANO% %HORA%:%MIN%

echo   Subiendo datos a GitHub...
git add Consolidado_AA_MAESTRO.xlsx Resumen_Pedidos_AA.xlsx 2>nul
git add app_pedidos.py maestro_aa.py aa_colors.py AUTO_SSASUR.py 2>nul
git add requirements.txt .gitignore 2>nul
git add CONFIGURAR_GITHUB.bat PUBLICAR_DATOS.bat INSTALAR_TAREA_PROGRAMADA.bat TAREA_PROGRAMADA_AUTO.xml 2>nul
git add feedback.json 2>nul

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