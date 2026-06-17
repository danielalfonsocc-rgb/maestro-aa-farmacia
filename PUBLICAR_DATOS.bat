@echo off
title Publicar Datos - Maestro AA
cd /d "%~dp0"

:: ── Verificar que estamos en un repo Git ─────────────────────────────────────
if not exist ".git" (
    echo  [AVISO] Git no esta configurado en esta carpeta.
    echo  Ejecuta primero: CONFIGURAR_GITHUB.bat
    exit /b 0
)

:: ── Fecha y hora para el mensaje de commit ───────────────────────────────────
for /f "tokens=1-3 delims=/ " %%a in ("%DATE%") do (
    set DIA=%%a
    set MES=%%b
    set ANO=%%c
)
for /f "tokens=1-2 delims=:, " %%a in ("%TIME%") do (
    set HORA=%%a
    set MIN=%%b
)

set MSG=Datos %DIA%/%MES%/%ANO% %HORA%:%MIN%

:: ── Subir solo los archivos de datos y la app ────────────────────────────────
echo   Subiendo datos a GitHub...
git add Consolidado_AA_MAESTRO.xlsx Resumen_Pedidos_AA.xlsx 2>nul
git add app_pedidos.py maestro_aa.py aa_colors.py AUTO_SSASUR.py 2>nul
git add requirements.txt *.bat .gitignore 2>nul
git add feedback.json 2>nul

git diff --cached --quiet
if errorlevel 1 (
    git commit -m "%MSG%"
    git push
    if errorlevel 1 (
        echo   [AVISO] Push fallo. Puede que necesites reconectar Git a GitHub.
        echo          Ejecuta CONFIGURAR_GITHUB.bat para reconfigurar.
    ) else (
        echo   ✓ Datos publicados: %MSG%
    )
) else (
    echo   (Sin cambios que publicar)
)
