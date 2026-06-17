@echo off
title Configurar GitHub - Maestro AA
cd /d "%~dp0"
color 0A

echo.
echo  ================================================================
echo   CONFIGURACION INICIAL GITHUB  -  Maestro AA Farmacia
echo  ================================================================
echo.
echo  Este script configura Git en esta carpeta por primera vez.
echo  Solo necesitas ejecutarlo UNA vez.
echo.
echo  ANTES de continuar necesitas:
echo    1. Tener una cuenta en https://github.com
echo    2. Crear un repositorio PRIVADO en GitHub
echo       (Ej: "maestro-aa-farmacia")
echo    3. Tener la URL del repo (la copias desde GitHub)
echo.
echo  ================================================================
echo.

:: ── Verificar Git ────────────────────────────────────────────────────────────
git --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Git no esta instalado.
    echo  Descargalo desde: https://git-scm.com/download/win
    echo  Luego cierra esta ventana y vuelve a ejecutar este archivo.
    pause & exit /b 1
)
echo  [OK] Git instalado:
git --version

echo.
set /p GH_URL="  Pega la URL de tu repositorio GitHub (Ej: https://github.com/tuusuario/maestro-aa-farmacia): "
set /p GH_NOMBRE="  Tu nombre completo (para los commits): "
set /p GH_EMAIL="  Tu email de GitHub: "

echo.
echo  Configurando Git...
git config user.name "%GH_NOMBRE%"
git config user.email "%GH_EMAIL%"

:: ── Inicializar repo si no existe ────────────────────────────────────────────
if not exist ".git" (
    git init -b main
    echo  [OK] Repositorio Git inicializado.
) else (
    echo  [OK] Git ya estaba inicializado.
)

:: ── Agregar remote ───────────────────────────────────────────────────────────
git remote remove origin >nul 2>&1
git remote add origin %GH_URL%
echo  [OK] Repositorio remoto configurado: %GH_URL%

:: ── Primer push ──────────────────────────────────────────────────────────────
echo.
echo  Subiendo archivos a GitHub por primera vez...
echo  (se te pedira tu usuario y contrasena/token de GitHub)
echo.
git add .
git commit -m "Configuracion inicial Maestro AA Farmacia"
git push -u origin main

if errorlevel 1 (
    echo.
    echo  [AVISO] Si el push fallo, puede que necesites un Token de GitHub:
    echo    1. Ve a github.com → Settings → Developer settings → Personal access tokens
    echo    2. Genera un token con permiso "repo"
    echo    3. Usa ese token como contrasena cuando Git lo pida
    echo.
) else (
    echo.
    echo  ================================================================
    echo   EXITO - Tu repositorio esta en GitHub
    echo.
    echo   Proximos pasos:
    echo   1. Ve a https://share.streamlit.io
    echo   2. Conecta tu cuenta de GitHub
    echo   3. Selecciona el repo: %GH_URL%
    echo   4. Archivo principal: app_pedidos.py
    echo   5. Haz clic en Deploy
    echo  ================================================================
)

echo.
pause
