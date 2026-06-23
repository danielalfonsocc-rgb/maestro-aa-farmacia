@echo off
title Setup Google Drive - Farmacia AA
cd /d "%~dp0"
chcp 65001 >nul

echo.
echo  ============================================================
echo   SETUP GOOGLE DRIVE — Farmacia AA
echo  ============================================================
echo.
echo  Instala dependencias y abre el navegador para autorizar
echo  el acceso a Google Drive (solo la primera vez).
echo.
echo  Prerequisitos:
echo    1. Descarga credentials.json desde Google Cloud Console:
echo       - console.cloud.google.com → Credenciales
echo       - Tipo: OAuth 2.0 Client ID (Desktop App)
echo    2. Guarda el archivo como credentials.json en esta carpeta.
echo.
echo  Presiona cualquier tecla cuando tengas credentials.json listo...
pause >nul

if not exist credentials.json (
    echo.
    echo  [ERROR] No se encontro credentials.json en esta carpeta.
    echo  Descargalo desde console.cloud.google.com
    pause
    exit /b 1
)

echo.
echo  Instalando dependencias de Google Drive API...
py -m pip install google-api-python-client google-auth-oauthlib google-auth-httplib2 --quiet
if errorlevel 1 (
    echo  [ERROR] No se pudieron instalar las dependencias.
    pause
    exit /b 1
)

echo.
echo  Abriendo navegador para autorizar acceso a Drive...
echo  (se abrira una ventana de Google — acepta los permisos)
echo.
py publicar_drive.py --setup

if errorlevel 1 (
    echo.
    echo  [ERROR] Fallo la autorizacion. Revisa credentials.json.
    pause
    exit /b 1
)

echo.
echo  ============================================================
echo   Listo! token_drive.json guardado.
echo   Desde ahora AUTO_SSASUR sube automaticamente a Drive.
echo  ============================================================
echo.
pause
