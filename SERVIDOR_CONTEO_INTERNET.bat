@echo off
title Conteo de Controlados - Servidor + Internet (Cloudflare Tunnel)
cd /d "%~dp0"

echo  ============================================================
echo   CONTEO DE CONTROLADOS - SERVIDOR + INTERNET
echo   Accesible desde cualquier red (no solo el hospital)
echo  ============================================================
echo.

:: -- Python ----------------------------------------------------------
py --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado. Ejecuta INSTALAR.bat primero.
    pause & exit /b 1
)

:: -- cloudflared -------------------------------------------------------
where cloudflared >nul 2>&1
if errorlevel 1 (
    if not exist "C:\Program Files (x86)\cloudflared\cloudflared.exe" (
        echo  Instalando cloudflared (una sola vez)...
        winget install --id Cloudflare.cloudflared -e --accept-source-agreements --accept-package-agreements
    )
)

py servidor_conteo_internet.py

if errorlevel 1 (
    echo.
    echo  [ERROR] El servidor cerro inesperadamente.
    pause
)
