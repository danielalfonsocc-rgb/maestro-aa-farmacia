@echo off
title Instalar Tarea Programada - Maestro AA
cd /d "%~dp0"
color 0A

echo.
echo  ================================================================
echo   INSTALAR TAREA PROGRAMADA  -  Maestro AA Farmacia
echo  ================================================================
echo.
echo  Esto instalara una tarea que abre AUTO_SSASUR.bat
echo  automaticamente de lunes a viernes a las 07:30.
echo.
echo  Recuerda: TU debes logarte en SSASUR cuando se abra el navegador.
echo  El resto corre solo.
echo.
echo  Necesitas ejecutar este archivo como Administrador.
echo.
pause

:: Importar la tarea desde el XML
schtasks /Create /XML "%~dp0TAREA_PROGRAMADA_AUTO.xml" /TN "MaestroAA_AutoSSASUR" /F

if errorlevel 1 (
    echo.
    echo  [ERROR] No se pudo instalar la tarea.
    echo  Asegurate de ejecutar este archivo como Administrador:
    echo  Click derecho ^> "Ejecutar como administrador"
) else (
    echo.
    echo  ================================================================
    echo   EXITO - Tarea instalada correctamente
    echo.
    echo   Nombre de la tarea : MaestroAA_AutoSSASUR
    echo   Horario            : Lunes a Viernes, 07:30
    echo.
    echo   Para modificar la hora:
    echo   Inicio ^> Programador de tareas ^> MaestroAA_AutoSSASUR ^> Propiedades
    echo  ================================================================
)

echo.
pause
