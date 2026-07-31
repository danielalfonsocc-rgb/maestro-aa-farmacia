@echo off
title Clozapina - Consolidado Hemogramas MINSAL
cd /d "%~dp0"
rem Corre AUTO_SSASUR completo (recetas+stock+GT) MAS el paso de Clozapina
rem (HCE: busca el hemograma mas reciente de cada paciente y arma el Excel
rem listo para "Ingreso de Hemogramas en Bloque" del MINSAL).
call AUTO_SSASUR.bat --clozapina
