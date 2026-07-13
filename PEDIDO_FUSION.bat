@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo  Pedido Fusion AA — SGLI + Pedido del Dia
echo ============================================================
echo.
echo  Genera: Pedido_Fusion_AA_<fecha>.xlsx
echo  Hoja 1 Farm_Bod    : Farmacia AA - Bodega AA (ajustado al dia semana)
echo  Hoja 2 Bod_Farmacos: Bodega AA - Bodega Farmacos (ciclo 2 semanas)
echo  Hoja 3 Dialisis    : dialisis mensual (solo S3, o con --forzar-dialisis)
echo.
echo  Opciones:
echo    --forzar-dialisis  incluir hoja dialisis aunque no sea S3
echo    --todos            incluir meds sin necesidad de pedido hoy
echo.

py pedido_fusion.py %*

echo.
pause
