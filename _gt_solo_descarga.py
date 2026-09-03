"""Descarga UNICA del informe de modalidad de despacho (GT), sin el cruce_gt.py
--generar automatico que dispara `AUTO_SSASUR.py --gt`. Uso puntual pedido por
el usuario (11-08-2026): "descarga el reporte y ya esta", sin generar planillas
ni sincronizar gt_maestro.xlsx. Reusa entrar_receta/paso_gt/sesion guardada del
script principal - no reimplementa nada."""
import asyncio
import sys
from datetime import date, timedelta

from playwright.async_api import async_playwright

import AUTO_SSASUR as A


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
        )
        if A.SESSION_FILE.exists():
            context = await browser.new_context(accept_downloads=True, storage_state=str(A.SESSION_FILE))
            print("\n  (Sesion guardada encontrada - puede que no necesites logarte)")
        else:
            context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        print("\n[1/2] Logeate en SSASUR si hace falta (tienes 5 minutos)...")
        await page.goto(A.DASHBOARD_URL)
        try:
            await page.wait_for_selector(
                'button:has-text("ABASTECIMIENTO"), div:has-text("ABASTECIMIENTO")',
                timeout=A.TIMEOUT_LOGIN,
            )
        except Exception:
            if not await page.evaluate("document.body.innerText.includes('ABASTECIMIENTO')"):
                raise
        await context.storage_state(path=str(A.SESSION_FILE))
        print("  OK sesion detectada")

        today = date.today()
        desde_gt = A.fmt(A.dia_habil_anterior(today))
        hasta_gt = A.fmt(today + timedelta(days=13))

        print("\n[2/2] Descargando informe de modalidad de despacho (GT)...")
        dest, n = await A.paso_gt(page, desde_gt, hasta_gt, debug=False)
        await browser.close()

        print("\n" + "=" * 62)
        if dest:
            print(f"  OK {dest.name}  ({n} recetas)")
            print(f"  Archivo: {dest}")
        elif n == 0:
            print("  Sin recetas en el listado para esas fechas - no se descargo nada.")
        else:
            print("  ERROR: no se genero el Excel (revisa debug_gt.png).")
        print(f"  Carpeta: {A.GT_DIR}")
        print("=" * 62)


if __name__ == "__main__":
    asyncio.run(main())
