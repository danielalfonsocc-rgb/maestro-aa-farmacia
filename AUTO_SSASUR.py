#!/usr/bin/env python3
"""
AUTO_SSASUR.py  —  Maestro AA Farmacia
═══════════════════════════════════════════════════════════════
Descarga automáticamente RECETA y ABASTECIMIENTO desde SSASUR
y actualiza los Excel del Maestro AA.

Flujo:
  1. Se abre Chromium en la página de login de SSASUR
  2. Tú te logeas manualmente (el script no toca tus credenciales)
  3. El script trabaja UN módulo a la vez en UNA sola pestaña
     (SSASUR es una SPA que exige usar una única pestaña):
       a) Clic en la tarjeta RECETA → informe de recetas → descarga
       b) Vuelve al dashboard → clic en ABASTECIMIENTO → stock → descarga
  4. Ejecuta maestro_aa.py para actualizar el Consolidado
  5. Publica los datos en GitHub (si está configurado)

NOTA TÉCNICA — cómo navega SSASUR (verificado por inspección del DOM):
  · Las tarjetas del dashboard son <button> de Vue (no <a>): hay que
    clicarlas por texto, p.ej. button:has-text("ABASTECIMIENTO").
  · El clic NO abre pestaña nueva: navega ESTA misma pestaña al módulo
    (www.ssasur.cl/<modulo>) y acuña la sesión del módulo vía SSO.
  · Ir directo por URL a un módulo sin pasar por el clic da 403 / rebote
    al login, porque la cookie de sesión solo cubre login.ssasur.cl.
  · El reporte de stock NO requiere firma electrónica; basta elegir
    bodega = TODAS (value 0) y pulsar "Generar XLS".
═══════════════════════════════════════════════════════════════
"""
import asyncio
import os
import subprocess
import sys
from pathlib import Path
from datetime import date

# Forzar UTF-8 en la salida — si no, al redirigir a un archivo/pipe Windows usa
# cp1252 y los caracteres ═ ✓ → ✗ revientan el script antes de empezar.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# ── Configuración ─────────────────────────────────────────────────────────────
MAESTRO_DIR      = Path(__file__).parent
SESSION_FILE     = MAESTRO_DIR / ".ssasur_session.json"
DASHBOARD_URL    = "https://login.ssasur.cl/dashboard"
RECETA_INFORME   = "https://www.ssasur.cl/receta/informes/sabana"
STOCK_REPORTE    = "https://www.ssasur.cl/abastecimiento/reportes/stock_en_momento_bodega"
ESTAB_PITRUFQUEN = "59"   # slctHeaderEstab → "PITRUFQUEN HOSP." (única opción)
BODEGA_TODAS     = "0"    # bodega → "TODAS"
TIMEOUT_LOGIN    = 300_000   # 5 minutos para logarse
TIMEOUT_DESCARGA = 600_000   # 10 minutos por descarga


def fmt(d: date) -> str:
    return d.strftime("%d/%m/%Y")


async def entrar_modulo(page, nombre: str):
    """
    Entra a un módulo desde el dashboard. SSASUR es una SPA de UNA pestaña:
    la tarjeta es un <button>; al clicarlo, ESTA misma pestaña navega al
    módulo (www.ssasur.cl/<modulo>) y se acuña su sesión. No abre tab nueva.
    """
    btn = page.locator(f'button:has-text("{nombre}")').first
    await btn.wait_for(state="visible", timeout=30_000)
    await btn.click()
    # Esperar a que la SPA navegue fuera del dashboard (login.ssasur.cl → www.ssasur.cl)
    try:
        await page.wait_for_url(lambda u: "login.ssasur.cl" not in u, timeout=30_000)
    except Exception:
        pass
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(2_000)
    print(f"  ✓ En módulo {nombre}: {page.url}")


async def descargar(page, dest_dir: Path, accion, etiqueta: str, timeout=TIMEOUT_DESCARGA):
    """Envuelve un clic que dispara una descarga y guarda el archivo."""
    async with page.expect_download(timeout=timeout) as dl_info:
        await accion()
    dl   = await dl_info.value
    dest = dest_dir / dl.suggested_filename
    await dl.save_as(dest)
    size_kb = dest.stat().st_size // 1024
    print(f"  ✓ {dl.suggested_filename}  ({size_kb:,} KB)")
    return dest


async def main():
    # ── Verificar Playwright ───────────────────────────────────────────────────
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print("\n[ERROR] Playwright no está instalado.")
        print("  Ejecuta AUTO_SSASUR.bat para instalarlo automáticamente.\n")
        try:
            input("Presiona Enter para cerrar...")
        except EOFError:
            pass
        sys.exit(1)

    today        = date.today()
    fecha_inicio = fmt(today.replace(day=1))
    fecha_fin    = fmt(today)

    print()
    print("═" * 62)
    print("  AUTO SSASUR  ·  Maestro AA Farmacia")
    print(f"  Período: {fecha_inicio}  →  {fecha_fin}")
    print("═" * 62)

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
        )

        if SESSION_FILE.exists():
            context = await browser.new_context(
                accept_downloads=True,
                storage_state=str(SESSION_FILE),
            )
            print("\n  (Sesión guardada encontrada — puede que no necesites logarte)")
        else:
            context = await browser.new_context(accept_downloads=True)

        page = await context.new_page()

        # ── PASO 1 — DETECTAR LOGIN ────────────────────────────────────────────
        print("\n[1/5] Logéate en SSASUR (tienes 5 minutos)...")
        await page.goto(DASHBOARD_URL)
        try:
            await page.wait_for_selector(
                'button:has-text("ABASTECIMIENTO"), div:has-text("ABASTECIMIENTO")',
                timeout=TIMEOUT_LOGIN,
            )
        except Exception:
            await page.wait_for_function(
                "document.body.innerText.includes('ABASTECIMIENTO')",
                timeout=TIMEOUT_LOGIN,
            )
        await context.storage_state(path=str(SESSION_FILE))
        print("  ✓ Sesión detectada")

        # ════════════════════════════════════════════════════════════════════
        #  PASO 2 — RECETA  (entrar → informe → descargar)
        # ════════════════════════════════════════════════════════════════════
        print(f"\n[2/5] Módulo RECETA  ({fecha_inicio} → {fecha_fin})...")
        await entrar_modulo(page, "RECETA")

        await page.goto(RECETA_INFORME)
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(2_000)

        try:
            await page.fill("#fechaInicio", fecha_inicio)
            await page.fill("#fechaTermino", fecha_fin)
            print(f"  Fechas: {fecha_inicio} → {fecha_fin}")
        except Exception as e:
            print(f"  [AVISO] No se pudieron rellenar las fechas: {e}")

        print("  Descargando recetas... (puede tardar varios minutos)")
        try:
            # Caso A: "Buscar" dispara la descarga directamente
            await descargar(
                page, MAESTRO_DIR,
                lambda: page.click('button:has-text("Buscar"), input[value="Buscar"], button[type="submit"]'),
                "RECETA", timeout=120_000,
            )
        except Exception:
            # Caso B: "Buscar" solo corre la consulta; el archivo sale con "Descargar Excel"
            try:
                await page.wait_for_timeout(2_000)
                await descargar(
                    page, MAESTRO_DIR,
                    lambda: page.click('button:has-text("Descargar Excel")'),
                    "RECETA",
                )
            except Exception as e:
                print(f"  [ERROR] Descarga RECETA falló: {e}")
                await page.screenshot(path=str(MAESTRO_DIR / "debug_receta.png"))
                print("  Screenshot guardado: debug_receta.png")

        # ════════════════════════════════════════════════════════════════════
        #  PASO 3 — ABASTECIMIENTO  (volver al dashboard → entrar → stock)
        # ════════════════════════════════════════════════════════════════════
        print("\n[3/5] Módulo ABASTECIMIENTO...")
        await page.goto(DASHBOARD_URL)
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(1_500)
        await entrar_modulo(page, "ABASTECIMIENTO")

        await page.goto(STOCK_REPORTE)
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(2_500)

        if "login.ssasur.cl" in page.url:
            # Por si la sesión del módulo no quedó lista: reintentar una vez
            print("  [AVISO] Rebote al login — reintentando entrar a ABASTECIMIENTO...")
            await page.goto(DASHBOARD_URL)
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1_500)
            await entrar_modulo(page, "ABASTECIMIENTO")
            await page.goto(STOCK_REPORTE)
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(2_500)

        # Configurar el reporte: establecimiento Pitrufquén + bodega TODAS
        try:
            await page.select_option("#slctHeaderEstab", ESTAB_PITRUFQUEN)
            await page.wait_for_timeout(1_500)
        except Exception as e:
            print(f"  [info] establecimiento: {e}")
        try:
            await page.select_option("#bodega", BODEGA_TODAS)
            await page.wait_for_timeout(1_000)
            print("  Bodega: TODAS")
        except Exception as e:
            print(f"  [AVISO] No se pudo seleccionar bodega TODAS: {e}")

        print("  Generando XLS de stock... (puede tardar varios minutos)")
        try:
            await descargar(
                page, MAESTRO_DIR,
                lambda: page.click('button:has-text("Generar XLS"), button:has-text("XLS"), button:has-text("Excel")'),
                "ABASTECIMIENTO",
            )
        except Exception as e:
            print(f"  [ERROR] Descarga ABASTECIMIENTO falló: {e}")
            await page.screenshot(path=str(MAESTRO_DIR / "debug_stock.png"))
            print("  Screenshot guardado: debug_stock.png")

        await browser.close()

    # ── PASO 4 — MAESTRO AA ────────────────────────────────────────────────────
    print("\n[4/5] Actualizando Maestro AA...")
    env_utf8 = {**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}
    result = subprocess.run(
        [sys.executable, str(MAESTRO_DIR / "maestro_aa.py")],
        cwd=str(MAESTRO_DIR),
        env=env_utf8,
    )

    print()
    if result.returncode == 0:
        print("═" * 62)
        print("  ✓ COMPLETADO")
        print("  → Consolidado_AA_MAESTRO.xlsx  actualizado")
        print("  → Resumen_Pedidos_AA.xlsx       actualizado")
        print("═" * 62)

        # ── PASO 5 — PUBLICAR EN GITHUB ───────────────────────────────────────
        git_dir  = MAESTRO_DIR / ".git"
        publicar = MAESTRO_DIR / "PUBLICAR_DATOS.bat"
        if git_dir.exists() and publicar.exists():
            print("\n[5/5] Publicando datos en GitHub...")
            pub = subprocess.run(
                ["cmd", "/c", str(publicar)],
                cwd=str(MAESTRO_DIR),
            )
            if pub.returncode != 0:
                print("  [AVISO] Publicación falló — ejecuta PUBLICAR_DATOS.bat manualmente.")
        else:
            print("\n[5/5] GitHub no configurado — omitiendo publicación.")
            print("      Ejecuta CONFIGURAR_GITHUB.bat para activarlo.")
    else:
        print("  [ERROR] maestro_aa.py falló — revisa los mensajes arriba")

    try:
        input("\nPresiona Enter para cerrar...")
    except EOFError:
        pass


if __name__ == "__main__":
    asyncio.run(main())
