#!/usr/bin/env python3
"""
AUTO_SSASUR.py  —  Maestro AA Farmacia
═══════════════════════════════════════════════════════════════
Descarga automáticamente RECETA y ABASTECIMIENTO desde SSASUR
y actualiza los Excel del Maestro AA.

Flujo:
  1. Se abre Chromium en la página de login de SSASUR
  2. Tú te logeas manualmente (el script no toca tus credenciales)
  3. El script trabaja UN módulo a la vez (SSASUR exige usar una
     sola pestaña — abrir dos módulos a la vez causa errores):
       a) Abre RECETA desde el dashboard → descarga recetas → cierra
       b) Vuelve al dashboard → abre ABASTECIMIENTO → descarga stock
  4. Ejecuta maestro_aa.py para actualizar el Consolidado
  5. Publica los datos en GitHub (si está configurado)

NOTA TÉCNICA — por qué serializamos:
  La cookie de sesión de SSASUR vive solo en login.ssasur.cl. La
  sesión de cada módulo (www.ssasur.cl/...) se "acuña" al hacer clic
  en su tarjeta del dashboard. El dashboard advierte que la app solo
  debe usarse en UNA pestaña: tener RECETA y ABASTECIMIENTO abiertos
  a la vez rompe la sesión y rebota al login. Por eso abrimos y
  cerramos cada módulo por separado, justo antes de usarlo.
═══════════════════════════════════════════════════════════════
"""
import asyncio
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
TIMEOUT_LOGIN    = 300_000   # 5 minutos para logarse
TIMEOUT_DESCARGA = 600_000   # 10 minutos por descarga


def fmt(d: date) -> str:
    return d.strftime("%d/%m/%Y")


async def fill_date(page, label_text: str, value: str):
    """Rellena un campo de fecha buscando por su etiqueta."""
    idx = 0 if "Inicio" in label_text else 1
    filled = False
    try:
        field = page.get_by_label(label_text, exact=False).first
        await field.fill(value)
        await field.press("Tab")
        filled = True
    except Exception:
        pass
    if not filled:
        try:
            inputs = page.locator('input[type="text"]')
            await inputs.nth(idx).fill(value)
            await inputs.nth(idx).press("Tab")
            filled = True
        except Exception:
            pass
    if not filled:
        try:
            inputs = page.locator('input[type="date"], input[name*="fecha"], input[placeholder*="/"]')
            await inputs.nth(idx).fill(value)
            await inputs.nth(idx).press("Tab")
            filled = True
        except Exception:
            pass
    if not filled:
        print(f"  [AVISO] No se pudo rellenar '{label_text}'")


async def ir_al_dashboard(page):
    """Deja la pestaña lanzadera en el dashboard de login.ssasur.cl."""
    if "login.ssasur.cl" not in page.url:
        await page.goto(DASHBOARD_URL)
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(1_200)


async def abrir_modulo(context, page_dashboard, nombre: str):
    """
    Hace clic en la tarjeta 'nombre' del dashboard. La tarjeta abre el
    módulo en una pestaña nueva (target="_blank") y acuña su sesión vía
    el handshake SSO — la capturamos con context.expect_page().
    Si no abre pestaña nueva, navega directo a la RAÍZ del módulo
    (no a una URL profunda, que rebotaría al login).
    Devuelve el handle de la pestaña del módulo.
    """
    try:
        async with context.expect_page(timeout=15_000) as tab_info:
            await page_dashboard.click(f'a:has-text("{nombre}")')
        nueva = await tab_info.value
        await nueva.wait_for_load_state("networkidle")
        await nueva.wait_for_timeout(2_000)
        print(f"  ✓ Pestaña {nombre}: {nueva.url}")
        return nueva
    except Exception as e:
        print(f"  [AVISO] {nombre} no abrió pestaña nueva ({e}) — navegando a la raíz del módulo")
        nueva = await context.new_page()
        await nueva.goto(f"https://www.ssasur.cl/{nombre.lower()}")
        await nueva.wait_for_load_state("networkidle")
        await nueva.wait_for_timeout(2_000)
        print(f"  Pestaña {nombre} (directo): {nueva.url}")
        return nueva


async def cerrar_pestana(page):
    """Cierra una pestaña de módulo de forma segura."""
    try:
        await page.close()
    except Exception:
        pass


async def main():
    # ── Verificar Playwright ───────────────────────────────────────────────────
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print("\n[ERROR] Playwright no está instalado.")
        print("  Ejecuta AUTO_SSASUR.bat para instalarlo automáticamente.\n")
        input("Presiona Enter para cerrar...")
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

        page = await context.new_page()   # pestaña "lanzadera" — se queda en el dashboard

        # ── PASO 1 — DETECTAR LOGIN ────────────────────────────────────────────
        print("\n[1/5] Logéate en SSASUR (tienes 5 minutos)...")
        await page.goto(DASHBOARD_URL)

        try:
            await page.wait_for_selector(
                'div:has-text("ABASTECIMIENTO"), button:has-text("ABASTECIMIENTO")',
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
        #  PASO 2 — MÓDULO RECETA  (abrir → descargar → CERRAR)
        # ════════════════════════════════════════════════════════════════════
        print(f"\n[2/5] Módulo RECETA  ({fecha_inicio} → {fecha_fin})...")
        page_receta = await abrir_modulo(context, page, "RECETA")

        # Llegar al Informe Completo. Primero intentamos la URL del informe;
        # si no carga el formulario, navegamos por el menú desde la raíz.
        await page_receta.goto("https://www.ssasur.cl/receta/informes/sabana")
        await page_receta.wait_for_load_state("networkidle")
        await page_receta.wait_for_timeout(1_500)

        cnt   = await page_receta.locator('input[type="text"], input[type="date"]').count()
        title = await page_receta.title()
        print(f"  {page_receta.url[:70]}  →  {cnt} inputs")

        if cnt < 2 and "receta" not in title.lower() and "informe" not in title.lower():
            print("  Navegando por menú...")
            try:
                await page_receta.goto("https://www.ssasur.cl/receta")
                await page_receta.wait_for_load_state("networkidle")
                await page_receta.wait_for_timeout(1_000)
                await page_receta.click('a:has-text("Reportes"), a:has-text("Informes")')
                await page_receta.wait_for_timeout(800)
                await page_receta.click('a:has-text("Informe Completo"), a:has-text("Sábana")')
                await page_receta.wait_for_load_state("networkidle")
                await page_receta.wait_for_timeout(1_500)
            except Exception as e:
                print(f"  [AVISO] Navegación menú: {e}")

        await page_receta.wait_for_timeout(1_000)
        await page_receta.screenshot(path=str(MAESTRO_DIR / "debug_formulario.png"))

        # Debug inputs
        try:
            vis = await page_receta.evaluate("""
                () => Array.from(document.querySelectorAll('input:not([type=hidden])'))
                    .map(i => ({type: i.type, name: i.name, id: i.id,
                                placeholder: i.placeholder, value: i.value,
                                label: i.labels?.[0]?.textContent?.trim() || ''}))
            """)
            print(f"  Inputs visibles: {len(vis)}")
            for inp in vis:
                print(f"    {inp}")
        except Exception as e:
            print(f"  [debug] {e}")

        # Rellenar fechas vía JS
        try:
            set_ok = await page_receta.evaluate(f"""
                () => {{
                    const inputs = Array.from(document.querySelectorAll('input:not([type=hidden])'));
                    const dateInputs = inputs.filter(i =>
                        i.type === 'date' || i.type === 'text' &&
                        (i.placeholder?.includes('/') || i.id?.toLowerCase().includes('fech') ||
                         i.name?.toLowerCase().includes('fech'))
                    );
                    if (dateInputs.length >= 2) {{
                        function setVal(el, val) {{
                            const nativeInputValueSetter = Object.getOwnPropertyDescriptor(
                                window.HTMLInputElement.prototype, 'value').set;
                            nativeInputValueSetter.call(el, val);
                            el.dispatchEvent(new Event('input', {{bubbles: true}}));
                            el.dispatchEvent(new Event('change', {{bubbles: true}}));
                        }}
                        setVal(dateInputs[0], '{fecha_inicio}');
                        setVal(dateInputs[1], '{fecha_fin}');
                        return `OK: ${{dateInputs[0].value}} / ${{dateInputs[1].value}}`;
                    }}
                    return `FAIL: ${{inputs.length}} inputs totales, ${{dateInputs.length}} de fecha`;
                }}
            """)
            print(f"  JS fill: {set_ok}")
        except Exception as e:
            print(f"  [AVISO] JS fill falló: {e}")
            await fill_date(page_receta, "Fecha Inicio", fecha_inicio)
            await fill_date(page_receta, "Fecha Término", fecha_fin)

        await page_receta.wait_for_timeout(800)

        try:
            inputs = page_receta.locator('input[type="text"], input[type="date"]')
            cnt = await inputs.count()
            if cnt >= 2:
                v0 = await inputs.nth(0).input_value()
                v1 = await inputs.nth(1).input_value()
                print(f"  Fechas: '{v0}'  →  '{v1}'")
        except Exception as e:
            print(f"  [AVISO] No se pudo verificar fechas: {e}")

        print("  Descargando... (puede tardar varios minutos)")
        try:
            async with page_receta.expect_download(timeout=TIMEOUT_DESCARGA) as dl_info:
                await page_receta.click('button:has-text("Buscar"), input[value="Buscar"], button[type="submit"]')
            dl = await dl_info.value
            dest = MAESTRO_DIR / dl.suggested_filename
            await dl.save_as(dest)
            size_kb = dest.stat().st_size // 1024
            print(f"  ✓ {dl.suggested_filename}  ({size_kb:,} KB)")
        except Exception as e:
            print(f"  [ERROR] Descarga RECETA falló: {e}")
            await page_receta.screenshot(path=str(MAESTRO_DIR / "debug_receta.png"))
            print("  Screenshot guardado: debug_receta.png")

        # CERRAR RECETA antes de abrir ABASTECIMIENTO (regla de única pestaña)
        await cerrar_pestana(page_receta)
        print("  Pestaña RECETA cerrada — liberando la sesión para ABASTECIMIENTO")

        # ════════════════════════════════════════════════════════════════════
        #  PASO 3 — MÓDULO ABASTECIMIENTO  (abrir → descargar)
        # ════════════════════════════════════════════════════════════════════
        print("\n[3/5] Módulo ABASTECIMIENTO...")
        await ir_al_dashboard(page)

        stock_url = "https://www.ssasur.cl/abastecimiento/reportes/stock_en_momento_bodega"
        page_abast = await abrir_modulo(context, page, "ABASTECIMIENTO")

        on_report = False
        for intento in range(2):
            # Asentar la sesión del módulo pasando por su RAÍZ antes del deep-link
            await page_abast.goto("https://www.ssasur.cl/abastecimiento")
            await page_abast.wait_for_load_state("networkidle")
            await page_abast.wait_for_timeout(1_500)

            if "login.ssasur.cl" not in page_abast.url:
                # Sesión del módulo OK → ir al reporte de stock
                await page_abast.goto(stock_url)
                await page_abast.wait_for_load_state("networkidle")
                await page_abast.wait_for_timeout(2_000)
                print(f"  URL stock (intento {intento + 1}): {page_abast.url}")
                if "login.ssasur.cl" not in page_abast.url:
                    on_report = True
                    break

            # Rebotó al login → reabrir el módulo desde el dashboard
            print(f"  [AVISO] Rebote a login (intento {intento + 1}) — reabriendo desde el dashboard...")
            await cerrar_pestana(page_abast)
            await ir_al_dashboard(page)
            page_abast = await abrir_modulo(context, page, "ABASTECIMIENTO")

        if not on_report:
            print("  [ERROR] No se pudo llegar al reporte de stock")
            await page_abast.screenshot(path=str(MAESTRO_DIR / "debug_stock.png"))
            print("  Screenshot guardado: debug_stock.png")
        else:
            print("  Generando XLS... (puede tardar varios minutos)")
            try:
                async with page_abast.expect_download(timeout=TIMEOUT_DESCARGA) as dl_info:
                    await page_abast.click(
                        'button:has-text("Generar XLS"), button:has-text("XLS"), button:has-text("Excel")'
                    )
                dl = await dl_info.value
                dest = MAESTRO_DIR / dl.suggested_filename
                await dl.save_as(dest)
                size_kb = dest.stat().st_size // 1024
                print(f"  ✓ {dl.suggested_filename}  ({size_kb:,} KB)")
            except Exception as e:
                print(f"  [ERROR] Descarga ABASTECIMIENTO falló: {e}")
                await page_abast.screenshot(path=str(MAESTRO_DIR / "debug_stock.png"))
                print("  Screenshot guardado: debug_stock.png")

        await browser.close()

    # ── PASO 4 — MAESTRO AA ────────────────────────────────────────────────────
    print("\n[4/5] Actualizando Maestro AA...")
    import os
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

    input("\nPresiona Enter para cerrar...")


if __name__ == "__main__":
    asyncio.run(main())
