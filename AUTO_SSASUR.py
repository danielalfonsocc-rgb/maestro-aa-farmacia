#!/usr/bin/env python3
"""
AUTO_SSASUR.py  —  Maestro AA Farmacia
═══════════════════════════════════════════════════════════════
Descarga automáticamente RECETA y ABASTECIMIENTO desde SSASUR
y actualiza los Excel del Maestro AA.

Flujo:
  1. Se abre Chromium en la página de login de SSASUR
  2. Tú te logeas manualmente (el script no toca tus credenciales)
  3. El script detecta el login y descarga los archivos
  4. Ejecuta maestro_aa.py para actualizar el Consolidado
═══════════════════════════════════════════════════════════════
"""
import asyncio
import subprocess
import sys
from pathlib import Path
from datetime import date


# ── Configuración ─────────────────────────────────────────────────────────────
MAESTRO_DIR   = Path(__file__).parent
SESSION_FILE  = MAESTRO_DIR / ".ssasur_session.json"   # Guarda cookies entre sesiones
TIMEOUT_LOGIN = 300_000   # 5 minutos para logarse
TIMEOUT_DESCARGA = 600_000  # 10 minutos por descarga


def fmt(d: date) -> str:
    return d.strftime("%d/%m/%Y")


async def fill_date(page, label_text: str, value: str):
    """Rellena un campo de fecha buscando por su etiqueta."""
    idx = 0 if "Inicio" in label_text else 1
    filled = False

    # Intento 1: fill() por label (más robusto)
    try:
        field = page.get_by_label(label_text, exact=False).first
        await field.fill(value)
        await field.press("Tab")
        filled = True
    except Exception:
        pass

    # Intento 2: fill() por índice
    if not filled:
        try:
            inputs = page.locator('input[type="text"]')
            await inputs.nth(idx).fill(value)
            await inputs.nth(idx).press("Tab")
            filled = True
        except Exception:
            pass

    # Intento 3: fill() por placeholder o nombre
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
    fecha_inicio = fmt(today.replace(day=1))   # Primer día del mes actual
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

        # Restaurar sesión si existe (evita re-login si las cookies son válidas)
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
        print("\n[1/4] Logéate en SSASUR (tienes 5 minutos)...")
        await page.goto("https://login.ssasur.cl/dashboard")

        # Esperar el dashboard (aparece el módulo ABASTECIMIENTO solo si está logueado)
        try:
            await page.wait_for_selector(
                'div:has-text("ABASTECIMIENTO"), button:has-text("ABASTECIMIENTO"), [class*="ABASTECIMIENTO"]',
                timeout=TIMEOUT_LOGIN,
            )
        except Exception:
            # Fallback: esperar cualquier texto del dashboard
            await page.wait_for_function(
                "document.body.innerText.includes('ABASTECIMIENTO')",
                timeout=TIMEOUT_LOGIN,
            )

        # Guardar cookies para la próxima vez
        await context.storage_state(path=str(SESSION_FILE))
        print("  ✓ Sesión detectada — iniciando descargas automáticas...")

        # ── PASO 1b — ENTRAR A MÓDULO RECETA (establece cookies de www.ssasur.cl) ──
        print("  Entrando al módulo RECETA...")
        try:
            await page.click('a:has-text("RECETA"), div:has-text("RECETA") >> nth=0')
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(2_000)
            print(f"  URL después de clic RECETA: {page.url}")
        except Exception as e:
            print(f"  [AVISO] No se pudo hacer clic en RECETA: {e}")
            # Intentar navegación directa
            await page.goto("https://www.ssasur.cl/receta")
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(2_000)

        # ── PASO 2 — RECETA: Informe Completo ─────────────────────────────────
        print(f"\n[2/4] Informe Completo Recetas  ({fecha_inicio} → {fecha_fin})...")

        # Intentar la URL correcta directamente
        receta_urls = [
            "https://www.ssasur.cl/receta/informes/sabana",
        ]
        loaded = False
        for url in receta_urls:
            await page.goto(url)
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1_500)
            # Verificar si hay inputs de fecha (más de 1 input o input con nombre de fecha)
            cnt = await page.locator('input[type="text"], input[type="date"]').count()
            title = await page.title()
            print(f"  URL probada: {url} → {cnt} inputs, title: {title[:50]}")
            if cnt >= 2 or "receta" in title.lower() or "informe" in title.lower():
                loaded = True
                print(f"  ✓ Formulario encontrado en: {url}")
                break

        if not loaded:
            # Navegar por el menú: Reportes → Informe Completo
            print("  Navegando por menú...")
            await page.goto("https://www.ssasur.cl/receta")
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1_000)
            try:
                await page.click('a:has-text("Reportes"), a:has-text("Informes")')
                await page.wait_for_timeout(800)
                await page.click('a:has-text("Informe Completo"), a:has-text("Sábana")')
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(1_500)
            except Exception as e:
                print(f"  [AVISO] Navegación menú: {e}")

        await page.wait_for_timeout(1_000)

        # Screenshot del formulario antes de rellenar
        await page.screenshot(path=str(MAESTRO_DIR / "debug_formulario.png"))

        # Debug: imprimir inputs visibles solamente
        try:
            visible_inputs = await page.evaluate("""
                () => Array.from(document.querySelectorAll('input:not([type=hidden])'))
                    .map(i => ({type: i.type, name: i.name, id: i.id,
                                placeholder: i.placeholder, value: i.value,
                                label: i.labels?.[0]?.textContent?.trim() || ''}))
            """)
            print(f"  Inputs visibles: {len(visible_inputs)}")
            for inp in visible_inputs:
                print(f"    {inp}")
        except Exception as e:
            print(f"  [debug] {e}")

        # Rellenar fechas con JS directo como fallback final
        try:
            set_ok = await page.evaluate(f"""
                () => {{
                    const inputs = Array.from(document.querySelectorAll('input:not([type=hidden])'));
                    const dateInputs = inputs.filter(i =>
                        i.type === 'date' || i.type === 'text' &&
                        (i.placeholder?.includes('/') || i.id?.toLowerCase().includes('fech') ||
                         i.name?.toLowerCase().includes('fech'))
                    );
                    if (dateInputs.length >= 2) {{
                        // Disparar eventos para que la app React/Vue detecte el cambio
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
                    return `FAIL: solo ${{inputs.length}} inputs totales, ${{dateInputs.length}} de fecha`;
                }}
            """)
            print(f"  JS fill: {set_ok}")
        except Exception as e:
            print(f"  [AVISO] JS fill falló: {e}")
            await fill_date(page, "Fecha Inicio", fecha_inicio)
            await fill_date(page, "Fecha Término", fecha_fin)

        await page.wait_for_timeout(800)

        # Verificar fechas (intentar con text o date)
        try:
            inputs = page.locator('input[type="text"], input[type="date"]')
            cnt = await inputs.count()
            if cnt >= 2:
                v0 = await inputs.nth(0).input_value()
                v1 = await inputs.nth(1).input_value()
                print(f"  Fechas en formulario: '{v0}'  →  '{v1}'")
            else:
                print(f"  [AVISO] Solo {cnt} input(s) encontrado(s)")
        except Exception as e:
            print(f"  [AVISO] No se pudo verificar fechas: {e}")

        print("  Descargando... (puede tardar varios minutos)")
        try:
            async with page.expect_download(timeout=TIMEOUT_DESCARGA) as dl_info:
                await page.click('button:has-text("Buscar"), input[value="Buscar"], button[type="submit"]')
            dl = await dl_info.value
            dest = MAESTRO_DIR / dl.suggested_filename
            await dl.save_as(dest)
            size_kb = dest.stat().st_size // 1024
            print(f"  ✓ {dl.suggested_filename}  ({size_kb:,} KB)")
        except Exception as e:
            print(f"  [ERROR] Descarga RECETA falló: {e}")
            await page.screenshot(path=str(MAESTRO_DIR / "debug_receta.png"))
            print("  Screenshot guardado: debug_receta.png")

        # ── PASO 3 — ABASTECIMIENTO: Stock actual ──────────────────────────────
        print("\n[3/4] Stock actual ABASTECIMIENTO...")

        # La tarjeta ABASTECIMIENTO en el dashboard de login abre una NUEVA PESTAÑA
        # (target="_blank"). Por eso page.url nunca cambiaba al hacer clic y el
        # script pensaba que la navegación había fallado. Capturamos la nueva pestaña
        # con context.expect_page(); esa pestaña ya tiene las cookies del módulo
        # www.ssasur.cl establecidas por el propio login.ssasur.cl al abrirla.
        print("  Volviendo al dashboard para abrir ABASTECIMIENTO (nueva pestaña)...")
        await page.goto("https://login.ssasur.cl/dashboard")
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(1_500)

        abast_page = None
        try:
            async with context.expect_page(timeout=8_000) as new_tab_info:
                await page.click('a:has-text("ABASTECIMIENTO")')
            abast_page = await new_tab_info.value
            await abast_page.wait_for_load_state("networkidle")
            await abast_page.wait_for_timeout(2_000)
            print(f"  URL nueva pestaña ABASTECIMIENTO: {abast_page.url}")
        except Exception as e:
            print(f"  [AVISO] No se capturó nueva pestaña ({e}); usando página actual.")
            abast_page = page

        # Si la pestaña aterrizó en login (inesperado), navegar directo al módulo
        if "login.ssasur.cl" in abast_page.url:
            print("  Navegando directo a raíz de ABASTECIMIENTO para establecer sesión...")
            await abast_page.goto("https://www.ssasur.cl/abastecimiento")
            await abast_page.wait_for_load_state("networkidle")
            await abast_page.wait_for_timeout(2_000)
            print(f"  URL tras goto directo: {abast_page.url}")

        # Navegar al reporte de stock con verificación (hasta 2 intentos)
        stock_url = "https://www.ssasur.cl/abastecimiento/reportes/stock_en_momento_bodega"
        on_report = False
        for intento in range(2):
            await abast_page.goto(stock_url)
            await abast_page.wait_for_load_state("networkidle")
            await abast_page.wait_for_timeout(2_000)
            print(f"  URL stock (intento {intento + 1}): {abast_page.url}")
            if "login.ssasur.cl" not in abast_page.url:
                on_report = True
                break
            print("  [AVISO] Redirigido al dashboard — reestableciendo sesión de ABASTECIMIENTO...")
            await abast_page.goto("https://www.ssasur.cl/abastecimiento")
            await abast_page.wait_for_load_state("networkidle")
            await abast_page.wait_for_timeout(2_000)

        if not on_report:
            print("  [ERROR] No se pudo llegar al reporte de stock (sigue en login.ssasur.cl)")
            await abast_page.screenshot(path=str(MAESTRO_DIR / "debug_stock.png"))
            print("  Screenshot guardado: debug_stock.png")
        else:
            print("  Generando XLS... (puede tardar varios minutos)")
            try:
                async with abast_page.expect_download(timeout=TIMEOUT_DESCARGA) as dl_info:
                    await abast_page.click('button:has-text("Generar XLS"), button:has-text("XLS"), button:has-text("Excel")')
                dl = await dl_info.value
                dest = MAESTRO_DIR / dl.suggested_filename
                await dl.save_as(dest)
                size_kb = dest.stat().st_size // 1024
                print(f"  ✓ {dl.suggested_filename}  ({size_kb:,} KB)")
            except Exception as e:
                print(f"  [ERROR] Descarga ABASTECIMIENTO falló: {e}")
                await abast_page.screenshot(path=str(MAESTRO_DIR / "debug_stock.png"))
                print("  Screenshot guardado: debug_stock.png")

        await browser.close()

    # ── PASO 4 — MAESTRO AA ────────────────────────────────────────────────────
    print("\n[4/4] Actualizando Maestro AA...")
    result = subprocess.run(
        [sys.executable, str(MAESTRO_DIR / "maestro_aa.py")],
        cwd=str(MAESTRO_DIR),
    )

    print()
    if result.returncode == 0:
        print("═" * 62)
        print("  ✓ COMPLETADO")
        print("  → Consolidado_AA_MAESTRO.xlsx  actualizado")
        print("  → Resumen_Pedidos_AA.xlsx       actualizado")
        print("═" * 62)

        # ── PASO 5 — PUBLICAR EN GITHUB (solo si el repo está configurado) ────
        git_dir = MAESTRO_DIR / ".git"
        publicar = MAESTRO_DIR / "PUBLICAR_DATOS.bat"
        if git_dir.exists() and publicar.exists():
            print("\n[5/5] Publicando datos en GitHub...")
            pub = subprocess.run(
                ["cmd", "/c", str(publicar)],
                cwd=str(MAESTRO_DIR),
                capture_output=False,
            )
            if pub.returncode != 0:
                print("  [AVISO] Publicación falló — ejecuta PUBLICAR_DATOS.bat manualmente.")
        else:
            print("\n[5/5] GitHub no configurado — omitiendo publicación.")
            print("      Ejecuta CONFIGURAR_GITHUB.bat para activar la publicación automática.")
    else:
        print("  [ERROR] maestro_aa.py falló — revisa los mensajes arriba")

    input("\nPresiona Enter para cerrar...")


if __name__ == "__main__":
    asyncio.run(main())
