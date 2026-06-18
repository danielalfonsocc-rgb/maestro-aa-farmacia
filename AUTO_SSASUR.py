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
from datetime import date, timedelta

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
BODEGA_TODAS     = "0"    # bodega → "TODAS" (todas las bodegas en un archivo)
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


# ── Selección de fecha por CALENDARIO (bootstrap-datepicker) ────────────────────
# El informe de recetas usa bootstrap-datepicker. La fecha SOLO queda registrada
# para el backend cuando se hace CLIC en una celda de día (eso dispara el evento
# interno 'changeDate'); asignar el .value por JS deja el texto visible pero el
# servidor la ignora y devuelve 0 filas. Por eso seleccionamos clic a clic.
_MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
             "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


async def _cerrar_calendario(page):
    """Cierra cualquier datepicker abierto: si no, el de un campo tapa al otro."""
    await page.keyboard.press("Escape")
    # Clic neutro en la franja superior vacía (lejos de campos y del calendario).
    try:
        await page.mouse.click(950, 130)
    except Exception:
        pass
    await page.wait_for_timeout(300)


async def seleccionar_fecha_calendario(page, input_id: str, d: date):
    """
    Selecciona una fecha haciendo CLIC REAL en el calendario del campo `input_id`.
    Abre el datepicker, navega al mes/año con « / » y clica el día (excluyendo
    los días atenuados de los meses contiguos).
    """
    # Abrir el calendario clicando el input.
    await page.click(f"#{input_id}")
    dp = page.locator(".datepicker:visible").first
    await dp.wait_for(state="visible", timeout=10_000)

    # Navegar al mes/año objetivo leyendo el encabezado «.datepicker-switch».
    objetivo = f"{_MESES_ES[d.month - 1]} {d.year}"
    switch = dp.locator(".datepicker-switch").first
    for _ in range(36):  # tope de seguridad (~3 años de saltos)
        actual = (await switch.inner_text()).strip().lower()
        if actual == objetivo:
            break
        partes = actual.split()
        try:
            m_act = _MESES_ES.index(partes[0]) + 1   # mes 1-12
            y_act = int(partes[1])                   # año
        except (ValueError, IndexError):
            break  # formato inesperado: confiar en el mes visible
        # Comparar (año, mes) para decidir dirección.
        if (y_act, m_act) > (d.year, d.month):
            await dp.locator(".prev").first.click()
        else:
            await dp.locator(".next").first.click()
        await page.wait_for_timeout(250)

    # Clic en el día del mes vigente (sin .old / .new = meses contiguos atenuados).
    dia = dp.locator(f"td.day:not(.old):not(.new):text-is('{d.day}')").first
    await dia.wait_for(state="visible", timeout=5_000)
    await dia.click()
    await page.wait_for_timeout(300)


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

    # Solo el DÍA ANTERIOR a la ejecución (evita re-descargar días repetidos):
    # si la tarea corre el 18/06, extrae las recetas del 17/06.
    today        = date.today()
    ayer         = today - timedelta(days=1)
    fecha_inicio = fmt(ayer)
    fecha_fin    = fmt(ayer)

    print()
    print("═" * 62)
    print("  AUTO SSASUR  ·  Maestro AA Farmacia")
    print(f"  Recetas del día: {fecha_inicio}")
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
        print(f"\n[2/5] Módulo RECETA  (día {fecha_inicio})...")
        await entrar_modulo(page, "RECETA")

        await page.goto(RECETA_INFORME)
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(2_000)

        def _contar_filas(ruta: Path) -> int:
            try:
                with open(ruta, encoding="latin-1") as _f:
                    return sum(1 for _ in _f) - 1
            except Exception:
                return -1

        # ── Selección de fechas POR CALENDARIO (clic real) ─────────────────────
        # Clave verificada: el backend SOLO registra la consulta si la fecha se
        # ELIGE clicando en el calendario. Por eso seleccionamos AYER clic a clic
        # en ambos campos (cerrando el calendario entre uno y otro). NO requiere
        # firma. El estado por defecto del formulario (Solicitado / sin domicilio /
        # no anuladas) es el correcto, así que no tocamos nada más.
        recetas_ok = False
        try:
            await seleccionar_fecha_calendario(page, "fechaInicio", ayer)
            await _cerrar_calendario(page)
            await seleccionar_fecha_calendario(page, "fechaTermino", ayer)
            await _cerrar_calendario(page)
            print(f"  ✓ Fechas seleccionadas en el calendario: {fecha_inicio} → {fecha_fin}")

            print("  Buscando y descargando recetas...")
            async with page.expect_download(timeout=180_000) as dl_info:
                await page.click(
                    'button:has-text("Buscar"), input[type="submit"][value*="Buscar"], '
                    'a:has-text("Buscar"), #btnBuscar',
                    force=True,
                )
            dl   = await dl_info.value
            dest = MAESTRO_DIR / dl.suggested_filename
            await dl.save_as(dest)
            n_filas = _contar_filas(dest)
            if n_filas <= 0:
                dest.unlink(missing_ok=True)
                print("  [AVISO] Informe VACÍO (0 filas) tras selección automática — paso a modo manual.")
            else:
                recetas_ok = True
                print(f"  ✓ {dl.suggested_filename}  ({dest.stat().st_size // 1024:,} KB · {n_filas:,} filas)")
        except Exception as e:
            print(f"  [AVISO] Selección automática de fechas falló: {e}")
            await page.screenshot(path=str(MAESTRO_DIR / "debug_receta.png"))

        # ── Respaldo SEMIautomático (si la automática no trajo datos) ──────────
        if not recetas_ok:
            print(f"  📋 RESPALDO MANUAL: selecciona AYER ({fecha_inicio}) en el CALENDARIO")
            print("     de 'Fecha Inicio' y 'Fecha Término' (CLIC en el día — no escribas")
            print("     la fecha a mano) y pulsa 'Buscar'.")
            print("     Esperando la descarga (hasta 3 min)... si hoy no la necesitas, ignóralo.")
            try:
                async with page.expect_download(timeout=180_000) as dl_info:
                    pass  # la descarga la dispara el usuario al pulsar Buscar
                dl   = await dl_info.value
                dest = MAESTRO_DIR / dl.suggested_filename
                await dl.save_as(dest)
                n_filas = _contar_filas(dest)
                if n_filas <= 0:
                    dest.unlink(missing_ok=True)
                    print("  [AVISO] Informe VACÍO (0 filas) — reselecciona la fecha en el calendario. Archivo descartado.")
                else:
                    print(f"  ✓ {dl.suggested_filename}  ({dest.stat().st_size // 1024:,} KB · {n_filas:,} filas)")
            except Exception:
                print("  [AVISO] Sin descarga de recetas. Continúo con el stock.")

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

        # Configurar el reporte: bodega = TODAS.
        # IMPORTANTE: NO seleccionar el establecimiento — hacerlo dispara un modal
        # ("Selección Proyecto") cuyo backdrop bloquea el botón Generar XLS.
        # El establecimiento ya viene en "PITRUFQUEN HOSP." (única opción).
        try:
            await page.select_option("#bodega", BODEGA_TODAS)
            await page.wait_for_timeout(1_000)
            print("  Bodega: TODAS")
        except Exception as e:
            print(f"  [AVISO] No se pudo seleccionar bodega TODAS: {e}")

        print("  Generando XLS de stock... (puede tardar varios minutos)")
        try:
            # force=True: el overlay "Generando reporte. Espere un momento."
            # puede interponerse al clic; force lo atraviesa.
            await descargar(
                page, MAESTRO_DIR,
                lambda: page.click("#generarXLS_stock", force=True),
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

        # ── PASO 4b — AUDITORÍA DE PRESCRIPCIÓN (para la pestaña en la nube) ──
        audit_py = MAESTRO_DIR / "auditoria_prescripcion.py"
        if audit_py.exists():
            print("\n  Generando auditoría de prescripción...")
            subprocess.run([sys.executable, str(audit_py)], cwd=str(MAESTRO_DIR), env=env_utf8)

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
