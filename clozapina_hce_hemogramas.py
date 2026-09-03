"""
clozapina_hce_hemogramas.py — Descarga automatizada de hemogramas desde HCE
(SSASur) para los pacientes con Clozapina 100mg despachada en Atención Abierta,
y arma el consolidado listo para copiar/pegar en "Ingreso de Hemogramas en
Bloque" del MINSAL.

Reutiliza la sesión guardada de AUTO_SSASUR.py (.ssasur_session.json) — si ya
corriste AUTO_SSASUR hoy, no necesitas volver a loguearte en SSASUR.

FIRMA ELECTRÓNICA: al entrar al módulo HCE, la plataforma pide Rut + Clave de
Firma Electrónica UNA VEZ por sesión de navegador. El script NUNCA la escribe
por ti — ingrésala tú manualmente cuando aparezca el diálogo. Después de eso,
recorre todos los pacientes sin más intervención (confirmado por el usuario:
la clave se pide una sola vez, no por cada paciente).

DESCARGA DEL PDF: el botón "Ver PDF" es un <form method="POST" target="_blank">
con un campo oculto "url" (base64) — confirmado en vivo. Esperar a que
Playwright detecte la pestaña nueva resultó frágil (a veces nunca se abre en
el contexto de Playwright), así que en vez de clicar y perseguir la pestaña,
se reproduce el MISMO POST directo con context.request (comparte las cookies
de sesión del browser). Esto NO abre ninguna pestaña ni toca el estado JS de
la página, así que no choca con la advertencia de la propia app de "esta
aplicación solo debe ser utilizada en una única pestaña del navegador".

ESTADO: validado en vivo contra los 38 pacientes reales de julio 2026 (38/38
hemogramas descargados). El modal de Laboratorio se cierra explícitamente
(API de Bootstrap/JS) después de cada paciente, sea éxito o fallo — dejarlo
abierto bloquea la búsqueda del paciente siguiente.

Uso:
    py clozapina_hce_hemogramas.py informe_completo_recetas_b20260701.csv [--desde 2026-07-01] [--hasta 2026-07-31]
"""
import argparse
import asyncio
import sys
from datetime import date, datetime
from pathlib import Path

from clozapina_consolidar import (
    cargar_despachos_clozapina, extraer_hemograma_pdf, construir_fila,
    guardar_excel, FilaConsolidada, ruta_archivo_mes,
)

MAESTRO_DIR = Path(__file__).parent
SESSION_FILE = MAESTRO_DIR / ".ssasur_session.json"
DASHBOARD_URL = "https://login.ssasur.cl/dashboard"
HEMOGRAMAS_DIR = MAESTRO_DIR / "_hemogramas_clozapina"
TIMEOUT_LOGIN = 300_000
TIMEOUT_FIRMA = 300_000


def _run_a_partes(run: str) -> tuple[str, str]:
    """'23797344-5' -> ('23797344', '5'). Acepta también '23.797.344-5'."""
    limpio = run.replace(".", "").strip()
    num, dv = limpio.rsplit("-", 1)
    return num.strip(), dv.strip()


def _cache_key(run_sin_guion: str, fecha_despacho) -> str:
    """Cache por (paciente, despacho) — NO solo por paciente. Bug real detectado
    en vivo (11-08-2026, caso Gina Cicarelli): con cache solo por RUN, un
    paciente que ya tenía un PDF descargado de un despacho anterior nunca
    volvía a consultarse en HCE en despachos posteriores, aunque el sistema ya
    tuviera un hemograma más nuevo — el reporte seguía mostrando el examen
    viejo (y su antigüedad calculada quedaba falsamente alta o baja). Cada
    despacho nuevo del mismo paciente debe buscar de nuevo en HCE."""
    if fecha_despacho is None or fecha_despacho != fecha_despacho:  # NaT check
        return run_sin_guion
    fecha = fecha_despacho.date() if hasattr(fecha_despacho, "date") else fecha_despacho
    return f"{run_sin_guion}_{fecha:%Y%m%d}"


async def entrar_hce(page):
    """Dashboard → tarjeta HCE. Igual que RECETA en AUTO_SSASUR.py: SSASur acuña
    la sesión del módulo al clicar la tarjeta, un deep-link directo puede fallar."""
    if "ssasur.cl/hce" in page.url:
        return
    await page.goto(DASHBOARD_URL, wait_until="networkidle")
    btn = page.locator('button:has-text("HCE")').first
    await btn.wait_for(state="visible", timeout=30_000)
    await btn.click()
    await page.wait_for_url(lambda u: "ssasur.cl/hce" in u, timeout=30_000)
    await page.wait_for_load_state("networkidle")
    print("  [OK] En modulo HCE")


async def esperar_firma_electronica(page):
    print("\n  Si aparece el dialogo 'Firma Electronica', ingresa tu Rut + Clave "
          "MANUALMENTE (el script nunca la escribe). Tienes 5 minutos.")
    await page.wait_for_selector('input[placeholder="Rut Paciente"]',
                                  timeout=TIMEOUT_FIRMA, state="visible")
    print("  [OK] Firma Electronica completada - listo para buscar pacientes.")


async def buscar_paciente(page, run_num: str, dv: str, debug: bool = False) -> bool:
    campo_rut = page.locator('input[placeholder="Rut Paciente"]').first
    campo_dv = page.locator('input[placeholder="dv"]').first
    await campo_rut.click()
    await campo_rut.fill("")
    await campo_rut.type(run_num, delay=30)
    await campo_dv.click()
    await campo_dv.fill("")
    await campo_dv.type(dv, delay=30)
    await page.click('button:has-text("Buscar")')
    await page.wait_for_timeout(2_000)
    # Tras buscar, si hay match el bloque de módulos del paciente (con el botón
    # "Laboratorio N") aparece directo bajo la fila de resultado — no hace falta
    # clicar ningún botón adicional para "abrir" la ficha.
    laboratorio_btn = page.locator('button:has-text("Laboratorio")').first
    try:
        await laboratorio_btn.wait_for(state="visible", timeout=8_000)
    except Exception:
        if debug:
            MAESTRO_DIR.mkdir(exist_ok=True)
            await page.screenshot(path=str(MAESTRO_DIR / f"_debug_busqueda_{run_num}.png"))
            texto = await page.evaluate("document.body.innerText")
            (MAESTRO_DIR / f"_debug_busqueda_{run_num}.txt").write_text(texto, encoding="utf-8")
            print(f"    [debug] rut_valor={await campo_rut.input_value()!r} dv_valor={await campo_dv.input_value()!r}")
        return False
    return True


async def _cerrar_modal(page):
    """Cierra el modal Bootstrap de Laboratorio ('#modalHce'). Confirmado en vivo:
    la tecla Escape NO lo cierra y deja el modal + su backdrop bloqueando clics
    del paciente siguiente (Locator.click timeout: '...subtree intercepts
    pointer events'). Se cierra por API de Bootstrap/jQuery si está disponible,
    y se limpian clases/backdrop a mano como respaldo por si no lo está."""
    try:
        await page.evaluate("""
            () => {
                const jq = window.jQuery || window.$;
                if (jq) { try { jq('.modal.in, .modal.show').modal('hide'); } catch (e) {} }
                document.querySelectorAll('.modal.in, .modal.show').forEach(m => {
                    m.classList.remove('in', 'show');
                    m.style.display = 'none';
                    m.setAttribute('aria-hidden', 'true');
                });
                document.querySelectorAll('.modal-backdrop').forEach(el => el.remove());
                document.body.classList.remove('modal-open');
            }
        """)
    except Exception:
        pass
    await page.wait_for_timeout(300)


async def _debug_estado(page, run_sin_guion: str, etapa: str):
    MAESTRO_DIR.mkdir(exist_ok=True)
    try:
        await page.screenshot(path=str(MAESTRO_DIR / f"_debug_{etapa}_{run_sin_guion}.png"))
        texto = await page.evaluate("document.body.innerText")
        (MAESTRO_DIR / f"_debug_{etapa}_{run_sin_guion}.txt").write_text(texto, encoding="utf-8")
    except Exception as e:
        print(f"    [debug] no se pudo capturar estado ({etapa}): {e}")


async def descargar_hemograma_mas_reciente(context, page, run_sin_guion: str, cache_key: str, debug: bool = False) -> Path | None:
    """Abre el modal de Laboratorio y captura el PDF del examen MAS RECIENTE
    (primera fila: la tabla ya viene ordenada descendente por fecha). SIEMPRE
    cierra el modal antes de retornar (éxito o fallo) — confirmado en vivo que
    si queda abierto, bloquea la búsqueda del paciente siguiente."""
    await page.click('button:has-text("Laboratorio")')
    try:
        return await _intentar_descarga_pdf(page, context, run_sin_guion, cache_key, debug)
    finally:
        await _cerrar_modal(page)


async def _intentar_descarga_pdf(page, context, run_sin_guion: str, cache_key: str, debug: bool) -> Path | None:
    # El modal "Históricos Clínicos de Exámenes" abre con la pestaña Laboratorio
    # activa por defecto, ya ordenada descendente por fecha — el primer botón
    # "Ver PDF" que aparezca en la página ES el examen más reciente.
    primer_ver_pdf = page.locator('button:has-text("Ver PDF")').first
    try:
        await primer_ver_pdf.wait_for(state="visible", timeout=15_000)
    except Exception:
        if debug:
            print(f"    [debug] {run_sin_guion}: no aparecio ningun boton 'Ver PDF' tras abrir Laboratorio")
            await _debug_estado(page, run_sin_guion, "filas")
        return None

    # El botón vive en <form class="pdfView" action="/hce/pdf" method="POST"
    # target="_blank"><input type="hidden" name="url" value="BASE64...">
    # (confirmado en vivo). Es un form NATIVO — pero esperar a que Playwright
    # detecte la pestaña nueva resultó frágil (no siempre se detecta). En vez
    # de clicar y perseguir la pestaña, reproducimos el MISMO POST directo con
    # context.request (comparte las cookies de sesión del browser, no abre
    # ninguna pestaña ni toca el estado JS de la página — no viola el aviso de
    # "una única pestaña" de la app, porque no es una pestaña).
    try:
        url_valor = await primer_ver_pdf.evaluate(
            "btn => { const f = btn.closest('form'); const i = f && f.querySelector('input[name=url]'); return i ? i.value : null; }"
        )
    except Exception as e:
        if debug:
            print(f"    [debug] {run_sin_guion}: no se pudo leer el campo 'url' del form: {e}")
        return None

    if not url_valor:
        if debug:
            print(f"    [debug] {run_sin_guion}: campo 'url' del form vino vacío")
        return None

    try:
        resp = await context.request.post("https://www.ssasur.cl/hce/pdf", form={"url": url_valor})
        if not resp.ok:
            if debug:
                print(f"    [debug] {run_sin_guion}: POST /hce/pdf devolvió status {resp.status}")
            return None
        pdf_bytes = await resp.body()
    except Exception as e:
        if debug:
            print(f"    [debug] {run_sin_guion}: fallo el POST /hce/pdf: {e}")
        return None

    HEMOGRAMAS_DIR.mkdir(exist_ok=True)
    dest = HEMOGRAMAS_DIR / f"{cache_key}.pdf"
    dest.write_bytes(pdf_bytes)
    return dest


async def limpiar_busqueda(page):
    try:
        await page.click('button:has-text("Limpiar")', timeout=3_000)
    except Exception:
        pass


def _construir_filas_desde_cache(por_paciente):
    """Arma filas_minsal/filas_obs solo con PDFs ya cacheados, sin tocar el
    navegador — usado cuando no hay ningún despacho nuevo que buscar."""
    hoy = date.today()
    filas_minsal, filas_obs, fallidos = [], [], []
    for _, row in por_paciente.iterrows():
        run = row["run_normalizado"]
        run_sin_guion = run.replace("-", "").replace(".", "")
        nombre = f"{row['Nombre']} {row['Apellido Paterno']} {row['Apellido Materno']}"
        pdf_cacheado = HEMOGRAMAS_DIR / f"{_cache_key(run_sin_guion, row['fecha_despacho'])}.pdf"
        hemograma = None
        if pdf_cacheado.exists():
            try:
                hemograma = extraer_hemograma_pdf(pdf_cacheado)
            except Exception as e:
                print(f"  [aviso] PDF cacheado de {run_sin_guion} no se pudo leer: {e}")
                fallidos.append(run_sin_guion)
        else:
            fallidos.append(run_sin_guion)
        fc = FilaConsolidada(
            run=run, nombre=nombre,
            fecha_despacho=row["fecha_despacho"].date() if row["fecha_despacho"] == row["fecha_despacho"] else None,
            dosis_mg_dia=row["dosis_mg_dia"], cuota_receta=row["cuota_receta"] or "",
            posologia_texto=row["posologia_texto"] or "", hemograma=hemograma,
        )
        fm, fo = construir_fila(fc, hoy)
        filas_minsal.append(fm)
        filas_obs.append(fo)
    return filas_minsal, filas_obs, fallidos


async def main_async(csv_recetas: str, desde, hasta, salida: str, solo_rut: str = None, forzar: bool = False):
    from playwright.async_api import async_playwright

    despachos = cargar_despachos_clozapina(csv_recetas, desde, hasta)
    por_paciente = despachos.sort_values("fecha_despacho").groupby("run_normalizado").last()
    por_paciente = por_paciente.reset_index().sort_values("fecha_despacho")  # planilla: más antigua -> más nueva
    if solo_rut:
        solo_rut_norm = solo_rut.replace(".", "").replace(" ", "").upper()
        por_paciente = por_paciente[por_paciente["run_normalizado"].str.replace(".", "", regex=False).str.replace(" ", "", regex=False).str.upper() == solo_rut_norm]
    print(f"{len(por_paciente)} pacientes con Clozapina (Atencion Abierta) en el periodo.")

    def _tiene_pdf_cacheado(run_sin_guion, fecha_despacho):
        return (HEMOGRAMAS_DIR / f"{_cache_key(run_sin_guion, fecha_despacho)}.pdf").exists()

    pendientes = [
        row for _, row in por_paciente.iterrows()
        if forzar or not _tiene_pdf_cacheado(row["run_normalizado"].replace("-", "").replace(".", ""), row["fecha_despacho"])
    ]

    if not pendientes:
        # Todos los pacientes ya tienen su hemograma cacheado de una corrida
        # anterior — no hace falta ni abrir el navegador ni pedir Firma
        # Electrónica para no buscar nada nuevo.
        print("Sin despachos nuevos — todos los pacientes ya tienen hemograma cacheado. "
              "No se abre el navegador.")
        filas_minsal, filas_obs, fallidos = _construir_filas_desde_cache(por_paciente)
        guardar_excel(filas_minsal, filas_obs, salida)
        print(f"\n{len(por_paciente) - len(fallidos)}/{len(por_paciente)} hemogramas listos -> {salida}"
              f"  ({len(por_paciente)} ya cacheados, 0 nuevos buscados en HCE)")
        if fallidos:
            print(f"[AVISO] sin hemograma para: {', '.join(fallidos)} (revisar manualmente en HCE)")
        return

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
        )
        if SESSION_FILE.exists():
            context = await browser.new_context(accept_downloads=True, storage_state=str(SESSION_FILE))
            print("(Sesion SSASUR guardada encontrada - puede que no necesites logearte)")
        else:
            context = await browser.new_context(accept_downloads=True)

        page = await context.new_page()
        print("\n[1/3] Logeate en SSASUR si te lo pide (tienes 5 minutos)...")
        await page.goto(DASHBOARD_URL)
        try:
            await page.wait_for_selector('button:has-text("HCE")', timeout=TIMEOUT_LOGIN)
        except Exception:
            print("[ERROR] No se detecto el dashboard tras el tiempo de espera.")
            await browser.close()
            return

        await context.storage_state(path=str(SESSION_FILE))

        print("\n[2/3] Entrando a HCE...")
        await entrar_hce(page)
        await esperar_firma_electronica(page)

        print(f"\n[3/3] Buscando hemogramas de {len(por_paciente)} pacientes...\n")
        filas_minsal, filas_obs = [], []
        hoy = date.today()
        ok, fallidos = 0, []

        omitidos_cache = 0
        for _, row in por_paciente.iterrows():
            run = row["run_normalizado"]
            run_sin_guion = run.replace("-", "").replace(".", "")
            run_num, dv = _run_a_partes(run)
            nombre = f"{row['Nombre']} {row['Apellido Paterno']} {row['Apellido Materno']}"

            cache_key = _cache_key(run_sin_guion, row["fecha_despacho"])
            pdf_cacheado = HEMOGRAMAS_DIR / f"{cache_key}.pdf"
            hemograma = None
            if not forzar and pdf_cacheado.exists():
                # Ya se descargó el hemograma para ESTE despacho puntual (mismo
                # paciente + misma fecha de despacho) en una corrida anterior —
                # no hace falta volver a buscarlo en HCE. Un despacho nuevo del
                # mismo paciente en otra fecha SIEMPRE vuelve a buscar en HCE
                # (bug real 11-08-2026: con cache solo por RUN se reutilizaba
                # para siempre el PDF del primer despacho, ignorando exámenes
                # más nuevos ya cargados en el sistema).
                try:
                    hemograma = extraer_hemograma_pdf(pdf_cacheado)
                    ok += 1
                    omitidos_cache += 1
                except Exception as e:
                    print(f"  [aviso] PDF cacheado de {run_sin_guion} no se pudo leer, se vuelve a buscar: {e}")
                    hemograma = None

            if hemograma is None:
                encontrado = await buscar_paciente(page, run_num, dv, debug=True)
                if encontrado:
                    pdf_path = await descargar_hemograma_mas_reciente(context, page, run_sin_guion, cache_key, debug=True)
                    if pdf_path:
                        try:
                            hemograma = extraer_hemograma_pdf(pdf_path)
                            ok += 1
                        except Exception as e:
                            print(f"  [aviso] no se pudo leer el PDF de {run_sin_guion}: {e}")
                            fallidos.append(run_sin_guion)
                    else:
                        fallidos.append(run_sin_guion)
                else:
                    print(f"  [aviso] paciente {run_sin_guion} no encontrado en HCE")
                    fallidos.append(run_sin_guion)

                await limpiar_busqueda(page)

            fc = FilaConsolidada(
                run=run, nombre=nombre,
                fecha_despacho=row["fecha_despacho"].date() if row["fecha_despacho"] == row["fecha_despacho"] else None,
                dosis_mg_dia=row["dosis_mg_dia"], cuota_receta=row["cuota_receta"] or "",
                posologia_texto=row["posologia_texto"] or "", hemograma=hemograma,
            )
            fm, fo = construir_fila(fc, hoy)
            filas_minsal.append(fm)
            filas_obs.append(fo)

        await browser.close()

    guardar_excel(filas_minsal, filas_obs, salida)
    nuevos = len(por_paciente) - omitidos_cache
    print(f"\n{ok}/{len(por_paciente)} hemogramas listos -> {salida}"
          f"  ({omitidos_cache} ya cacheados, {nuevos} nuevos buscados en HCE)")
    if fallidos:
        print(f"[AVISO] sin hemograma para: {', '.join(fallidos)} (revisar manualmente en HCE)")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("csv_recetas")
    ap.add_argument("--desde", default=None)
    ap.add_argument("--hasta", default=None)
    ap.add_argument("--salida", default=None, help="Por defecto: Clozapina_Reportes/Consolidado_..._<Mes>_<Año>.xlsx")
    ap.add_argument("--rut", default=None, help="Solo un paciente (debug), ej. 18198602-6")
    ap.add_argument("--forzar", action="store_true",
                     help="Vuelve a buscar en HCE aunque ya haya un PDF cacheado para el paciente "
                          "(por defecto se reusa el PDF ya descargado — no se re-busca en cada corrida).")
    args = ap.parse_args()
    desde = datetime.strptime(args.desde, "%Y-%m-%d").date() if args.desde else None
    hasta = datetime.strptime(args.hasta, "%Y-%m-%d").date() if args.hasta else None
    salida = args.salida or str(ruta_archivo_mes())
    asyncio.run(main_async(args.csv_recetas, desde, hasta, salida, solo_rut=args.rut, forzar=args.forzar))


if __name__ == "__main__":
    main()
