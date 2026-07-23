#!/usr/bin/env python3
"""
descargar_recetas_pdf.py — Descarga de SSASUR el PDF individual de cada
receta (por Nº) para completar una Revisión de Solicitudes con los
documentos reales, no solo la lista de números que arma
revision_solicitudes.py --xlsx.

ADVERTENCIA — a diferencia de GT y la sábana (ver memoria del proyecto
ssasur-navegacion.md, ambos verificados con corridas reales), la búsqueda de
UNA receta por número NO tiene un flujo verificado todavía en este proyecto.
Este script hace lo mismo que se hizo para descubrir GT: entra al módulo
RECETA, intenta ubicar la pantalla de consulta con varios candidatos
razonables, y si no encuentra el formulario esperado, vuelca la página real
con --debug para poder terminar de ajustar los selectores con datos reales
en vez de adivinar a ciegas.

Primera corrida (obligatoria): usar --debug con 1-2 recetas nomás. Revisa la
salida [DESCUBRIR ...] y los screenshots (debug_consulta_receta_*.png) — con
eso se ajustan los selectores de búsqueda/descarga antes de correr el lote
completo.

Uso:
  py descargar_recetas_pdf.py --estab "CESFAM FREIRE" --recetas 45858462,46045176 --debug
  py descargar_recetas_pdf.py --estab "CESFAM FREIRE" --feedback "C:/ruta/Feedback_Solicitud_CESFAM_FREIRE_2026-07-23.xlsx"
"""
import argparse
import asyncio
import os
import re
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import AUTO_SSASUR as AS
from agregar_gt_manual import _CARPETA_LOCAL, GT_SOLICITUDES_DIR

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MAESTRO_DIR = os.path.dirname(os.path.abspath(__file__))

# Candidatos razonables — NINGUNO verificado en vivo todavía (ver advertencia
# arriba). Se prueban en orden; si ninguno funciona, --debug vuelca la página
# real para terminar de ajustar esto con datos ciertos.
CANDIDATOS_URL = [
    "https://www.ssasur.cl/receta/consulta",
    "https://www.ssasur.cl/receta/informes/consultaReceta",
    "https://www.ssasur.cl/receta/consultarReceta",
]
CANDIDATOS_MENU = [
    'a:has-text("Consultar Receta")', 'button:has-text("Consultar Receta")',
    'a:has-text("Buscar Receta")', 'button:has-text("Buscar Receta")',
    'a:has-text("Historial")', 'button:has-text("Historial")',
]
CANDIDATOS_INPUT_RECETA = [
    "#numero_receta",  # confirmado en vivo 23-07-2026 (Consultar Receta -> Reportes)
    'input[name="numeroReceta"]', 'input[id*="receta" i]', 'input[placeholder*="receta" i]',
]


def _leer_recetas_de_feedback(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    header_idx = None
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row and any(str(c).strip() == "N° Receta" for c in row if c):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError(f"No encontré la columna 'N° Receta' en {path}")
    header = [str(c).strip() if c else "" for c in list(ws.iter_rows(values_only=True))[header_idx - 1]]
    col_receta = header.index("N° Receta")
    col_pac = header.index("Paciente") if "Paciente" in header else None
    out = []
    for row in list(ws.iter_rows(values_only=True))[header_idx:]:
        receta = row[col_receta] if col_receta < len(row) else None
        if not receta or str(receta).startswith("SIN_RECETA_"):
            continue
        paciente = row[col_pac] if col_pac is not None and col_pac < len(row) else ""
        out.append((str(receta).strip(), str(paciente or "").strip()))
    return out


async def _abrir_consulta_receta(page, debug):
    # Confirmado en vivo (23-07-2026): "Consultar Receta" existe pero está
    # anidado bajo el menú "Reportes" (mismo patrón que GT en
    # _abrir_reporte_gt) — hay que abrir "Reportes" primero.
    for sel in ('a:has-text("Reportes")', 'button:has-text("Reportes")',
                'li:has-text("Reportes")', 'span:has-text("Reportes")'):
        try:
            await page.click(sel, timeout=3_000)
            await page.wait_for_timeout(700)
            break
        except Exception:
            continue
    for sel in ('a:has-text("Consultar Receta")', 'button:has-text("Consultar Receta")',
                ':text("Consultar Receta")'):
        try:
            await page.click(sel, timeout=3_000)
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1_000)
            break
        except Exception:
            continue

    for isel in CANDIDATOS_INPUT_RECETA:
        if await page.query_selector(isel):
            print(f"  Formulario de consulta encontrado (input: {isel})")
            return True

    print("  [ERROR] No encontré la pantalla de consulta de receta con los candidatos actuales.")
    if debug:
        await AS._dump_formulario(page, "consulta-receta-no-encontrada")
        await page.screenshot(path=str(AS.MAESTRO_DIR / "debug_consulta_receta_no_encontrada.png"))
    return False


async def _descargar_una(page, n_receta, dest_path, debug):
    input_sel = None
    for sel in CANDIDATOS_INPUT_RECETA:
        if await page.query_selector(sel):
            input_sel = sel
            break
    if not input_sel:
        print(f"  [ERROR] {n_receta}: no encontré el campo de búsqueda.")
        return False

    # Confirmado en vivo 23-07-2026: el formulario tiene pestañas (Run / Nº de
    # Receta / Cta. Cte. / DAU / Mis Recetas Emitidas) y por defecto viene
    # activa la de Run — el campo #numero_receta existe en el DOM pero está
    # oculto hasta clickear la pestaña "Nº de Receta".
    for sel in ('a:has-text("Nº de Receta")', 'button:has-text("Nº de Receta")',
                'li:has-text("Nº de Receta")', ':text("Nº de Receta")'):
        try:
            await page.click(sel, timeout=3_000)
            break
        except Exception:
            continue
    try:
        await page.wait_for_selector(input_sel, state="visible", timeout=8_000)
    except Exception:
        print(f"  [AVISO] {n_receta}: el campo sigue sin ser visible tras clickear la pestaña. "
              f"Volcando estado para diagnosticar.")
        await AS._dump_formulario(page, f"receta-{n_receta}-input-no-visible")
        await page.screenshot(path=str(AS.MAESTRO_DIR / f"debug_receta_{n_receta}_no_visible.png"))
        return False
    await page.fill(input_sel, str(n_receta))
    # Varias pestañas comparten el texto "Buscar" — se acota al botón visible.
    await AS._click_primero(page, ('button:has-text("Buscar"):visible', 'a:has-text("Buscar"):visible'), "Buscar")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1_500)

    if debug:
        await AS._dump_formulario(page, f"consulta-receta-{n_receta}")
        await page.screenshot(path=str(AS.MAESTRO_DIR / f"debug_receta_{n_receta}.png"))

    # Confirmado en vivo 23-07-2026: los botones "Imprimir*" del sitio abren
    # el diálogo de impresión NATIVO del sistema operativo (window.print()),
    # que Playwright no puede controlar ni capturar. En vez de pelear con
    # eso, se usa page.pdf() (motor de Chromium vía CDP) directo sobre la
    # página ya cargada — funciona en modo headed y trae exactamente lo que
    # se ve en pantalla (paciente + tabla de prescripción), sin depender de
    # ningún botón del sitio.
    texto = await page.evaluate("() => document.body.innerText")
    if "Información Receta Número" not in texto:
        print(f"  [ERROR] {n_receta}: la búsqueda no devolvió una receta válida.")
        return False
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    await page.pdf(path=str(dest_path))
    print(f"  ✓ {dest_path.name}  ({dest_path.stat().st_size // 1024:,} KB)")
    return True


async def _main_async(estab, recetas, debug):
    carpeta_local = _CARPETA_LOCAL.get(estab.upper())
    if not carpeta_local:
        print(f"[ERROR] '{estab}' no está en el mapeo de establecimientos.")
        return
    base_dir = os.path.join(GT_SOLICITUDES_DIR, carpeta_local, "Revisión de Solicitudes")
    os.makedirs(base_dir, exist_ok=True)

    from playwright.async_api import async_playwright
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, args=["--start-maximized"])
        if AS.SESSION_FILE.exists():
            context = await browser.new_context(accept_downloads=True, storage_state=str(AS.SESSION_FILE))
            print("(Sesión guardada encontrada — puede que no necesites logarte)")
        else:
            context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        print("Logéate en SSASUR (tienes 5 minutos)...")
        await page.goto(AS.DASHBOARD_URL)
        await page.wait_for_selector('button:has-text("ABASTECIMIENTO"), div:has-text("ABASTECIMIENTO")',
                                      timeout=AS.TIMEOUT_LOGIN)
        await context.storage_state(path=str(AS.SESSION_FILE))
        print("✓ Sesión detectada")

        await AS.entrar_receta(page)
        if not await _abrir_consulta_receta(page, debug):
            print("\nNo se pudo ubicar la pantalla de consulta. Revisa el volcado/screenshot"
                  " (si corriste con --debug) y avisa para ajustar los candidatos de URL/selectores.")
            await browser.close()
            return

        ok, fallidos = 0, []
        for i, (n_receta, paciente) in enumerate(recetas):
            if i > 0:
                # Cada búsqueda navega a la página de impresión — hay que
                # volver a la pantalla de consulta antes de la siguiente.
                if not await _abrir_consulta_receta(page, debug):
                    print(f"  [ERROR] {n_receta}: no pude volver a la pantalla de consulta.")
                    fallidos.append(n_receta)
                    continue
            nombre_archivo = re.sub(r"\s+", "", paciente) or "Paciente"
            dest = Path(os.path.join(base_dir, f"Receta_{n_receta}_{nombre_archivo}.pdf"))
            print(f"\nReceta {n_receta} ({paciente})...")
            if await _descargar_una(page, n_receta, dest, debug):
                ok += 1
            else:
                fallidos.append(n_receta)

        print(f"\n{ok}/{len(recetas)} PDF descargados. Fallidos: {fallidos or 'ninguno'}")
        await browser.close()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--estab", required=True)
    ap.add_argument("--feedback", help="Ruta a un Feedback_Solicitud_*.xlsx generado por revision_solicitudes.py")
    ap.add_argument("--recetas", help="Lista de N° de receta separados por coma (alternativa a --feedback)")
    ap.add_argument("--debug", action="store_true", help="Vuelca formularios y screenshots para ajustar selectores")
    a = ap.parse_args()

    if a.feedback:
        recetas = _leer_recetas_de_feedback(a.feedback)
    elif a.recetas:
        recetas = [(n.strip(), "") for n in a.recetas.split(",") if n.strip()]
    else:
        print("[ERROR] Pasa --feedback o --recetas.")
        return
    if not recetas:
        print("No hay recetas para procesar.")
        return
    print(f"{len(recetas)} receta(s) a descargar.")
    asyncio.run(_main_async(a.estab, recetas, a.debug))


if __name__ == "__main__":
    main()
