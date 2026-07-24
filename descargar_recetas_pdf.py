#!/usr/bin/env python3
"""
descargar_recetas_pdf.py — Descarga de SSASUR el PDF OFICIAL de cada receta
(por Nº) — "RECETA MÉDICA N°..." con firma del médico y checklist de
retiro, el mismo formato que produce el sitio al hacer clic en "Imprimir" —
para completar una Revisión de Solicitudes con los documentos reales, no
solo la lista de números que arma revision_solicitudes.py --xlsx.

Confirmado en vivo 23-07-2026, contrastado byte a byte contra el PDF de
referencia de Toltén (18-07-2026): el botón "Imprimir" del sitio abre el
diálogo de impresión NATIVO del sistema operativo (window.print()), que
Playwright no puede controlar — pero ese botón en realidad pide, por debajo,
el endpoint https://www.ssasur.cl/receta/impresion/pdf/<receta>/undefined,
que devuelve el PDF oficial directo. Se pide ese endpoint con la sesión ya
autenticada, sin necesidad de navegar ni clicar nada del formulario.

Uso:
  py descargar_recetas_pdf.py --estab "CESFAM FREIRE" --recetas 45858462,46045176 --debug
  py descargar_recetas_pdf.py --estab "CESFAM FREIRE" --feedback "C:/ruta/Feedback_Solicitud_CESFAM_FREIRE_2026-07-23.xlsx"
  py descargar_recetas_pdf.py --estab "CESFAM FREIRE" --rut "6.385.207-4"
      (busca en vivo las recetas vigentes de ese RUT — para cuando la receta
      es tan nueva que ni siquiera aparece en el CSV local; confirmado en
      vivo 23-07-2026 vía Consultar Receta -> pestaña Run, filtrando Estado
      distinto de ENTREGADA/CERRADA-INCOMPLETA/ANULADA)
"""
import argparse
import asyncio
import datetime
import os
import re
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import AUTO_SSASUR as AS
from agregar_gt_manual import _CARPETA_LOCAL, GT_SOLICITUDES_DIR
from revision_solicitudes import _carpeta_salida

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MAESTRO_DIR = os.path.dirname(os.path.abspath(__file__))


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


ESTADOS_NO_VIGENTES = {"ENTREGADA", "CERRADA / INCOMPLETA", "ANULADA"}


def _split_rut(rut):
    """'6.385.207-4' / '6385207-4' -> ('6385207', '4')."""
    limpio = re.sub(r"[.\s]", "", str(rut)).upper()
    numero, dv = limpio.rsplit("-", 1) if "-" in limpio else (limpio, "")
    return numero, dv


async def _buscar_vigentes_por_rut(page, rut, origen_filtro="PITRUFQUEN HOSP."):
    """Busca por RUN o ID Único de Paciente (pestaña activa por defecto en
    Consultar Receta — confirmado en vivo 23-07-2026) y devuelve las
    recetas NO entregadas del establecimiento de origen indicado — el caso
    de uso real: encontrar la receta más nueva de un RUT que SSASUR aún no
    despachó, cuando ni siquiera aparece todavía en el CSV local. La tabla
    de resultados viene ordenada por Nº de receta descendente (las más
    nuevas primero), por eso basta con mirar la primera página."""
    numero, dv = _split_rut(rut)
    await page.fill("#rut", numero)
    await page.fill("#dv", dv)
    await AS._click_primero(page, ('button:has-text("Buscar"):visible', 'a:has-text("Buscar"):visible'), "Buscar")
    await page.wait_for_load_state("networkidle")
    try:
        await page.wait_for_function(
            "() => { const t = document.querySelector('#tablaResultados'); "
            "return t && t.querySelectorAll('tbody tr').length >= 1 && !document.body.innerText.includes('Cargando'); }",
            timeout=20_000,
        )
    except Exception:
        print(f"  [AVISO] RUT {rut}: no cargó la tabla de resultados a tiempo.")
        return []
    await page.wait_for_timeout(500)

    filas = await page.evaluate(r"""() => {
      const t = document.querySelector('#tablaResultados');
      return [...t.querySelectorAll('tbody tr')].map(tr =>
        [...tr.querySelectorAll('td')].map(td => td.innerText.trim()));
    }""")
    vigentes = []
    for r in filas:
        if len(r) <= 9:
            continue
        n_receta, origen, estado = r[0], r[1], r[9]
        if not n_receta or not n_receta.isdigit():
            continue
        if estado.upper() in ESTADOS_NO_VIGENTES:
            continue
        if origen_filtro and origen_filtro.upper() not in origen.upper():
            continue
        vigentes.append((n_receta, estado))
    return vigentes


PDF_ENDPOINT = "https://www.ssasur.cl/receta/impresion/pdf/{n_receta}/undefined"


async def _descargar_una(page, n_receta, dest_path, debug):
    """Confirmado en vivo 23-07-2026 (contrastado byte a byte contra el
    formato de referencia de Toltén 18-07-2026): los botones "Imprimir*" del
    sitio abren el diálogo de impresión NATIVO del sistema operativo
    (window.print()), que Playwright no puede controlar. Pero ese botón
    dispara internamente una petición a un endpoint que SÍ devuelve el PDF
    oficial directo — "RECETA MÉDICA N°..." con firma del médico y checklist
    de retiro, no la vista web de consulta. Se pide ese endpoint directo con
    la sesión ya autenticada, sin necesidad de navegar el formulario."""
    resp = await page.context.request.get(PDF_ENDPOINT.format(n_receta=n_receta))
    if resp.status != 200 or "pdf" not in (resp.headers.get("content-type") or "").lower():
        print(f"  [ERROR] {n_receta}: el endpoint no devolvió un PDF (status {resp.status}).")
        if debug:
            body = await resp.body()
            (AS.MAESTRO_DIR / f"debug_receta_{n_receta}_respuesta.bin").write_bytes(body)
        return False
    body = await resp.body()
    if not body.startswith(b"%PDF"):
        print(f"  [ERROR] {n_receta}: la respuesta no es un PDF válido.")
        return False
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    dest_path.write_bytes(body)
    print(f"  ✓ {dest_path.name}  ({dest_path.stat().st_size // 1024:,} KB)")
    return True


async def _main_async(estab, recetas, debug, rut=None):
    carpeta_local = _CARPETA_LOCAL.get(estab.upper())
    if not carpeta_local:
        print(f"[ERROR] '{estab}' no está en el mapeo de establecimientos.")
        return
    base_dir = os.path.join(GT_SOLICITUDES_DIR, carpeta_local, "Revisión de Solicitudes")
    salida_dir = _carpeta_salida(base_dir, datetime.date.today())

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

        if rut:
            for sel in ('a:has-text("Reportes")', 'button:has-text("Reportes")'):
                try:
                    await page.click(sel, timeout=3_000)
                    await page.wait_for_timeout(700)
                    break
                except Exception:
                    continue
            for sel in ('a:has-text("Consultar Receta")', 'button:has-text("Consultar Receta")'):
                try:
                    await page.click(sel, timeout=3_000)
                    await page.wait_for_load_state("networkidle")
                    break
                except Exception:
                    continue
            print(f"\nBuscando recetas vigentes del RUT {rut} (origen PITRUFQUEN HOSP.)...")
            vigentes = await _buscar_vigentes_por_rut(page, rut)
            if not vigentes:
                print("  No se encontraron recetas vigentes (no ENTREGADA/CERRADA/ANULADA) para ese RUT.")
                await browser.close()
                return
            for n_receta, estado in vigentes:
                print(f"  Encontrada: receta {n_receta} — Estado: {estado}")
            recetas = list(recetas) + [(n, "") for n, _ in vigentes]

        if not recetas:
            print("No hay recetas para descargar.")
            await browser.close()
            return

        # _descargar_una pide el endpoint PDF directo — no navega el
        # formulario, así que no hace falta "volver a la pantalla de
        # consulta" entre una receta y otra como antes.
        ok, fallidos = 0, []
        for n_receta, paciente in recetas:
            nombre_archivo = re.sub(r"\s+", "", paciente) or "Paciente"
            dest = Path(os.path.join(salida_dir, f"Receta_{n_receta}_{nombre_archivo}.pdf"))
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
    ap.add_argument("--rut", help="Busca las recetas VIGENTES (no entregadas) de este RUT en vivo en SSASUR — "
                                   "para cuando la receta es tan nueva que todavía no está en el CSV local")
    ap.add_argument("--debug", action="store_true", help="Vuelca formularios y screenshots para ajustar selectores")
    a = ap.parse_args()

    if a.feedback:
        recetas = _leer_recetas_de_feedback(a.feedback)
    elif a.recetas:
        recetas = [(n.strip(), "") for n in a.recetas.split(",") if n.strip()]
    elif a.rut:
        recetas = []
    else:
        print("[ERROR] Pasa --feedback, --recetas o --rut.")
        return
    if not recetas and not a.rut:
        print("No hay recetas para procesar.")
        return
    if recetas:
        print(f"{len(recetas)} receta(s) a descargar.")
    asyncio.run(_main_async(a.estab, recetas, a.debug, rut=a.rut))


if __name__ == "__main__":
    main()
