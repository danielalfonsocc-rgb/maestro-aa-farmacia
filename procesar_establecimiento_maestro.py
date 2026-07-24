#!/usr/bin/env python3
"""
procesar_establecimiento_maestro.py — Para establecimientos que NO envían
solicitud propia (ej. DSM Loncoche): en vez de esperar una planilla o PDFs
del establecimiento, saca el listado de pacientes directamente del maestro
de Gestión Territorial (gt_maestro), del último mes con filas para ese
establecimiento, y genera el mismo PDF consolidado + Feedback que el resto
del workflow.

Uso:
  py procesar_establecimiento_maestro.py --estab "DSM LONCOCHE" --dry-run
  py procesar_establecimiento_maestro.py --estab "DSM LONCOCHE"
"""
import argparse
import asyncio
import datetime
import os
import re
import sys

from pypdf import PdfWriter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import gt_maestro as GM
from agregar_gt_manual import _CARPETA_LOCAL, GT_SOLICITUDES_DIR
from revision_solicitudes import _escribir_feedback, _carpeta_salida
import descargar_recetas_pdf as DRP


def _filas_establecimiento_ultimo_mes(estab):
    """Recorre las hojas de mes del maestro (más reciente primero) y
    devuelve las filas del PRIMER mes que tenga alguna fila para ese
    establecimiento — "el último mes que se envió"."""
    wb, _ = GM.cargar_maestro()
    estab_norm = GM._norm(estab)
    hojas_mes = [s for s in wb.sheetnames if s != GM.HOJA_HISTORIAL]
    # Orden cronológico por nombre de hoja ("ENERO 2026" ... "JULIO 2026")
    def _clave(nombre):
        mes, anio = nombre.rsplit(" ", 1)
        return (int(anio), GM.MESES_ES.index(mes) if mes in GM.MESES_ES else -1)
    hojas_mes.sort(key=_clave, reverse=True)

    headers = GM.HEADERS
    col_receta = headers.index("Nº de receta")
    col_paciente = headers.index("Paciente")
    col_rut = headers.index("Rut")
    col_destino = headers.index("Establecimiento de destino")
    col_periodo = headers.index("Periodo Receta")
    col_especialidad = headers.index("Especialidad")

    for sh in hojas_mes:
        ws = wb[sh]
        filas = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[col_receta]:
                continue
            if GM._norm(row[col_destino]) != estab_norm:
                continue
            filas.append({
                "receta": str(row[col_receta]).strip(),
                "paciente": str(row[col_paciente] or "").strip(),
                "rut": str(row[col_rut] or "").strip(),
                "periodo": str(row[col_periodo] or "").strip(),
                "especialidad": str(row[col_especialidad] or "").strip(),
            })
        if filas:
            return sh, filas
    return None, []


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--estab", required=True, help='Establecimiento tal como aparece en el maestro, ej. "DSM LONCOCHE"')
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    carpeta_local = _CARPETA_LOCAL.get(a.estab.upper())
    if not carpeta_local:
        print(f"[ERROR] '{a.estab}' no está en el mapeo de establecimientos.")
        return
    base_dir = os.path.join(GT_SOLICITUDES_DIR, carpeta_local, "Revisión de Solicitudes")
    if not os.path.isdir(base_dir):
        print(f"[ERROR] No existe {base_dir}")
        return

    hoja, filas = _filas_establecimiento_ultimo_mes(a.estab)
    if not filas:
        print(f"[{a.estab}] No encontré filas en ningún mes del maestro.")
        return
    print(f"[{a.estab}] {len(filas)} receta(s) en el maestro, hoja '{hoja}' (mes más reciente con datos).")

    alertas = GM.detectar_alertas_mismo_rut([
        {**f, "medicamentos": []} for f in filas
    ])
    alerta_por_receta = {}
    for al in alertas:
        for r in al["recetas"]:
            alerta_por_receta.setdefault(r, []).append(al["nota"])
    print(f"  {len(alertas)} alerta(s) de posible duplicado/ambigüedad detectada(s).")

    filas_feedback = [(f, f"Del maestro GT ({hoja}) — establecimiento sin solicitud propia") for f in filas]

    fecha_hoy = datetime.date.today()
    fecha_str = fecha_hoy.strftime("%Y-%m-%d")

    if a.dry_run:
        print(f"  [DRY-RUN] Generaría Feedback_Solicitud_{carpeta_local}_{fecha_str}.xlsx "
              f"({len(filas_feedback)} fila(s)) y descargaría {len(filas)} PDF de recetas. No se hace nada.")
        return

    salida_dir = _carpeta_salida(base_dir, fecha_hoy)

    feedback_path = os.path.join(salida_dir, f"Feedback_Solicitud_{carpeta_local}_{fecha_str}.xlsx")
    _escribir_feedback(feedback_path, a.estab, fecha_hoy, filas_feedback, alerta_por_receta)
    print(f"  Guardado: {feedback_path}")

    recetas = [(f["receta"], f["paciente"]) for f in filas]
    print(f"\n  Descargando {len(recetas)} PDF de recetas desde SSASUR (necesitas loguearte)...")
    asyncio.run(DRP._main_async(a.estab, recetas, debug=False))

    # descargar_recetas_pdf.py también resuelve su propio salida_dir vía
    # _carpeta_salida(base_dir, hoy) — mismo día, misma carpeta, así que los
    # PDF individuales ya quedan en el subdirectorio MES/FECHA correcto.
    individuales = [os.path.join(salida_dir, f"Receta_{r}_{re.sub(r'\s+', '', p) or 'Paciente'}.pdf")
                    for r, p in recetas]
    individuales = [p for p in individuales if os.path.isfile(p)]
    if individuales:
        combinado = os.path.join(salida_dir, f"Recetas_Combinadas_{carpeta_local}_{fecha_str}.pdf")
        writer = PdfWriter()
        for p in individuales:
            writer.append(p)
        with open(combinado, "wb") as fh:
            writer.write(fh)
        print(f"  Guardado: {combinado}  ({len(individuales)} receta(s))")

        pdfs_dir = os.path.join(salida_dir, "PDFs individuales")
        os.makedirs(pdfs_dir, exist_ok=True)
        for p in individuales:
            os.replace(p, os.path.join(pdfs_dir, os.path.basename(p)))
        print(f"  {len(individuales)} PDF individual(es) movido(s) a 'PDFs individuales'.")
    else:
        print("  [AVISO] No se descargó ningún PDF — no se generó el combinado.")


if __name__ == "__main__":
    main()
