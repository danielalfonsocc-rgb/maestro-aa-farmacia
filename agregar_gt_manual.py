#!/usr/bin/env python3
"""
agregar_gt_manual.py — Registra a mano recetas que SSASUR NO etiqueta con un
establecimiento de destino (el campo "Gestión Territorial Despacho" viene
vacío), pero que igual se están despachando. Como no traen destino, no
aparecen en el reporteGestionTerritorial_*.xlsx ni se sincronizan solas con
gt_maestro.sincronizar_gt_report() — hay que registrarlas a mano una vez.

Hace dos cosas con cada fila del CSV:
  1. Las agrega/actualiza en el maestro (gt_maestro.upsert_receta_maestro) —
     quedan trackeadas igual que todo lo demás: color por establecimiento,
     MAYÚSCULAS, Historial si cambia el Estado, Sheets/Drive al republicar.
  2. Genera la Nómina de Envío (mismo formato que usa el skill_gt para las
     planillas automáticas) agrupada por establecimiento, guardada en
     04_Farmacia_Gestion_Territorial/<ESTABLECIMIENTO>/Nóminas de Envío/

Uso:
  1. Llena una copia de _gt_manual_plantilla.csv con las recetas
     (una fila por receta). Columnas:
       receta,paciente,rut,destino,periodo,especialidad,n_presc,telefono,pendiente
     - destino: nombre del establecimiento tal como quieres que aparezca
       (ej. "CESFAM TEODORO SCHMIDT", "HOSPITAL TOLTEN").
     - pendiente: opcional, texto libre (ej. medicamento que falta).

  2. py agregar_gt_manual.py --csv mi_lote.csv --dry-run   (preview)
     py agregar_gt_manual.py --csv mi_lote.csv             (aplica + genera nóminas)
"""
import argparse
import csv
import datetime
import os
import sys

import openpyxl

import gt_maestro as GM
import generar as G

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MAESTRO_DIR = os.path.dirname(os.path.abspath(__file__))
GT_SOLICITUDES_DIR = os.path.join(os.path.dirname(MAESTRO_DIR), "04_Farmacia_Gestion_Territorial")

# Mismo mapeo que publicar_drive.py — nombre de carpeta local (gt_maestro)
# vs. nombre de carpeta ya usado en Drive (convención SSASUR/skill_gt).
# Incluye AMBOS órdenes de palabra ("GORBEA DSM" y "DSM GORBEA", etc.) porque
# la columna "Establecimiento de destino" del maestro gt_maestro usa
# "TIPO LUGAR" (DSM GORBEA, HOSPITAL LONCOCHE) mientras que esta convención
# histórica (Drive/skill_gt) usa "LUGAR TIPO" (GORBEA DSM, LONCOCHE HOSP) —
# confirmado 23-07-2026 al revisar los valores reales del maestro.
_CARPETA_LOCAL = {
    "CESFAM FREIRE": "CESFAM_FREIRE",
    "CESFAM HUALPIN": "CESFAM_HUALPIN",
    "CESFAM QUEPE": "CESFAM_QUEPE",
    "CESFAM TEODORO SCHMIDT": "CESFAM_TEODORO_SCHMIDT",
    "GORBEA DSM": "DSM_GORBEA",
    "DSM GORBEA": "DSM_GORBEA",
    "LONCOCHE DSM": "DSM_LONCOCHE",
    "DSM LONCOCHE": "DSM_LONCOCHE",
    "TOLTEN DSM": "DSM_TOLTEN",
    "DSM TOLTEN": "DSM_TOLTEN",
    "GORBEA HOSP": "HOSPITAL_GORBEA",
    "HOSPITAL GORBEA": "HOSPITAL_GORBEA",
    "LONCOCHE HOSP": "HOSPITAL_LONCOCHE",
    "HOSPITAL LONCOCHE": "HOSPITAL_LONCOCHE",
    "TOLTEN HOSP": "HOSPITAL_TOLTEN",
    "HOSPITAL TOLTEN": "HOSPITAL_TOLTEN",
    "PSR COMUY": "PSR_COMUY",
    "PSR QUEULE": "PSR_QUEULE",
}


def _leer_csv(path):
    filas = []
    with open(path, encoding="utf-8-sig", newline="") as fh:
        for row in csv.DictReader(fh):
            receta = (row.get("receta") or "").strip()
            if not receta:
                continue
            filas.append({
                "receta": receta,
                "paciente": (row.get("paciente") or "").strip(),
                "rut": (row.get("rut") or "").strip(),
                "destino": (row.get("destino") or "").strip(),
                "periodo": (row.get("periodo") or "").strip(),
                "especialidad": (row.get("especialidad") or "").strip(),
                "n_presc": (row.get("n_presc") or "1").strip(),
                "telefono": (row.get("telefono") or "").strip(),
                "pendiente": (row.get("pendiente") or "").strip(),
            })
    return filas


def _generar_nomina(destino, filas_destino, fecha_hoy):
    """Genera un .xlsx de Nómina de Envío (mismo formato que skill_gt) para
    un establecimiento, con las filas manuales de esta corrida."""
    regs = [{
        "receta": f["receta"], "paciente": f["paciente"], "run": f["rut"],
        "especialidad": f["especialidad"], "periodo": f["periodo"],
        "n_presc": int(f["n_presc"] or 1), "pendiente": f["pendiente"],
    } for f in filas_destino]

    wb = openpyxl.Workbook()
    titulo = f"GESTIÓN TERRITORIAL — {destino}"
    subtitulo = (f"Origen: Farmacia Hospital de Pitrufquén   |   Destino: {destino}   |   "
                 f"Nómina manual (sin destino en SSASUR) — {fecha_hoy.strftime('%d/%m/%Y')}")
    G.hoja_funcionarios(wb, regs, destino, titulo, subtitulo)

    # "Nóminas de Envío" se organiza en dos niveles: carpeta por mes
    # ("JULIO 2026") y adentro carpeta por fecha de extracción (DD-MM-YYYY) —
    # mismo esquema que "Revisión de Solicitudes" (ver
    # revision_solicitudes._carpeta_salida). La nómina manual queda ahí, con
    # prefijo "Nomina_Manual_" para distinguirla de las que baja skill_gt.
    carpeta_mes = f"{GM.MESES_ES[fecha_hoy.month - 1]} {fecha_hoy.year}"
    carpeta_fecha = fecha_hoy.strftime("%d-%m-%Y")
    carpeta_local = _CARPETA_LOCAL.get(destino)
    if carpeta_local:
        destino_dir = os.path.join(GT_SOLICITUDES_DIR, carpeta_local, "Nóminas de Envío", carpeta_mes, carpeta_fecha)
    else:
        destino_dir = os.path.join(GT_SOLICITUDES_DIR, "_sin_carpeta_conocida", carpeta_mes, carpeta_fecha)
        print(f"  [AVISO] '{destino}' no está en el mapeo de carpetas — guardando en {destino_dir}")
    os.makedirs(destino_dir, exist_ok=True)

    nombre = f"Nomina_Manual_{GM._norm(destino)}_{fecha_hoy.strftime('%Y-%m-%d')}.xlsx"
    ruta = os.path.join(destino_dir, nombre)
    wb.save(ruta)
    return ruta


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--csv", required=True, help="CSV con las recetas sin destino en el sistema")
    ap.add_argument("--estado", default="EN PREPARACIÓN",
                     help="Estado a asignar (default: EN PREPARACIÓN)")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    filas = _leer_csv(a.csv)
    if not filas:
        print(f"[ERROR] {a.csv} no tiene filas válidas (falta la columna 'receta' o está vacía).")
        return

    sin_destino = [f["receta"] for f in filas if not f["destino"]]
    if sin_destino:
        print(f"[ERROR] Estas recetas no tienen 'destino' en el CSV — complétalo antes de correr: {sin_destino}")
        return

    print(f"{'[DRY-RUN] ' if a.dry_run else ''}{len(filas)} receta(s) en {a.csv}")

    wb, path = GM.cargar_maestro()
    hoy = datetime.date.today()
    hojas_tocadas = {}
    for f in filas:
        receta_dict = {k: v for k, v in f.items() if k != "receta" and v} | {"receta": f["receta"]}
        if a.dry_run:
            ws_existente = GM.buscar_receta_en_maestro(wb, f["receta"])
            print(f"  {'actualizaría' if ws_existente else 'nueva'} -> {f['receta']} | {f['paciente']} | destino={f['destino']}")
            continue
        ws, resultado, _ = GM.upsert_receta_maestro(wb, receta_dict, estado=a.estado, fecha_fallback=hoy)
        hojas_tocadas[ws.title] = ws
        print(f"  {resultado} -> {ws.title} | {f['receta']} | {f['paciente']}")

    if a.dry_run:
        print("\nNada escrito (dry-run). Nóminas no generadas en dry-run.")
        return

    for ws in hojas_tocadas.values():
        GM.aplicar_formato_maestro(ws)
    GM.formatear_historial(wb)
    GM.guardar(wb, path)
    print(f"\nGuardado: {path}")

    por_destino = {}
    for f in filas:
        por_destino.setdefault(f["destino"], []).append(f)
    print("\nGenerando Nóminas de Envío:")
    for destino, filas_destino in por_destino.items():
        ruta = _generar_nomina(destino, filas_destino, hoy)
        print(f"  {destino}: {len(filas_destino)} receta(s) -> {ruta}")


if __name__ == "__main__":
    main()
