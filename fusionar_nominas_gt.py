#!/usr/bin/env python3
"""
fusionar_nominas_gt.py — Fusiona en UNA sola Nómina de Envío las recetas que
llegaron por las dos vías separadas de Gestión Territorial para el mismo
destino y la misma fecha:

  1. La automática (`cruce_gt.py --generar`, recetas con Estab. Destino
     etiquetado en SSASUR) -> "<slug(destino)>_Planilla.xlsx"
  2. La manual (`agregar_gt_manual.py`, invocada por el QF a mano, recetas
     SIN destino en SSASUR) -> "Nomina_Manual_<destino>_<fecha>.xlsx"

Hasta ahora ningún script leía el resultado del otro, así que cuando un
destino tenía recetas de ambos tipos el mismo día quedaban en 2 archivos
separados y había que fusionarlos a mano antes de enviar (caso real
25-08-2026: CESFAM Freire, automática=4 pacientes + manual=24 -> 28 reales).

Este script NO decide qué receta va en la nómina (eso lo hace SSASUR o el
QF vía agregar_gt_manual.py) — solo LEE los dos archivos ya generados y
escribe uno solo, con el mismo formato (`generar.hoja_funcionarios`),
SIEMPRE bajo el nombre oficial "<slug(destino)>_Planilla.xlsx" (el que
espera quien recibe en destino) — incluso si ese día solo hubo nómina
"manual" (backlog GT o registro a mano), se renombra igual. El .xlsx
"Nomina_Manual_..." y su letrero quedan borrados una vez fusionados: es
solo un archivo de staging interno, nunca el entregable final — "manual"
ya no es una distinción válida de cara a Gestión Territorial (corregido
07-09-2026, ver AUTO_SSASUR.py PASO 5c2). Si hay una receta repetida en
ambos archivos (no debería pasar — dedup por N° Receta en cada pipeline),
se conserva la versión manual y se avisa.

Uso:
  py fusionar_nominas_gt.py --destino "CESFAM FREIRE"                  # hoy
  py fusionar_nominas_gt.py --destino "CESFAM FREIRE" --fecha 2026-08-25
  py fusionar_nominas_gt.py --todos                                    # todos los destinos conocidos, fecha de hoy
"""
import argparse
import datetime
import glob
import os
import re

import gt_maestro as GM
import generar as G
from agregar_gt_manual import _CARPETA_LOCAL, _leer_nomina_existente

MAESTRO_DIR = os.path.dirname(os.path.abspath(__file__))
GT_DIR = os.path.join(os.path.dirname(MAESTRO_DIR), "04_Farmacia_Gestion_Territorial")

_RE_PERIODO_DE = re.compile(r"^\s*(\d+)\s+de\s+(\d+)\s*$", re.IGNORECASE)


def _normalizar_periodo(p):
    """agregar_gt_manual.py a veces trae el período como 'N de M' (tipeado a
    mano en la solicitud) mientras que el reporte automático de SSASUR usa
    'N/M' — se normaliza a 'N/M' para que la nómina fusionada quede
    uniforme (detectado 25-08-2026 comparando contra la nómina real de
    CESFAM Freire que el usuario armó a mano)."""
    m = _RE_PERIODO_DE.match(str(p or ""))
    return f"{m.group(1)}/{m.group(2)}" if m else p


def _limpiar_manual(ruta_manual, dirp):
    """Borra el .xlsx 'Manual' (staging interno de agregar_gt_manual.py) y su
    letrero PDF una vez que sus filas ya quedaron en el archivo oficial —
    nada rotulado 'manual' debe sobrevivir como entregable final."""
    try:
        os.remove(ruta_manual)
    except OSError:
        pass
    base = os.path.splitext(os.path.basename(ruta_manual))[0]
    letrero_pdf = os.path.join(dirp, f"{base}_Letrero.pdf")
    if os.path.exists(letrero_pdf):
        try:
            os.remove(letrero_pdf)
        except OSError:
            pass


def _carpeta_destino(destino, fecha):
    carpeta_local = _CARPETA_LOCAL.get(destino)
    if not carpeta_local:
        raise SystemExit(
            f"[ERROR] '{destino}' no está en el mapeo _CARPETA_LOCAL de agregar_gt_manual.py — "
            "usa el valor EXACTO tal como aparece en gt_maestro.xlsx (columna Establecimiento de destino)."
        )
    carpeta_mes = f"{GM.MESES_ES[fecha.month - 1]} {fecha.year}"
    carpeta_fecha = fecha.strftime("%d-%m-%Y")
    return os.path.join(GT_DIR, carpeta_local, "Nóminas de Envío", carpeta_mes, carpeta_fecha)


def fusionar(destino, fecha):
    dirp = _carpeta_destino(destino, fecha)
    if not os.path.isdir(dirp):
        print(f"{destino}: no existe la carpeta de hoy ({os.path.basename(dirp)}) — nada que fusionar.")
        return None

    # Detección por glob, NO por reconstruir el nombre esperado a partir de
    # slug(destino)/_norm(destino) — el string "destino" tiene variantes
    # (con/sin punto, "GORBEA DSM" vs "DSM GORBEA", etc.) que producen un
    # nombre de archivo distinto al que realmente generó cada pipeline.
    # Reconstruir el nombre a mano puede fallar en silencio (busca un
    # archivo que no existe -> "nada que fusionar" siendo falso). Mirar la
    # carpeta real evita ese desvío.
    candidatos_auto = sorted(f for f in glob.glob(os.path.join(dirp, "*_Planilla.xlsx"))
                              if not os.path.basename(f).startswith("Nomina_Manual"))
    candidatos_manual = sorted(f for f in glob.glob(os.path.join(dirp, "Nomina_Manual_*.xlsx"))
                                if "_Letrero" not in os.path.basename(f))

    if len(candidatos_auto) > 1 or len(candidatos_manual) > 1:
        print(f"  [AVISO] {destino}: hay más de un archivo candidato en {dirp} — "
              f"reviso a mano, NO fusiono para no elegir mal: "
              f"auto={candidatos_auto} manual={candidatos_manual}")
        return None

    ruta_auto = candidatos_auto[0] if candidatos_auto else os.path.join(dirp, f"{G.slug(destino)}_Planilla.xlsx")
    ruta_manual = candidatos_manual[0] if candidatos_manual else None
    hay_auto = bool(candidatos_auto)
    hay_manual = bool(candidatos_manual)

    if not hay_auto and not hay_manual:
        print(f"{destino}: nada que fusionar hoy (no hay ni automática ni manual en {dirp}).")
        return None
    if hay_auto and not hay_manual:
        print(f"{destino}: solo automática — nada que fusionar ({os.path.basename(ruta_auto)}).")
        return ruta_auto

    if hay_manual and not hay_auto:
        # Solo hay "manual" (backlog GT o registro a mano) — igual se renombra
        # al nombre oficial: el entregable final NUNCA debe quedar rotulado
        # "Manual" (07-09-2026: la QF reportó que las nóminas de Gestión
        # Territorial le seguían llegando así pese a que ya no aplica esa
        # distinción de cara al establecimiento destino).
        regs = list({r["receta"]: r for r in _leer_nomina_existente(ruta_manual)}.values())
        for r in regs:
            r["periodo"] = _normalizar_periodo(r.get("periodo"))
        from openpyxl import Workbook
        wb = Workbook()
        titulo = f"GESTIÓN TERRITORIAL - {destino.upper()}"
        subtitulo = f"Origen: Farmacia Hospital de Pitrufquén   |   Destino: {destino}   |   Fecha de entrega: {fecha.strftime('%d/%m/%Y')}"
        G.hoja_funcionarios(wb, regs, destino, titulo, subtitulo, modo="todos")
        wb.save(ruta_auto)
        _limpiar_manual(ruta_manual, dirp)
        print(f"{destino}: {len(regs)} paciente(s) -> {os.path.basename(ruta_auto)} "
              f"(renombrada desde {os.path.basename(ruta_manual)} — ya no queda 'manual')")
    else:
        regs_auto = {r["receta"]: r for r in _leer_nomina_existente(ruta_auto)}
        regs_manual = {r["receta"]: r for r in _leer_nomina_existente(ruta_manual)}
        choque = set(regs_auto) & set(regs_manual)
        ya_fusionada = bool(choque) and choque == set(regs_manual)
        if ya_fusionada:
            # Re-corrida del mismo día: la Planilla automática ya es el resultado
            # de una fusión anterior (la sobreescribimos con las 2 fuentes) y
            # contiene TODAS las recetas manuales — no hay que reescribir la
            # planilla, pero igual se asegura el letrero más abajo.
            print(f"{destino}: ya estaba fusionada ({len(regs_auto)} pacientes) — no se repite la planilla.")
            regs = list(regs_auto.values())
        else:
            if choque:
                print(f"  [AVISO] {len(choque)} receta(s) presentes en AMBOS archivos (no debería pasar) — "
                      f"se conserva la versión manual: {sorted(choque)}")
            regs = list({**regs_auto, **regs_manual}.values())
            for r in regs:
                r["periodo"] = _normalizar_periodo(r.get("periodo"))

            from openpyxl import Workbook
            wb = Workbook()
            titulo = f"GESTIÓN TERRITORIAL - {destino.upper()}"
            subtitulo = f"Origen: Farmacia Hospital de Pitrufquén   |   Destino: {destino}   |   Fecha de entrega: {fecha.strftime('%d/%m/%Y')}"
            G.hoja_funcionarios(wb, regs, destino, titulo, subtitulo, modo="todos")
            wb.save(ruta_auto)
            print(f"{destino}: fusionado -> {os.path.basename(ruta_auto)} "
                  f"({len(regs)} pacientes = {len(regs_auto)} automática + {len(regs_manual)} manual"
                  f"{f' - {len(choque)} choque' if choque else ''})")
        _limpiar_manual(ruta_manual, dirp)

    # El letrero se (re)genera SIEMPRE que se toca este destino — no solo si
    # hay refrigerado — mismo criterio que el pipeline automático normal
    # (generar.py llama a letrero() para todo destino; "lleva" solo decide
    # si el aviso ❄ aparece, no si el letrero existe).
    lleva = any((r.get("refrigerado") or "").strip() for r in regs)
    detalle_refri = G._detalle_refrigerados(regs, {}) if lleva else []
    ruta_letrero = os.path.join(dirp, f"{G.slug(destino)}_Letrero.xlsx")
    G.FECHA = fecha.strftime("%d/%m/%Y")
    G.letrero(destino, lleva, ruta_letrero, detalle_refri)
    if G.to_pdf(ruta_letrero, dirp):
        try:
            os.remove(ruta_letrero)
        except OSError:
            pass
        print(f"  Letrero -> {os.path.basename(ruta_letrero).replace('.xlsx', '.pdf')}")

    return ruta_auto


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--destino", help='Destino exacto tal como está en gt_maestro.xlsx, ej. "CESFAM FREIRE"')
    ap.add_argument("--todos", action="store_true", help="Fusiona todos los destinos conocidos (_CARPETA_LOCAL)")
    ap.add_argument("--fecha", help="YYYY-MM-DD (default: hoy)")
    a = ap.parse_args()
    if not a.destino and not a.todos:
        print("[ERROR] Pasa --destino \"<DESTINO>\" o --todos.")
        return
    fecha = datetime.date.fromisoformat(a.fecha) if a.fecha else datetime.date.today()
    if a.todos:
        # _CARPETA_LOCAL tiene 2 alias para varios destinos (mismo lugar,
        # distinto orden de palabra) — deduplicar por carpeta para no
        # procesar la misma carpeta 2 veces con títulos distintos.
        vistas = set()
        destinos = []
        for d, carpeta in _CARPETA_LOCAL.items():
            if carpeta not in vistas:
                vistas.add(carpeta)
                destinos.append(d)
    else:
        destinos = [a.destino]
    for destino in destinos:
        fusionar(destino, fecha)


if __name__ == "__main__":
    main()
