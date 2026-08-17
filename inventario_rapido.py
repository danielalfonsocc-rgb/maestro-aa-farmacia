#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inventario_rapido.py — Instrumento de conteo + diferencia + pedido, SIN Consolidado
=====================================================================================
Versión liviana de `programacion_aa.py` para cuando NO se tiene acceso al
Consolidado_AA_MAESTRO.xlsx ni al historial de recetas (p.ej. en un
computador sin el repo Maestro AA clonado). Solo necesita los dos reportes
que se descargan directo de SSASUR y que cualquiera puede tener a mano:

  1. Reporte de Programación ("cantidad_de_productos_consumidos_en_centro_de
     _costo_farmacia*.xlsx", Reportes → Consumo por centro de costo →
     Centro de Costo = FARMACIA → Generar XLS)
       → Cantidad Programada / Cantidad Solicitada del ciclo.
  2. Reporte de stock ("reporte_de_stock_*.xlsx")
       → Stock Sistema de la bodega (por defecto BODEGA AT ABIERTA).

A diferencia de programacion_aa.py, este script NO calcula Consumo Promedio
Mensual ni Sugerencia de programación (eso requiere el historial completo de
recetas vía maestro_aa.py) — solo entrega el instrumento de conteo y, tras
contar físicamente, la Diferencia y una Cantidad a Pedir simple.

Uso — paso 1, generar el instrumento de conteo:
    py inventario_rapido.py --programacion reporte_prog.xlsx --stock reporte_stock.xlsx

    Genera Programacion_AA/Instrumento_Conteo_AA_<fecha>.xlsx. Imprímelo (o
    ábrelo en Excel/Sheets) y cuenta físicamente, anotando en la columna
    "Stock Real".

Uso — paso 2, aplicar el conteo:
    py inventario_rapido.py --aplicar-conteo Instrumento_Conteo_AA_<fecha>.xlsx
        (si ya llenaste la columna "Stock Real" directo en ese Excel)
  o
    py inventario_rapido.py --aplicar-conteo conteo.json
        conteo.json = {"MEDICAMENTO TAL COMO SALE EN LA PLANILLA": 123, ...}
        (si prefieres pasar los conteos como texto/JSON; se aplican sobre el
        Instrumento_Conteo_AA_*.xlsx más reciente, o el que indiques con
        --plantilla)

    Genera Programacion_AA/Resumen_Conteo_AA_<fecha_hora>.xlsx con:
      - Diferencia = Stock Real − Stock Sistema (discrepancia de inventario)
      - Cantidad a Pedir = max(0, objetivo − Stock Real), donde objetivo es
        Cantidad Programada (o Solicitada si no hay Programada). Es una
        fórmula simple de "reponer hasta lo programado" — no reemplaza el
        modelo SGLI completo (sgli.py), que sí usa consumo histórico.

Otras opciones:
    --bodega "NOMBRE"   Filtra el reporte de stock por otra bodega (default:
                         BODEGA AT ABIERTA — debe calzar con BODEGA_FISICA_AA
                         de programacion_aa.py). Si el nombre no aparece en
                         el archivo, el script lista las bodegas disponibles.

Sin llamadas a IA — solo pandas + openpyxl.
"""
import os
import sys
import json
import glob
import argparse
import datetime as dt

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

WORK_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, WORK_DIR)
from utils_aa import norm_erp, HOMOLOGACION, setup_stdout  # noqa: E402

setup_stdout()

OUT_DIR = os.path.join(WORK_DIR, 'Programacion_AA')
BODEGA_DEFAULT = 'BODEGA AT ABIERTA'  # debe calzar con BODEGA_FISICA_AA de programacion_aa.py

THIN  = Side(style='thin', color='DCDCDC')
BRD   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HFILL = PatternFill('solid', fgColor='D1FAF5')
HFONT = Font(bold=True, color='065F46', name='Arial', size=10)
PEDIR_FILL = PatternFill('solid', fgColor='FFE0B2')
PEDIR_FONT = Font(name='Arial', size=10, color='B45309', bold=True)
DIFF_FILL  = PatternFill('solid', fgColor='F4B3B3')
DIFF_FONT  = Font(name='Arial', size=10, color='7F1D1D', bold=True)
FALTA_FILL = PatternFill('solid', fgColor='E5E7EB')


def _key(nombre):
    n = norm_erp(nombre)
    return HOMOLOGACION.get(n, n)


# ─────────────── lectura de los dos reportes crudos ─────────────────────────

def _leer_programacion(ruta):
    """Lee el reporte 'Consumos por centro de costo' de SSASUR (Generar XLS).
    Fila 0 = título, fila 1 = metadata con el mes/año, fila 2 = encabezados.
    Devuelve {key: {'Medicamento': nombre_original, 'Cantidad Programada': x,
    'Cantidad Solicitada': y}} y el texto de metadata (mes/año) para mostrar."""
    meta = pd.read_excel(ruta, header=None, nrows=2, engine='openpyxl')
    meta_txt = str(meta.iloc[1, 0]).strip() if meta.shape[0] > 1 else ''

    df = pd.read_excel(ruta, header=2, engine='openpyxl')
    df = df.rename(columns=lambda c: str(c).strip())
    faltantes = [c for c in ('Producto', 'Total de Productos Programados', 'Productos Solicitado')
                 if c not in df.columns]
    if faltantes:
        print(f'[ERROR] El reporte de Programación no tiene las columnas esperadas {faltantes}.')
        print(f'  Columnas encontradas: {list(df.columns)}')
        print('  ¿Es el archivo correcto? (Reportes → Consumo por centro de costo → Generar XLS)')
        sys.exit(1)
    if 'Centro Costo' in df.columns:
        df = df[df['Centro Costo'].astype(str).str.strip().str.upper() == 'FARMACIA']
    df['Producto'] = df['Producto'].astype(str).str.strip()
    df = df[df['Producto'] != '']
    df['_key'] = df['Producto'].map(_key)
    df['Total de Productos Programados'] = pd.to_numeric(df['Total de Productos Programados'], errors='coerce')
    df['Productos Solicitado'] = pd.to_numeric(df['Productos Solicitado'], errors='coerce')

    prog = {}
    for row in df.to_dict('records'):
        k = row['_key']
        cur = prog.setdefault(k, {'Medicamento': row['Producto'], 'Cantidad Programada': 0.0, 'Cantidad Solicitada': 0.0})
        p = row['Total de Productos Programados']
        s = row['Productos Solicitado']
        if not pd.isna(p):
            cur['Cantidad Programada'] += p
        if not pd.isna(s):
            cur['Cantidad Solicitada'] += s
    return prog, meta_txt


def _leer_stock(ruta, bodega):
    """Lee un reporte_de_stock_*.xlsx crudo de SSASUR y devuelve
    {key: {'Medicamento': nombre_original, 'Stock Sistema': cantidad}} para
    la bodega indicada, sumando por lote. Si la bodega no aparece, lista las
    disponibles y aborta."""
    df = pd.read_excel(ruta, header=2, engine='openpyxl')
    df.columns = [str(c).strip() for c in df.columns]
    try:
        descr_col = next(c for c in df.columns if 'Descrip' in c)
        bod_col = next(c for c in df.columns if 'Bodega' in c)
    except StopIteration:
        print(f'[ERROR] El reporte de stock no tiene columnas de Descripción/Bodega reconocibles.')
        print(f'  Columnas encontradas: {list(df.columns)}')
        sys.exit(1)
    if 'Cantidad' not in df.columns:
        print(f'[ERROR] El reporte de stock no tiene columna "Cantidad". Columnas: {list(df.columns)}')
        sys.exit(1)

    df[bod_col] = df[bod_col].astype(str).str.strip().str.upper()
    bodega_norm = bodega.strip().upper()
    disponibles = sorted(df[bod_col].dropna().unique())
    if bodega_norm not in disponibles:
        print(f'[ERROR] No se encontró la bodega "{bodega}" en el reporte de stock.')
        print('  Bodegas disponibles en el archivo:')
        for b in disponibles:
            print(f'    - {b}')
        print('  Reintenta con --bodega "NOMBRE EXACTO"')
        sys.exit(1)

    sub = df[df[bod_col] == bodega_norm].copy()
    sub['Cantidad'] = pd.to_numeric(sub['Cantidad'], errors='coerce').fillna(0)
    sub['_key'] = sub[descr_col].astype(str).str.strip().map(_key)

    stock = {}
    for row in sub.to_dict('records'):
        k = row['_key']
        cur = stock.setdefault(k, {'Medicamento': row[descr_col].strip(), 'Stock Sistema': 0.0})
        cur['Stock Sistema'] += row['Cantidad']
    return stock


# ─────────────── modo 1: generar instrumento de conteo ──────────────────────

HDRS = [
    ('Medicamento',            46),
    ('Cantidad Programada',    17),
    ('Cantidad Solicitada',    17),
    ('Stock Sistema',          15),
    ('Stock Real',             13),
]


def generar(ruta_prog, ruta_stock, bodega):
    prog, meta_txt = _leer_programacion(ruta_prog)
    stock = _leer_stock(ruta_stock, bodega)

    universo = {}
    for k, v in prog.items():
        universo[k] = {'Medicamento': v['Medicamento'],
                        'Cantidad Programada': v['Cantidad Programada'],
                        'Cantidad Solicitada': v['Cantidad Solicitada'],
                        'Stock Sistema': None,
                        'Stock Real': None}
    n_solo_stock = 0
    for k, v in stock.items():
        if k in universo:
            universo[k]['Stock Sistema'] = v['Stock Sistema']
        else:
            n_solo_stock += 1
            universo[k] = {'Medicamento': v['Medicamento'],
                            'Cantidad Programada': None,
                            'Cantidad Solicitada': None,
                            'Stock Sistema': v['Stock Sistema'],
                            'Stock Real': None}

    filas = sorted(universo.values(), key=lambda f: f['Medicamento'])

    hoy = dt.date.today()
    os.makedirs(OUT_DIR, exist_ok=True)
    sal = os.path.join(OUT_DIR, f'Instrumento_Conteo_AA_{hoy.strftime("%Y%m%d")}.xlsx')
    sub = (f'Programación: {os.path.basename(ruta_prog)} ({meta_txt or "sin metadata de mes"})  ·  '
           f'Stock: {os.path.basename(ruta_stock)}  ·  Bodega: {bodega}')
    _escribir(sal, filas, 'INSTRUMENTO DE CONTEO — Bodega AA', sub, hoy, resumen=False)

    print(f'HOY = {hoy}')
    print(f'{len(filas)} medicamentos ({len(prog)} en Programación, {n_solo_stock} solo en Stock — sin programar)')
    print(f'\nExcel: {os.path.basename(sal)}  (carpeta {os.path.basename(OUT_DIR)}\\)')
    print('Imprime o abre esta planilla, cuenta físicamente y llena "Stock Real" a mano.')
    print(f'Luego corre: py inventario_rapido.py --aplicar-conteo "{os.path.basename(sal)}"')
    return sal


# ─────────────── modo 2: aplicar conteo → resumen con pedido ────────────────

def _mas_reciente_instrumento():
    cand = [f for f in glob.glob(os.path.join(OUT_DIR, 'Instrumento_Conteo_AA_*.xlsx'))
            if not os.path.basename(f).startswith('~$')]
    return max(cand, key=os.path.getmtime) if cand else None


def _cantidad_a_pedir(programada, solicitada, stock_real):
    if stock_real is None:
        return None
    objetivo = programada if programada not in (None,) and not (isinstance(programada, float) and pd.isna(programada)) else solicitada
    if objetivo is None or (isinstance(objetivo, float) and pd.isna(objetivo)):
        return None
    return max(0, round(objetivo - stock_real))


def aplicar_conteo(ruta, ruta_plantilla=None):
    hoy = dt.date.today()

    if ruta.lower().endswith('.xlsx'):
        # El propio instrumento, ya con "Stock Real" llenado a mano.
        plantilla = ruta
        wb = openpyxl.load_workbook(plantilla, data_only=True)
        ws = wb.active
        filas_base = []
        for row in ws.iter_rows(min_row=4, max_col=5):
            med_c, prog_c, sol_c, sbod_c, sreal_c = row
            if med_c.value is None:
                continue
            filas_base.append({
                'Medicamento': med_c.value,
                'Cantidad Programada': prog_c.value,
                'Cantidad Solicitada': sol_c.value,
                'Stock Sistema': sbod_c.value,
                'Stock Real': sreal_c.value,
            })
    else:
        # JSON medicamento→cantidad, aplicado sobre el instrumento más reciente.
        plantilla = ruta_plantilla or _mas_reciente_instrumento()
        if not plantilla:
            print('[ERROR] No hay ningún Instrumento_Conteo_AA_*.xlsx generado todavía.')
            print('  Corre primero: py inventario_rapido.py --programacion ... --stock ...')
            sys.exit(1)
        try:
            with open(ruta, encoding='utf-8') as fh:
                valores = json.load(fh)
        except (OSError, json.JSONDecodeError) as e:
            print(f'[ERROR] No se pudo leer {ruta} como JSON: {e}')
            sys.exit(1)
        valores_key = {_key(k): v for k, v in valores.items()}

        wb = openpyxl.load_workbook(plantilla, data_only=True)
        ws = wb.active
        filas_base = []
        for row in ws.iter_rows(min_row=4, max_col=5):
            med_c, prog_c, sol_c, sbod_c, sreal_c = row
            if med_c.value is None:
                continue
            k = _key(med_c.value)
            filas_base.append({
                'Medicamento': med_c.value,
                'Cantidad Programada': prog_c.value,
                'Cantidad Solicitada': sol_c.value,
                'Stock Sistema': sbod_c.value,
                'Stock Real': valores_key.get(k, sreal_c.value),
            })

    filas = []
    n_diff = 0
    n_pedir = 0
    n_falta = 0
    total_pedir = 0
    for f in filas_base:
        sreal = f['Stock Real']
        sbod = f['Stock Sistema']
        diff = None
        if sreal is not None and sbod is not None:
            diff = sreal - sbod
            if diff:
                n_diff += 1
        pedir = _cantidad_a_pedir(f['Cantidad Programada'], f['Cantidad Solicitada'], sreal)
        if sreal is None:
            n_falta += 1
        elif pedir:
            n_pedir += 1
            total_pedir += pedir
        filas.append({**f, 'Diferencia': diff, 'Cantidad a Pedir': pedir})

    os.makedirs(OUT_DIR, exist_ok=True)
    sal = os.path.join(OUT_DIR, f'Resumen_Conteo_AA_{hoy.strftime("%Y%m%d_%H%M")}.xlsx')
    sub = f'Basado en {os.path.basename(plantilla)}  ·  {n_falta} sin contar aún'
    _escribir(sal, filas, 'RESUMEN CONTEO vs PROGRAMACIÓN — Bodega AA', sub, hoy, resumen=True)

    print(f'Planilla base       : {os.path.basename(plantilla)}')
    print(f'Medicamentos        : {len(filas)}  ({n_falta} sin "Stock Real" todavía)')
    print(f'Con diferencia stock: {n_diff}  (Stock Real ≠ Stock Sistema)')
    print(f'A pedir             : {n_pedir} medicamento(s), {total_pedir} unidades en total')
    print(f'\nExcel: {os.path.basename(sal)}  (carpeta {os.path.basename(OUT_DIR)}\\)')
    if n_pedir:
        print('\nTop qué pedir (mayor cantidad primero):')
        top = sorted([f for f in filas if f['Cantidad a Pedir']],
                     key=lambda f: f['Cantidad a Pedir'], reverse=True)[:15]
        for f in top:
            print(f'  - {f["Medicamento"]}: {f["Cantidad a Pedir"]} ud')
    return sal


# ─────────────── escritura de Excel (instrumento y resumen comparten look) ──

def _escribir(sal, filas, titulo, subtitulo, hoy, resumen):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Resumen' if resumen else 'Conteo'

    hdrs = list(HDRS)
    if resumen:
        hdrs = hdrs + [('Diferencia (Real − Sistema)', 20), ('Cantidad a Pedir', 16)]
    ncols = len(hdrs)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws.cell(1, 1, f'{titulo}  ·  {hoy.strftime("%d/%m/%Y")}')
    ws.cell(1, 1).font = Font(bold=True, size=12, color='065F46', name='Arial')
    ws.row_dimensions[1].height = 22

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    ws.cell(2, 1, subtitulo)
    ws.cell(2, 1).font = Font(italic=True, size=9, color='555555', name='Arial')
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[2].height = 28

    for j, (label, w) in enumerate(hdrs, 1):
        c = ws.cell(3, j, label)
        c.fill = HFILL; c.font = HFONT; c.border = BRD
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[3].height = 26

    for i, f in enumerate(filas, 4):
        vals = [f['Medicamento'], f['Cantidad Programada'], f['Cantidad Solicitada'],
                f['Stock Sistema'], f['Stock Real']]
        if resumen:
            vals += [f.get('Diferencia'), f.get('Cantidad a Pedir')]
        for j, v in enumerate(vals, 1):
            c = ws.cell(i, j, v)
            c.border = BRD
            c.font = Font(name='Arial', size=10)
            if j >= 2:
                c.alignment = Alignment(horizontal='center')
        if resumen:
            if f['Stock Real'] is None:
                for j in range(1, ncols + 1):
                    ws.cell(i, j).fill = FALTA_FILL
            else:
                if f.get('Diferencia'):
                    dc = ws.cell(i, len(HDRS) + 1)
                    dc.fill = DIFF_FILL; dc.font = DIFF_FONT
                if f.get('Cantidad a Pedir'):
                    pc = ws.cell(i, len(HDRS) + 2)
                    pc.fill = PEDIR_FILL; pc.font = PEDIR_FONT

    last = 3 + len(filas)
    ws.freeze_panes = 'A4'
    if filas:
        ws.auto_filter.ref = f'A3:{get_column_letter(ncols)}{last}'
        ws.print_area = f'A1:{get_column_letter(ncols)}{last}'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = 'landscape'
    wb.save(sal)


# ─────────────── main ────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(
        description='Instrumento de conteo + diferencia + pedido, sin depender del Consolidado_AA_MAESTRO.')
    ap.add_argument('--programacion', default=None, metavar='XLSX',
                     help='Reporte SSASUR de Programación (Consumo por centro de costo → FARMACIA)')
    ap.add_argument('--stock', default=None, metavar='XLSX',
                     help='Reporte de stock crudo de SSASUR (reporte_de_stock_*.xlsx)')
    ap.add_argument('--bodega', default=BODEGA_DEFAULT,
                     help=f'Nombre de la bodega a filtrar en el reporte de stock (default: {BODEGA_DEFAULT})')
    ap.add_argument('--aplicar-conteo', default=None, metavar='XLSX_O_JSON',
                     help='Aplica el conteo: pásale el Instrumento_Conteo_AA_*.xlsx ya lleno, '
                          'o un JSON medicamento→cantidad')
    ap.add_argument('--plantilla', default=None, metavar='XLSX',
                     help='Fuerza el Instrumento_Conteo_AA_*.xlsx a usar con --aplicar-conteo JSON '
                          '(si no, usa el más reciente)')
    args = ap.parse_args()

    if args.aplicar_conteo:
        aplicar_conteo(args.aplicar_conteo, args.plantilla)
    else:
        if not args.programacion or not args.stock:
            print('[ERROR] Falta --programacion y/o --stock. Usa --help para ver el uso.')
            sys.exit(1)
        generar(args.programacion, args.stock, args.bodega)


if __name__ == '__main__':
    main()
