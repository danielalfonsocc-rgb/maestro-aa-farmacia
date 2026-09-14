# -*- coding: utf-8 -*-
"""
programacion_aa.py — Planilla de Programación del ciclo Bodega AA
===========================================================================
Al reiniciarse el ciclo de pedidos Bodega AA → Bodega Fármacos (mismo ciclo
de 2 semanas / 10 días hábiles que usa pedido_fusion.py), genera una planilla
con:

    Medicamento (orden alfabético) | Cantidad Programada | Cantidad Solicitada
    | Stock Bodega AA | Stock Real (en blanco, a contar) | Consumo Promedio
    Mensual | Sugerencia

  - Cantidad Programada / Cantidad Solicitada: salen del reporte mensual de
    SSASUR "Consumos por centro de costo" (Reportes → Reporte de consumo por
    centro de costo → Centro de Costo = FARMACIA → Generar XLS). Lo baja
    AUTO_SSASUR.py todos los días (PASO 4b) y aquí se detecta el más reciente.
    (Script eliminado 04-09-2026 y restaurado 14-09-2026: la planilla se sigue
    usando para el conteo físico de cada reinicio de ciclo.)
  - Stock Bodega AA: Consolidado_AA_MAESTRO → Pedido_Repos_Bodega → Stock_Bod_Actual.
  - Stock Real: se llena a mano tras el conteo físico y luego se aplica con
    --aplicar-conteo (ver más abajo).
  - Consumo Promedio Mensual: CMP_Mensual_22d, ya calculado/afinado en el
    Consolidado — es el requerimiento real que sustenta la Sugerencia.
  - Sugerencia: compara el Consumo Promedio Mensual contra la Cantidad
    Programada del reporte, e incluye la cantidad sugerida de programación
    (= el propio Consumo Promedio Mensual, redondeado):
      · "Subir programación a <N> ud"  si el real supera lo programado (+15%)
      · "Bajar programación a <N> ud"  si el real es menor a lo programado (-15%)
      · en blanco                       si están dentro del rango
      · "Incorporar a programación: <N> ud" si el medicamento NO aparece en
        el reporte de programación y lleva ≥3 meses consecutivos con demanda
        real (CMP_Mensual_22d > 0) sin estar programado — la racha se
        guarda en _historial_programacion.json.

Sin llamadas a IA — solo pandas + openpyxl.

Uso:
    py programacion_aa.py                        # solo genera si hoy es inicio de ciclo
    py programacion_aa.py --forzar                # genera igual, fuera del inicio de ciclo
    py programacion_aa.py --reporte ruta.xlsx      # fuerza el reporte SSASUR a usar
    py programacion_aa.py --auto                  # lo usa AUTO_SSASUR (PASO 5g): una vez
                                                  # por ciclo, con recuperación hasta 7 días

Segundo paso (tras contar físicamente y escanear la planilla impresa):
    py programacion_aa.py --aplicar-conteo conteo.json
        conteo.json = {"MEDICAMENTO TAL COMO SALE EN LA PLANILLA": 123, ...}
    Genera Resumen_Programacion_AA_<fecha>.xlsx en la carpeta Programacion_AA,
    con Diferencia (Stock Bodega AA − Stock Real) y las mismas Sugerencias.
    (Queda solo local: las carpetas "Programacion AA" de Drive/Escritorio se
    eliminaron el 04-09-2026 y no se restauraron.)
"""
import os
import re
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
from pedido_fusion import _feriados, _habil, CICLO_INICIO  # noqa: E402

setup_stdout()

OUT_DIR = os.path.join(WORK_DIR, 'Programacion_AA')
HIST_JSON = os.path.join(WORK_DIR, '_historial_programacion.json')
TOL_PCT = 0.15   # ±15% de tolerancia entre requerimiento real y lo programado
RACHA_MIN = 3    # meses consecutivos de demanda sin programación → incorporar

MESES_ES = {
    'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6,
    'JULIO': 7, 'AGOSTO': 8, 'SEPTIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11,
    'DICIEMBRE': 12,
}

THIN  = Side(style='thin', color='DCDCDC')
BRD   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HFILL = PatternFill('solid', fgColor='D1FAF5')
HFONT = Font(bold=True, color='065F46', name='Arial', size=10)

SUG_COL = {
    'Subir programación'          : ('FFE0B2', 'B45309'),
    'Bajar programación'          : ('E1F5FE', '01579B'),
    'Incorporar a programación'   : ('F4B3B3', '7F1D1D'),
    ''                             : ('F9FAFB', '374151'),
}


def _color_sugerencia(texto):
    """La Sugerencia incluye la cantidad sugerida (ej. 'Subir programación a 1200 ud'),
    así que el color se resuelve por el prefijo, no por igualdad exacta."""
    texto = texto or ''
    for prefijo, colores in SUG_COL.items():
        if prefijo and texto.startswith(prefijo):
            return colores
    return SUG_COL['']


def _pfill(hx):
    h = str(hx).lstrip('#')
    return PatternFill('solid', fgColor=('FF' + h) if len(h) == 6 else h)


def _key(nombre):
    n = norm_erp(nombre)
    return HOMOLOGACION.get(n, n)


# ─────────────── helpers de archivos ────────────────────────────────────────

def _mas_reciente(patron, extra_dirs=()):
    dirs = [WORK_DIR, *extra_dirs]
    cand = []
    for d in dirs:
        cand += [f for f in glob.glob(os.path.join(d, patron))
                 if not os.path.basename(f).startswith('~$')]
    return max(cand, key=os.path.getmtime) if cand else None


def _downloads_dir():
    perfil = os.environ.get('USERPROFILE', os.path.expanduser('~'))
    d = os.path.join(perfil, 'Downloads')
    return d if os.path.isdir(d) else WORK_DIR


def _universo_bodega():
    mae = _mas_reciente('Consolidado_AA_MAESTRO*.xlsx')
    if not mae:
        print('[ERROR] No se encontró Consolidado_AA_MAESTRO*.xlsx — corre maestro_aa.py primero.')
        sys.exit(1)
    df = pd.read_excel(mae, sheet_name='Pedido_Repos_Bodega', engine='openpyxl')
    df = df[['Medicamento', 'Stock_Bod_Actual', 'CMP_Mensual_22d']].copy()
    df['Medicamento'] = df['Medicamento'].astype(str).str.strip()
    df = df[df['Medicamento'] != ''].drop_duplicates(subset=['Medicamento'])
    return df.sort_values('Medicamento').reset_index(drop=True), mae


def _leer_reporte_ssasur(ruta):
    """Lee el reporte 'Consumos por centro de costo' de SSASUR (Generar XLS).
    Fila 0 = título, fila 1 = metadata con el mes/año, fila 2 = encabezados."""
    meta = pd.read_excel(ruta, header=None, nrows=2, engine='openpyxl')
    texto_meta = str(meta.iloc[1, 0]) if meta.shape[0] > 1 else ''
    m = re.search(r'mes de (\w+) de (\d{4})', texto_meta, re.IGNORECASE)
    periodo = None
    if m:
        mes_num = MESES_ES.get(m.group(1).strip().upper())
        if mes_num:
            periodo = f'{m.group(2)}-{mes_num:02d}'

    df = pd.read_excel(ruta, header=2, engine='openpyxl')
    df = df.rename(columns=lambda c: str(c).strip())
    if 'Centro Costo' in df.columns:
        df = df[df['Centro Costo'].astype(str).str.strip().str.upper() == 'FARMACIA']
    df['_key'] = df['Producto'].astype(str).map(_key)
    prog = dict(zip(df['_key'], pd.to_numeric(df['Total de Productos Programados'], errors='coerce')))
    sol  = dict(zip(df['_key'], pd.to_numeric(df['Productos Solicitado'], errors='coerce')))
    return prog, sol, periodo, texto_meta.strip()


def _cargar_historial():
    if os.path.isfile(HIST_JSON):
        try:
            with open(HIST_JSON, encoding='utf-8') as fh:
                return json.load(fh)
        except (OSError, json.JSONDecodeError):
            pass
    return {}


def _guardar_historial(hist):
    try:
        with open(HIST_JSON, 'w', encoding='utf-8') as fh:
            json.dump(hist, fh, ensure_ascii=False, indent=2)
    except OSError as e:
        print(f'  [aviso] no se pudo guardar {os.path.basename(HIST_JSON)}: {e}')


def _sugerencia(key, requerimiento_real, programado, periodo, hist):
    """Devuelve el texto de Sugerencia, con la cantidad sugerida de programación
    (= el Consumo Promedio Mensual ya calculado/afinado, redondeado a unidades)."""
    cant = round(requerimiento_real) if requerimiento_real else 0
    if programado is None or (isinstance(programado, float) and pd.isna(programado)):
        # No está en el reporte de programación: rastrea la racha de demanda.
        entry = hist.get(key, {'racha': 0, 'ultimo_mes': None})
        if requerimiento_real and requerimiento_real > 0:
            if periodo and entry.get('ultimo_mes') != periodo:
                entry['racha'] = entry.get('racha', 0) + 1
                entry['ultimo_mes'] = periodo
            hist[key] = entry
            if entry['racha'] >= RACHA_MIN:
                return f'Incorporar a programación: {cant} ud'
            return ''
        else:
            hist.pop(key, None)
            return ''
    else:
        # Está programado: se resetea cualquier racha acumulada de antes.
        hist.pop(key, None)
        if requerimiento_real is None or pd.isna(requerimiento_real):
            return ''
        if programado <= 0:
            return f'Subir programación a {cant} ud' if requerimiento_real > 0 else ''
        ratio = requerimiento_real / programado
        if ratio > 1 + TOL_PCT:
            return f'Subir programación a {cant} ud'
        if ratio < 1 - TOL_PCT:
            return f'Bajar programación a {cant} ud'
        return ''


# ─────────────── modo 1: generar planilla del ciclo ────────────────────────

# Anchos pensados para imprimir en CARTA VERTICAL (ajuste a 1 página de ancho):
# Medicamento y Sugerencia van con ajuste de texto en vez de columnas anchas.
HDRS = [
    ('Medicamento',                40),
    ('Cantidad Programada',        11),
    ('Cantidad Solicitada',        11),
    ('Stock Bodega AA',            10),
    ('Stock Real',                 12),
    ('Consumo Promedio Mensual',   11),
    ('Sugerencia',                 22),
]


def _es_inicio_ciclo(hoy, fer):
    """True si hoy es el primer día hábil del ciclo Bod_Farmacos.

    El ciclo tiene límite de CALENDARIO (cada 14 días desde CICLO_INICIO, ver
    pedido_fusion._dias_ciclo). Antes esto exigía _dias_ciclo(hoy) == 10, y un
    ciclo con feriado adentro nunca llega a 10: el 14-09-2026 (18-09 feriado)
    quedaba con 9 y la planilla no se reconocía como inicio de ciclo. Ahora se
    busca el primer día hábil desde el lunes de inicio de la ventana, así que
    también cubre un lunes feriado (el inicio pasa al martes).
    """
    if hoy < CICLO_INICIO or not _habil(hoy, fer):
        return False
    inicio, _ = _ventana_ciclo(hoy)
    return hoy == _primer_habil(inicio, fer)


def _ventana_ciclo(hoy):
    """(lunes de inicio, domingo de cierre) de la ventana de 14 días de `hoy`."""
    inicio = CICLO_INICIO + dt.timedelta(days=14 * ((hoy - CICLO_INICIO).days // 14))
    return inicio, inicio + dt.timedelta(days=13)


def _primer_habil(desde, fer):
    d = desde
    while not _habil(d, fer):
        d += dt.timedelta(days=1)
    return d


# ─────────────── modo --auto (lo llama AUTO_SSASUR.py en cada corrida) ─────────

CATCHUP_DIAS = 7   # días corridos desde el lunes de inicio en que aún se recupera la planilla


def _planilla_del_ciclo(inicio, fin):
    """Programacion_AA_<AAAAMMDD>.xlsx ya generada con fecha dentro de la ventana
    del ciclo (ignora .bak y Resumen_*), o None."""
    for f in glob.glob(os.path.join(OUT_DIR, 'Programacion_AA_*.xlsx')):
        m = re.fullmatch(r'Programacion_AA_(\d{8})\.xlsx', os.path.basename(f))
        if m:
            fecha = dt.datetime.strptime(m.group(1), '%Y%m%d').date()
            if inicio <= fecha <= fin:
                return f
    return None


def generar_auto():
    """Genera la planilla UNA vez por ciclo, sin intervención.

    Normalmente el primer día hábil del ciclo. Si ese día la corrida falló o no
    corrió (PC apagado, sesión SSASUR caída, etc.), se recupera en la primera
    corrida siguiente dentro de CATCHUP_DIAS — mismo patrón de catch-up que
    Centinela. "Ya generada" = existe una Programacion_AA_<fecha>.xlsx con fecha
    dentro de la ventana del ciclo, así que tampoco duplica una hecha a mano.
    Solo usa un reporte de Programación bajado HOY: con uno viejo la Cantidad
    Solicitada (o el mes entero) estaría desactualizada, y es mejor reintentar
    en la próxima corrida.
    """
    hoy = dt.date.today()
    fer = _feriados()
    if hoy < CICLO_INICIO or not _habil(hoy, fer):
        print('Hoy no es día hábil — no corresponde hoja de inventario.')
        return
    inicio, fin = _ventana_ciclo(hoy)
    primer = _primer_habil(inicio, fer)
    ya = _planilla_del_ciclo(inicio, fin)
    if ya:
        print(f'Ciclo del {inicio.strftime("%d/%m")}: ya existe {os.path.basename(ya)} — omitido.')
        return
    if hoy > inicio + dt.timedelta(days=CATCHUP_DIAS - 1):
        print(f'Ciclo del {inicio.strftime("%d/%m")} sin planilla, pero ya pasó la ventana de '
              f'recuperación ({CATCHUP_DIAS} días) — se genera en el próximo ciclo '
              f'(o a mano con --forzar).')
        return
    reporte = _mas_reciente('cantidad_de_productos_consumidos_en_centro_de_costo_farmacia*.xlsx')
    if not reporte or dt.date.fromtimestamp(os.path.getmtime(reporte)) != hoy:
        print('[aviso] No hay reporte de Programación AA descargado HOY (¿falló el PASO 4b o se '
              'usó --no-programacion?) — no genero la hoja; se reintenta en la próxima corrida.')
        return
    nota = None if hoy == primer else (
        f'generada en recuperación: el ciclo empezó el {primer.strftime("%d/%m/%Y")}')
    generar(reporte, forzar=True, nota=nota)


def generar(ruta_reporte=None, forzar=False, nota=None):
    hoy = dt.date.today()
    fer = _feriados()
    inicio_ciclo = _es_inicio_ciclo(hoy, fer)
    if not forzar and not inicio_ciclo:
        print(f'Hoy ({hoy.strftime("%d/%m/%Y")}) no es el inicio del ciclo Bodega AA → Bodega Fármacos.')
        print('No se genera planilla. Usa --forzar para generarla igual.')
        return
    if nota is None and forzar and not inicio_ciclo:
        nota = 'generada fuera del inicio de ciclo (--forzar)'

    universo, mae = _universo_bodega()
    ruta_reporte = ruta_reporte or _mas_reciente(
        'cantidad_de_productos_consumidos_en_centro_de_costo_farmacia*.xlsx',
        extra_dirs=[_downloads_dir()])
    if not ruta_reporte:
        print('[ERROR] No se encontró el reporte de SSASUR '
              '"cantidad_de_productos_consumidos_en_centro_de_costo_farmacia*.xlsx" '
              'en la carpeta del proyecto ni en Descargas.')
        print('  Descárgalo desde SSASUR → Reportes → Consumo por centro de costo → '
              'Centro de Costo = FARMACIA → Generar XLS.')
        sys.exit(1)

    prog, sol, periodo, meta_txt = _leer_reporte_ssasur(ruta_reporte)
    hist = _cargar_historial()

    filas = []
    n_sin_reporte = 0
    for r in universo.itertuples(index=False):
        key = _key(r.Medicamento)
        programado = prog.get(key)
        solicitado = sol.get(key)
        req_real = r.CMP_Mensual_22d if not pd.isna(r.CMP_Mensual_22d) else 0
        if key not in prog:
            n_sin_reporte += 1
        sugerencia = _sugerencia(key, req_real, programado, periodo, hist)
        filas.append({
            'Medicamento': r.Medicamento,
            'Cantidad Programada': None if programado is None or pd.isna(programado) else int(programado),
            'Cantidad Solicitada': None if solicitado is None or pd.isna(solicitado) else int(solicitado),
            'Stock Bodega AA': int(r.Stock_Bod_Actual) if not pd.isna(r.Stock_Bod_Actual) else 0,
            'Stock Real': None,
            'Consumo Promedio Mensual': round(req_real),
            'Sugerencia': sugerencia,
        })

    _guardar_historial(hist)
    os.makedirs(OUT_DIR, exist_ok=True)
    sal = os.path.join(OUT_DIR, f'Programacion_AA_{hoy.strftime("%Y%m%d")}.xlsx')
    _escribir_planilla(sal, filas, mae, ruta_reporte, meta_txt, hoy, nota)
    sal_pdf = os.path.splitext(sal)[0] + '.pdf'
    sub_pdf = (f'Reporte SSASUR: {meta_txt or os.path.basename(ruta_reporte)}  ·  '
               f'Stock Bodega AA según {os.path.basename(mae)}  ·  '
               f'Sugerencia con tolerancia ±{int(TOL_PCT*100)}%'
               + (f'  ·  {nota}' if nota else ''))
    try:
        _escribir_pdf(sal_pdf, filas, hoy, sub_pdf)
    except PermissionError:
        print(f'  [aviso] {os.path.basename(sal_pdf)} está abierto — no se pudo regenerar el PDF.')

    n_subir = sum(1 for f in filas if f['Sugerencia'].startswith('Subir programación'))
    n_bajar = sum(1 for f in filas if f['Sugerencia'].startswith('Bajar programación'))
    n_incorp = sum(1 for f in filas if f['Sugerencia'].startswith('Incorporar a programación'))
    print(f'HOY = {hoy}')
    print(f'Sistema  : {os.path.basename(mae)}')
    print(f'Reporte  : {os.path.basename(ruta_reporte)}  ({meta_txt or "sin metadata de mes"})')
    print(f'{len(filas)} medicamentos | {n_sin_reporte} sin programación en el reporte')
    print(f'Sugerencias: Subir {n_subir} | Bajar {n_bajar} | Incorporar a programación {n_incorp}')
    print(f'\nExcel: {sal}')
    print(f'PDF (carta vertical, para imprimir): {sal_pdf}')
    print('\nImprime esta planilla, cuenta físicamente Bodega AA y llena "Stock Real" a mano.')
    print('Cuando esté escaneada, avisa para transcribirla y correr --aplicar-conteo.')


def _escribir_planilla(sal, filas, mae, ruta_reporte, meta_txt, hoy, nota=None,
                         es_resumen=False, n_diferencias=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Resumen' if es_resumen else 'Programacion'

    hdrs = list(HDRS)
    if es_resumen:
        hdrs = hdrs + [('Diferencia (Real - Bod.AA)', 20)]
    ncols = len(hdrs)

    titulo = ('RESUMEN CONTEO vs PROGRAMACIÓN — Bodega AA' if es_resumen
              else 'PLANILLA DE PROGRAMACIÓN — Bodega AA')
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws.cell(1, 1, f'{titulo}  ·  {hoy.strftime("%d/%m/%Y")}')
    ws.cell(1, 1).font = Font(bold=True, size=12, color='065F46', name='Arial')
    ws.row_dimensions[1].height = 22

    sub = (f'Sistema: {os.path.basename(mae)}  ·  Reporte SSASUR: {os.path.basename(ruta_reporte)} '
           f'({meta_txt or "sin metadata de mes"})  ·  Tolerancia sugerencia: ±{int(TOL_PCT*100)}%  ·  '
           f'"Incorporar a programación" = ≥{RACHA_MIN} meses seguidos con demanda sin estar programado')
    if nota and not es_resumen:
        sub += f'  ·  {nota}'
    if es_resumen and n_diferencias is not None:
        sub += f'  ·  {n_diferencias} medicamento(s) con diferencia entre Stock Bodega AA y Stock Real'
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    ws.cell(2, 1, sub)
    ws.cell(2, 1).font = Font(italic=True, size=9, color='555555', name='Arial')
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[2].height = 30

    for j, (label, w) in enumerate(hdrs, 1):
        c = ws.cell(3, j, label)
        c.fill = HFILL; c.font = HFONT; c.border = BRD
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[3].height = 26

    for i, f in enumerate(filas, 4):
        vals = [f['Medicamento'], f['Cantidad Programada'], f['Cantidad Solicitada'],
                f['Stock Bodega AA'], f['Stock Real'], f['Consumo Promedio Mensual'],
                f['Sugerencia']]
        if es_resumen:
            vals.append(f.get('Diferencia'))
        bg, fg = _color_sugerencia(f['Sugerencia'])
        for j, v in enumerate(vals, 1):
            c = ws.cell(i, j, v)
            c.border = BRD
            if j >= 2:
                c.fill = _pfill(bg)
                c.font = Font(name='Arial', size=10, color=fg)
                c.alignment = Alignment(horizontal='center', vertical='center',
                                        wrap_text=(j == 7))
            else:
                c.font = Font(name='Arial', size=10)
                c.alignment = Alignment(vertical='center', wrap_text=True)
        if es_resumen and f.get('Diferencia') not in (None, 0):
            dc = ws.cell(i, len(hdrs))
            dc.font = Font(name='Arial', size=10, color='7F1D1D', bold=True)
            dc.fill = _pfill('F4B3B3')

    last = 3 + len(filas)
    ws.freeze_panes = 'A4'
    if filas:
        ws.auto_filter.ref = f'A3:{get_column_letter(ncols)}{last}'
        ws.print_area = f'A1:{get_column_letter(ncols)}{last}'
    # Impresión: CARTA VERTICAL, 1 página de ancho, encabezado repetido en cada hoja.
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.print_title_rows = '3:3'
    ws.print_options.horizontalCentered = True
    ws.oddFooter.center.text = 'Página &P de &N'
    wb.save(sal)


# ─────────────── PDF para imprimir (carta vertical) ─────────────────────────

def _fuente_pdf():
    """Arial del sistema (tildes, ñ, ±) si está disponible; si no, Helvetica."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    fonts = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts')
    try:
        pdfmetrics.registerFont(TTFont('ArialAA', os.path.join(fonts, 'arial.ttf')))
        pdfmetrics.registerFont(TTFont('ArialAA-Bold', os.path.join(fonts, 'arialbd.ttf')))
        return 'ArialAA', 'ArialAA-Bold'
    except Exception:
        return 'Helvetica', 'Helvetica-Bold'


def _escribir_pdf(sal_pdf, filas, hoy, subtitulo):
    """Misma planilla que el Excel, lista para imprimir en CARTA VERTICAL:
    encabezado repetido en cada página, columna "Stock Real" en blanco con
    espacio para escribir a mano, y colores tenues por sugerencia (paleta de
    impresión económica)."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from xml.sax.saxutils import escape

    fuente, fuente_b = _fuente_pdf()
    margen = 10 * mm
    titulo = f'HOJA DE INVENTARIO — Bodega AA  ·  Ciclo {hoy.strftime("%d/%m/%Y")}'

    st_cel = ParagraphStyle('cel', fontName=fuente, fontSize=7.5, leading=9)
    st_sug = ParagraphStyle('sug', fontName=fuente, fontSize=6.5, leading=8, alignment=1)
    st_hdr = ParagraphStyle('hdr', fontName=fuente_b, fontSize=7, leading=8.5, alignment=1,
                            textColor=colors.HexColor('#065F46'))
    st_tit = ParagraphStyle('tit', fontName=fuente_b, fontSize=12, leading=15,
                            textColor=colors.HexColor('#065F46'))
    st_sub = ParagraphStyle('sub', fontName=fuente, fontSize=7, leading=9,
                            textColor=colors.HexColor('#555555'))

    encabezados = ['Medicamento', 'Cant. Programada', 'Cant. Solicitada', 'Stock Bodega AA',
                   'Stock Real (conteo)', 'Consumo Prom. Mensual', 'Sugerencia']
    datos = [[Paragraph(h, st_hdr) for h in encabezados]]
    estilos = []
    for i, f in enumerate(filas, 1):
        num = lambda v: '' if v is None else f'{v:,}'.replace(',', '.')
        datos.append([
            Paragraph(escape(str(f['Medicamento'])), st_cel),
            num(f['Cantidad Programada']), num(f['Cantidad Solicitada']),
            num(f['Stock Bodega AA']), '', num(f['Consumo Promedio Mensual']),
            Paragraph(escape(f['Sugerencia'] or ''), st_sug),
        ])
        bg, fg = _color_sugerencia(f['Sugerencia'])
        if f['Sugerencia']:
            estilos.append(('BACKGROUND', (6, i), (6, i), colors.HexColor('#' + bg)))
        elif i % 2 == 0:
            estilos.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor('#F7F7F7')))

    ancho = letter[0] - 2 * margen
    cols = [0.33, 0.10, 0.10, 0.085, 0.12, 0.085, 0.18]   # "Programada" no cabe en menos
    tabla = Table(datos, colWidths=[ancho * c for c in cols], repeatRows=1, rowHeights=None)
    tabla.setStyle(TableStyle([
        ('FONT', (0, 1), (-1, -1), fuente, 8),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D1FAF5')),
        ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#BDBDBD')),
        ('BOX', (4, 0), (4, -1), 0.9, colors.HexColor('#424242')),   # columna a llenar a mano
        ('BACKGROUND', (4, 1), (4, -1), colors.white),
        ('ALIGN', (1, 1), (5, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 1), (-1, -1), 4), ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
        *estilos,
    ]))

    def _pie(canvas, doc):
        canvas.saveState()
        canvas.setFont(fuente, 7)
        canvas.setFillColor(colors.HexColor('#666666'))
        canvas.drawString(margen, 6 * mm, titulo)
        canvas.drawRightString(letter[0] - margen, 6 * mm, f'Página {doc.page}')
        canvas.restoreState()

    doc = SimpleDocTemplate(sal_pdf, pagesize=letter, leftMargin=margen, rightMargin=margen,
                            topMargin=margen, bottomMargin=12 * mm,
                            title=titulo, author='Farmacia AT Abierta — Hospital de Pitrufquén')
    doc.build([Paragraph(escape(titulo), st_tit), Paragraph(escape(subtitulo), st_sub),
               Spacer(1, 3 * mm), tabla], onFirstPage=_pie, onLaterPages=_pie)


BODEGA_FISICA_AA = 'BODEGA AT ABIERTA'  # debe calzar con BODEGAS_AA_BODEGA en maestro_aa.py


def _stock_bodega_desde_reporte(ruta_reporte):
    """Lee un reporte_de_stock_*.xlsx crudo de SSASUR (no el Consolidado) y
    devuelve {key: cantidad} para BODEGA_FISICA_AA, sumando por lote."""
    df = pd.read_excel(ruta_reporte, header=2, engine='openpyxl')
    df.columns = [str(c).strip() for c in df.columns]
    descr_col = next(c for c in df.columns if 'Descrip' in c)
    bod_col = next(c for c in df.columns if 'Bodega' in c)
    df[bod_col] = df[bod_col].astype(str).str.strip().str.upper()
    sub = df[df[bod_col] == BODEGA_FISICA_AA].copy()
    sub['Cantidad'] = pd.to_numeric(sub['Cantidad'], errors='coerce').fillna(0)
    sub['_key'] = sub[descr_col].astype(str).map(_key)
    return sub.groupby('_key')['Cantidad'].sum().to_dict()


def _mejor_stock_reporte(valores_key):
    """AUTO_SSASUR puede correr varias veces entre que se imprime la planilla
    y se termina el conteo físico, y el stock del sistema se mueve entre una
    corrida y otra — "la más reciente" no siempre es la que mejor refleja lo
    que había en la bodega al momento de contar. Compara TODOS los
    reporte_de_stock_*.xlsx descargados contra el conteo físico y usa el que
    más coincidencias tiene (empate → el más reciente por fecha).
    Devuelve (stock_key, archivo_elegido, ranking) o ({}, None, []) si no hay
    ningún reporte_de_stock_*.xlsx disponible."""
    candidatos = [f for f in glob.glob(os.path.join(WORK_DIR, 'reporte_de_stock_*.xlsx'))
                  if not os.path.basename(f).startswith('~$')]
    if not candidatos:
        return {}, None, []

    ranking = []
    for f in candidatos:
        try:
            stock = _stock_bodega_desde_reporte(f)
        except (OSError, ValueError, KeyError, StopIteration) as e:
            print(f'  [aviso] no se pudo leer {os.path.basename(f)}: {e}')
            continue
        coincide = sum(1 for k, real in valores_key.items()
                        if k in stock and int(stock[k]) == real)
        total = sum(1 for k in valores_key if k in stock)
        ranking.append((f, stock, coincide, total, os.path.getmtime(f)))

    if not ranking:
        return {}, None, []

    ranking.sort(key=lambda r: (r[2], r[4]), reverse=True)
    mejor_f, mejor_stock, mejor_c, mejor_t, _ = ranking[0]
    resumen = [(os.path.basename(f), c, t) for f, _, c, t, _ in ranking]
    return {k: int(v) for k, v in mejor_stock.items()}, mejor_f, resumen


# ─────────────── modo 2: aplicar conteo (tras escanear) ────────────────────

def aplicar_conteo(ruta_json, ruta_plantilla=None):
    # Si un nuevo ciclo generó una Programacion_AA_*.xlsx más reciente mientras
    # el conteo físico de la anterior seguía en curso, "la más reciente" ya NO
    # es la que se imprimió y contó — hay que poder fijarla a mano.
    plantilla = ruta_plantilla or _mas_reciente('Programacion_AA_*.xlsx', extra_dirs=[OUT_DIR])
    if not plantilla:
        print('[ERROR] No hay ninguna Programacion_AA_*.xlsx generada todavía. '
              'Corre primero: py programacion_aa.py --forzar')
        sys.exit(1)
    if not ruta_plantilla:
        otras = [f for f in glob.glob(os.path.join(WORK_DIR, 'Programacion_AA_*.xlsx'))
                 if os.path.abspath(f) != os.path.abspath(plantilla)]
        if otras:
            print(f'  [aviso] hay {len(otras)} otra(s) Programacion_AA_*.xlsx en la carpeta — '
                  f'se está usando la más reciente ({os.path.basename(plantilla)}). '
                  'Si el conteo físico corresponde a otra, usa --plantilla ruta.xlsx.')
    try:
        with open(ruta_json, encoding='utf-8') as fh:
            valores = json.load(fh)
    except (OSError, json.JSONDecodeError) as e:
        print(f'[ERROR] No se pudo leer {ruta_json}: {e}')
        sys.exit(1)

    valores_key = {_key(k): v for k, v in valores.items()}

    # Stock Bodega AA ACTUAL (no el congelado en la planilla impresa): el conteo
    # físico suele hacerse varios días después de imprimir la planilla, y el
    # stock del sistema se mueve en el intertanto. Comparar contra el número
    # ya viejo produce "diferencias" falsas que solo reflejan el paso del
    # tiempo, no un problema real de inventario.
    #
    # Además, entre que se imprime la planilla y se termina de contar puede
    # correr AUTO_SSASUR varias veces — se compara contra TODAS las corridas
    # de reporte_de_stock_*.xlsx disponibles y se usa la que más coincide con
    # el conteo físico, no automáticamente "la más reciente".
    universo, mae = _universo_bodega()
    stock_actual_key, archivo_stock, ranking = _mejor_stock_reporte(valores_key)
    if archivo_stock:
        mejor_c, mejor_t = next((c, t) for f, c, t in ranking if f == os.path.basename(archivo_stock))
        print(f'  Stock Bodega AA: {os.path.basename(archivo_stock)} '
              f'— la que más coincide con el conteo ({mejor_c}/{mejor_t})')
        for nombre, c, t in ranking:
            if nombre != os.path.basename(archivo_stock):
                print(f'    (descartada: {nombre} → {c}/{t})')
    else:
        stock_actual_key = {
            _key(r.Medicamento): (0 if pd.isna(r.Stock_Bod_Actual) else int(r.Stock_Bod_Actual))
            for r in universo.itertuples(index=False)
        }
        print(f'  [aviso] no se encontraron reporte_de_stock_*.xlsx sueltos; '
              f'usando Stock_Bod_Actual de {os.path.basename(mae)}')

    wb = openpyxl.load_workbook(plantilla)
    ws = wb.active
    hoy = dt.date.today()
    filas = []
    n_diff = 0
    # key → índice en filas: evita filas duplicadas cuando la planilla tiene un
    # medicamento con dos nombres (uno programado, uno no programado) que
    # HOMOLOGACION colapsa al mismo key. Gana el que tiene Programada≠None.
    key_idx: dict[str, int] = {}
    for row in ws.iter_rows(min_row=4, max_col=7):
        med_c, prog_c, sol_c, sbod_c, sreal_c, cpm_c, sug_c = row
        if med_c.value is None:
            continue
        key = _key(med_c.value)
        stock_real = valores_key.get(key)
        sbod = stock_actual_key.get(key, sbod_c.value or 0)
        diff = None if stock_real is None else stock_real - sbod
        fila = {
            'Medicamento': med_c.value,
            'Cantidad Programada': prog_c.value,
            'Cantidad Solicitada': sol_c.value,
            'Stock Bodega AA': sbod,
            'Stock Real': stock_real,
            'Consumo Promedio Mensual': cpm_c.value,
            'Sugerencia': sug_c.value or '',
            'Diferencia': diff,
        }
        if key in key_idx:
            prev = filas[key_idx[key]]
            # si la fila previa es huérfana (sin Programada) y esta tiene Programada, reemplazar
            if prev['Cantidad Programada'] is None and prog_c.value is not None:
                if prev.get('Diferencia') not in (None, 0):
                    n_diff -= 1
                filas[key_idx[key]] = fila
                if diff not in (None, 0):
                    n_diff += 1
            # si la fila actual es huérfana, descartarla silenciosamente
        else:
            key_idx[key] = len(filas)
            filas.append(fila)
            if diff not in (None, 0):
                n_diff += 1

    fuente_stock = archivo_stock or mae
    os.makedirs(OUT_DIR, exist_ok=True)
    sal = os.path.join(OUT_DIR, f'Resumen_Programacion_AA_{hoy.strftime("%Y%m%d_%H%M")}.xlsx')
    _escribir_planilla(sal, filas, fuente_stock, plantilla, '', hoy, None,
                        es_resumen=True, n_diferencias=n_diff)

    print(f'Planilla base : {os.path.basename(plantilla)}')
    print(f'Stock Bodega AA: refrescado desde {os.path.basename(fuente_stock)} (no el congelado de la planilla impresa)')
    print(f'Conteo aplicado: {len(valores_key)} medicamento(s) en {os.path.basename(ruta_json)}')
    print(f'Diferencias detectadas: {n_diff} de {len(filas)} medicamentos')
    print(f'\nExcel: {os.path.basename(sal)}  (carpeta {os.path.basename(OUT_DIR)}\\)')
    print('Sube esta carpeta a Drive y Escritorio con publicar_drive.py / publicar_escritorio.py.')


# ─────────────── main ────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(
        description='Planilla de Programación del ciclo Bodega AA. Sin IA.')
    ap.add_argument('--forzar', action='store_true',
                     help='Genera la planilla aunque hoy no sea el inicio del ciclo')
    ap.add_argument('--auto', action='store_true',
                     help='Modo AUTO_SSASUR: genera una vez por ciclo (con recuperación si '
                          'el primer día hábil no corrió), solo con reporte bajado hoy')
    ap.add_argument('--reporte', default=None,
                     help='Ruta a un reporte SSASUR específico (si no, usa el más reciente)')
    ap.add_argument('--aplicar-conteo', default=None, metavar='JSON',
                     help='Aplica el conteo físico (JSON medicamento→cantidad) y genera el Resumen final')
    ap.add_argument('--plantilla', default=None, metavar='XLSX',
                     help='Fuerza la Programacion_AA_*.xlsx a usar con --aplicar-conteo '
                          '(si no, usa la más reciente — cuidado si ya empezó un ciclo nuevo)')
    args = ap.parse_args()

    if args.aplicar_conteo:
        aplicar_conteo(args.aplicar_conteo, args.plantilla)
    elif args.auto:
        generar_auto()
    else:
        generar(args.reporte, args.forzar)


if __name__ == '__main__':
    main()
