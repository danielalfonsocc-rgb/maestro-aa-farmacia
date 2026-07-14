# -*- coding: utf-8 -*-
"""
pedido_fusion.py v2 — Planilla Simplificada de Pedidos AA
==========================================================
Genera  Pedido_Fusion_AA_<fecha>.xlsx  con 3 hojas:

  1 "Farm_Bod"       Farmacia AA → Bodega AA
      Factor_Empaque = Unidades_Caja de SGLI_Estres (ICP CENABAST)
      Pedir Hoy ajustado al día de la semana (no hay columna Req.5d)
  2 "Bod_Farmacos"   Bodega AA → Bodega Fármacos (ciclo 2 semanas)
  3 "Dialisis"       Diálisis mensual: Farmacia AA + Bodega AA
      Consumo = Consumo_5D_Solo_Dialisis / 5 × 30 (días naturales)
      Activa en S3 o con --forzar-dialisis

Sin llamadas a IA — solo pandas + openpyxl.

Uso:
    py pedido_fusion.py
    py pedido_fusion.py --forzar-dialisis
    py pedido_fusion.py --todos
"""
import os, math, datetime as dt, argparse, glob
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

WORK_DIR     = os.path.dirname(os.path.abspath(__file__))
FERIADOS_CSV = os.path.join(WORK_DIR, 'feriados_chile.csv')
BUFFER_SS    = 1    # días de safety stock (blindaje reapertura lunes)
EXTRA_CRIT   = 1    # días extra SS para criticidad ≤ 2
CICLO_INICIO = dt.date(2026, 7, 13)   # inicio ciclo Bod→BodFarm; repite cada 10d hábiles
# Recalibrado 2026-07-13: el ancla anterior (2026-06-29) cayó en feriado
# (San Pedro y San Pablo), lo que restó 1 día hábil al 1er ciclo y corrió
# el límite del 2° ciclo de lunes 13-jul a martes 14-jul. Confirmado con
# el usuario que el nuevo período de pedido Bod→BodFarm arranca esta semana.

# ─────────────── helpers ────────────────────────────────────────────────────

def _maestro():
    c = [f for f in glob.glob(os.path.join(WORK_DIR, 'Consolidado_AA_MAESTRO*.xlsx'))
         if not os.path.basename(f).startswith('~$')]
    return max(c, key=os.path.getmtime) if c else \
           os.path.join(WORK_DIR, 'Consolidado_AA_MAESTRO.xlsx')

def _feriados():
    f = {}
    try:
        with open(FERIADOS_CSV, encoding='utf-8') as fh:
            next(fh)
            for ln in fh:
                p = ln.rstrip().split(';')
                if p[0].strip():
                    f[dt.date.fromisoformat(p[0].strip())] = p[1].strip() if len(p) > 1 else ''
    except FileNotFoundError:
        pass
    return f

def _habil(d, fer):
    return d.weekday() < 5 and d not in fer

def _dias_ef(hoy, fer):
    """Días hábiles desde hoy (inclusive) hasta el viernes; 5 si es finde/feriado."""
    if not _habil(hoy, fer):
        return 5
    d, n = hoy, 0
    while d.weekday() < 5:
        if _habil(d, fer):
            n += 1
        d += dt.timedelta(days=1)
    return max(1, n)

def _dias_ciclo(hoy, fer):
    """Días hábiles restantes en el ciclo Bod→BodFarm (10d hábiles, inicio CICLO_INICIO, repite cada 2 semanas)."""
    if hoy < CICLO_INICIO:
        return 10
    d = CICLO_INICIO
    habiles = 0
    while d < hoy:
        if _habil(d, fer):
            habiles += 1
        d += dt.timedelta(days=1)
    pos = habiles % 10        # posición 0-9 dentro del ciclo
    return max(1, 10 - pos)  # días restantes en el ciclo actual

def _semana(d):
    return min((d.day - 1) // 7 + 1, 4)

def _nivel(c):
    s = str(c).strip().upper()
    if len(s) >= 2 and s[0] in '12345' and s[1] == '-':
        return int(s[0])
    if 'CRITICO' in s: return 1
    if 'URGENTE' in s or 'ALTO' in s: return 2
    if 'MODERADO' in s: return 3
    if 'BAJO' in s: return 4
    return 5

def _n(v, d=0.0):
    v = pd.to_numeric(v, errors='coerce')
    return d if pd.isna(v) else float(v)

def _ceil_fe(raw, fe):
    fe = max(1, int(fe))
    if raw <= 0:
        return 0
    return int(math.ceil(raw / fe)) * fe

# ─────────────── paleta de colores tenues ───────────────────────────────────

# (fondo muy claro, texto discreto) — sin colores saturados en urgentes
CPAL = {
    1: ('FFF0EF', 'B91C1C'),   # rojo pálido / vino
    2: ('FFF5EC', 'B45309'),   # durazno pálido / ocre
    3: ('FFFDE7', '856404'),   # crema / dorado
    4: ('F0FDF4', '166534'),   # verde muy claro / esmeralda
    5: ('F9FAFB', '374151'),   # gris casi blanco / gris oscuro
}
DIAL_BG = ('EFF6FF', '1E40AF')   # azul muy claro para diálisis
THIN    = Side(style='thin', color='DCDCDC')
BRD     = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HFILL   = PatternFill('solid', fgColor='D1FAF5')
HFONT   = Font(bold=True, color='065F46', name='Arial', size=10)

def _pfill(hx):
    h = str(hx).lstrip('#')
    if len(h) == 6:
        h = 'FF' + h   # openpyxl necesita ARGB de 8 dígitos; sin FF queda transparente
    return PatternFill('solid', fgColor=h)

def _totals(ws, data_start, data_end, ncols, sum_cols):
    """Fila de totales en negrita inmediatamente después de los datos."""
    tr = data_end + 1
    for j in range(1, ncols + 1):
        c = ws.cell(tr, j)
        c.fill = _pfill('D1FAF5'); c.border = BRD
        if j == 1:
            c.value = 'TOTAL'
            c.font = Font(bold=True, name='Arial', size=10, color='065F46')
        elif j in sum_cols:
            c.value = f'=SUM({get_column_letter(j)}{data_start}:{get_column_letter(j)}{data_end})'
            c.font = Font(bold=True, name='Arial', size=10, color='065F46')
            c.alignment = Alignment(horizontal='center')


def _hdr(ws, row, hdrs):
    for j, (label, w) in enumerate(hdrs, 1):
        c = ws.cell(row, j, label)
        c.fill = HFILL; c.font = HFONT; c.border = BRD
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[row].height = 26

def _titulo(ws, txt, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    t = ws.cell(1, 1, txt)
    t.font = Font(bold=True, size=12, color='065F46', name='Arial')
    ws.row_dimensions[1].height = 22

def _subtit(ws, txt, ncols, height=36):
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(2, 1, txt)
    c.font = Font(italic=True, size=9, color='555555', name='Arial')
    c.alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[2].height = height

FMT1D = '0.0'   # 1 decimal para coberturas y CDL

def _fila_crit(ws, row_i, vals, crit, cols_center, cols_fmt1d=None):
    nv = _nivel(crit)
    bg, fg = CPAL.get(nv, CPAL[5])
    for j, v in enumerate(vals, 1):
        c = ws.cell(row_i, j, v)
        c.border = BRD
        c.fill = _pfill(bg)
        c.font = Font(name='Arial', size=10, color=fg)
        if j in cols_center:
            c.alignment = Alignment(horizontal='center')
        if cols_fmt1d and j in cols_fmt1d:
            c.number_format = FMT1D

def _fila_dial(ws, row_i, vals, cols_center, cols_fmt1d=None):
    bg, fg = DIAL_BG
    for j, v in enumerate(vals, 1):
        c = ws.cell(row_i, j, v)
        c.border = BRD
        c.fill = _pfill(bg)
        c.font = Font(name='Arial', size=10, color=fg)
        if j in cols_center:
            c.alignment = Alignment(horizontal='center')
        if cols_fmt1d and j in cols_fmt1d:
            c.number_format = FMT1D

# ─────────────── carga de datos ─────────────────────────────────────────────

def _leer(maestro):
    def sh(name):
        try:
            return pd.read_excel(maestro, sheet_name=name, engine='openpyxl')
        except Exception:
            return pd.DataFrame()
    return {k: sh(v) for k, v in {
        'sgli' : 'SGLI_Estres',
        'farm' : 'Pedido_Farm_Bodega',
        'bod'  : 'Pedido_Repos_Bodega',
        'dfarm': 'Dialisis_Pedido_Farm',
        'dbod' : 'Dialisis_Pedido_Bod',
    }.items()}

# ─────────────── Hoja 1: Farmacia AA → Bodega AA ────────────────────────────

HDRS1 = [
    ('Medicamento',              44),
    ('Criticidad',               12),
    ('Stock Farm. AA (ud)',      14),   # Stock_Farm_Actual
    ('Cob. actual (días)',       13),   # Stock / CDL_Trend
    ('CDL Tend. (ud/d)',         12),
    ('A Pedir (ud)',             13),   # redondeado al Fe ICP CENABAST (invisible)
    ('Farm → Bod AA',            34),   # traspaso Bodega AA → Farmacia
    ('Bod AA ← Bod.Fármacos',   34),   # reposición que Bodega AA necesita de BodFarm
]

def calc_h1(df_farm, fe_map, dias_ef, todos=False):
    rows = []
    for _, r in df_farm.iterrows():
        med   = str(r.get('Medicamento', '')).strip()
        crit  = str(r.get('Criticidad', '5-OK'))
        nv    = _nivel(crit)
        # CDL de tendencia (Consumo_5D_Trend ya tiene factor de carga semanal aplicado)
        trend5d = _n(r.get('Consumo_5D_Trend', 0))
        cdl     = (trend5d / 5) if trend5d > 0 else _n(r.get('CDL_DiasHab', 0))
        sfarm   = int(_n(r.get('Stock_Farm_Actual', 0)))
        sbod    = int(_n(r.get('Stock_Bodega_Disponible', 0)))

        # Cobertura actual (antes del pedido)
        cob_actual = round(sfarm / cdl, 1) if cdl > 0 else 0.0

        ss_dias = BUFFER_SS + (EXTRA_CRIT if nv <= 2 else 0)
        ss      = math.ceil(cdl * ss_dias) if cdl > 0 else 0
        raw     = max(0, math.ceil(cdl * dias_ef) + ss - sfarm) if cdl > 0 else 0

        # Redondear al múltiplo del factor de empaque ICP CENABAST (invisible)
        fe  = int(fe_map.get(med, 1)) or 1
        ud  = _ceil_fe(raw, fe)

        if not todos and ud <= 0:
            continue

        if ud <= 0:
            accion1, accion2 = '', ''
        elif sbod >= ud:
            accion1 = f'Traspasar {ud} ud → Farmacia'
            accion2 = ''
        elif sbod > 0:
            accion1 = f'Traspasar {sbod} ud → Farmacia'
            accion2 = f'Reponer Bod.AA: {ud - sbod} ud de Bod.Fármacos'
        else:
            accion1 = 'SIN STOCK en Bodega AA'
            accion2 = f'COMPRA URGENTE {ud} ud de Bod.Fármacos'

        rows.append({
            'v': (med, crit, sfarm, cob_actual, round(cdl, 1), ud, accion1, accion2),
            '_nv': nv, '_ud': ud,
        })

    rows.sort(key=lambda x: x['v'][0])
    return [x['v'] for x in rows]


def write_h1(ws, rows, dias_ef, hoy, semana):
    ndia = ['Lun','Mar','Mié','Jue','Vie','Sáb','Dom'][hoy.weekday()]
    _titulo(ws,
        f'FARMACIA AA → BODEGA AA  ·  {ndia} {hoy.strftime("%d/%m/%Y")}  ·  '
        f'S{semana}  ·  {dias_ef} día(s) hábil(es) restante(s)',
        len(HDRS1))
    _subtit(ws,
        f'Stock Farm.AA = stock actual en Farmacia | Cob.actual = Stock/CDL (días antes del pedido) | '
        f'A Pedir = CDL×{dias_ef}d+SS−Stock, redondeado al empaque CENABAST | '
        f'Col 7: traspaso Farm←Bod.AA | Col 8: reposición Bod.AA←Bod.Fármacos '
        f'(rosa = sin stock, ámbar = compra urgente)',
        len(HDRS1))
    hdrs = list(HDRS1)
    hdrs[5] = (f'A Pedir ({dias_ef}d hab., ud)', 13)
    _hdr(ws, 3, hdrs)
    for i, vals in enumerate(rows, 4):
        # vals: (med, crit, sfarm, cob_actual, cdl, ud, accion1, accion2)
        _fila_crit(ws, i, vals, vals[1], {2, 3, 4, 5, 6}, cols_fmt1d={4, 5})
        a1, a2 = str(vals[6]), str(vals[7])
        if 'SIN STOCK' in a1:
            c = ws.cell(i, 7)
            c.fill = _pfill('FFDAD6')
            c.font = Font(name='Arial', size=10, color='9B1C1C', bold=True)
        if 'COMPRA URGENTE' in a2:
            c = ws.cell(i, 8)
            c.fill = _pfill('FEF08A')
            c.font = Font(name='Arial', size=10, color='854D0E', bold=True)
        elif 'Reponer' in a2:
            c = ws.cell(i, 8)
            c.fill = _pfill('FFF9C4')
            c.font = Font(name='Arial', size=10, color='78350F')
    ws.freeze_panes = 'A4'
    if rows:
        last = 3 + len(rows)
        ws.auto_filter.ref = f'A3:{get_column_letter(len(HDRS1))}{last}'
        _totals(ws, 4, last, len(HDRS1), {6})
        ws.print_area = f'A1:{get_column_letter(len(HDRS1))}{last + 1}'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = 'landscape'


# ─────────────── Hoja 2: Bodega AA → Bodega Fármacos ────────────────────────

HDRS2 = [
    ('Medicamento',              44),
    ('Criticidad',               12),
    ('Stock Bod. AA',            12),
    ('Stock Farm. AA',           12),
    ('Stock Bod. Fármacos',      14),
    ('Req. ciclo (ud)',          13),   # CDL × dias_ciclo; header dinámico en write_h2
    ('A Reponer (ud)',           13),   # max(0, req_ciclo - (sbod+sfarm)), redondeado al Fe
    ('Accion',                   44),
]

def calc_h2(df_bod, fe_map, hoy, fer):
    dc = _dias_ciclo(hoy, fer)
    rows = []
    for _, r in df_bod.iterrows():
        med    = str(r.get('Medicamento', '')).strip()
        crit   = str(r.get('Criticidad', '5-OK'))
        sbod   = int(_n(r.get('Stock_Bod_Actual', 0)))
        sfarm  = int(_n(r.get('Stock_Farm_Actual', 0)))
        sbfarm = int(_n(r.get('Stock_BODEGA_FARMACOS', 0)))
        cons10 = _n(r.get('Consumo_10D_Trend', 0))
        req2   = _n(r.get('Req_2_Semanas', 0))

        # CDL de tendencia (Consumo_10D_Trend ya trae el ajuste por semana del mes,
        # igual que Consumo_5D_Trend en Farm_Bod); cae al req plano de 2 semanas
        # (10d hábiles) solo si no hay dato de tendencia.
        cdl = (cons10 / 10) if cons10 > 0 else (req2 / 10 if req2 > 0 else 0.0)

        # Stock requerido para el ciclo actual (días restantes). Se descuenta TODO
        # el stock de Atención Abierta (Bodega AA + Farmacia AA) — el de Bodega
        # Fármacos vive en otra ubicación física y solo decide la acción (traspaso
        # vs. compra externa), no reduce la necesidad.
        req_ciclo   = math.ceil(cdl * dc) if cdl > 0 else 0
        fe          = int(fe_map.get(med, 1)) or 1
        rep         = _ceil_fe(max(0, req_ciclo - (sbod + sfarm)), fe)

        if rep <= 0:
            accion = ''
        elif sbfarm >= rep:
            accion = f'Pedir {rep} ud a Bod.Fármacos'
        else:
            falt = rep - sbfarm
            accion = f'Bod.Fármacos: {sbfarm} ud disponibles | COMPRA EXTERNA: {falt} ud'

        rows.append({
            'v': (med, crit, sbod, sfarm, sbfarm, req_ciclo, rep, accion),
            '_nv': _nivel(crit), '_rep': rep,
        })

    rows.sort(key=lambda x: x['v'][0])
    return dc, [x['v'] for x in rows]


def write_h2(ws, rows, hoy, semana, dias_ciclo):
    _titulo(ws,
        f'BODEGA AA → BODEGA FÁRMACOS  ·  Ciclo {dias_ciclo}d hábiles restantes  ·  '
        f'{hoy.strftime("%d/%m/%Y")}  ·  S{semana}',
        len(HDRS2))
    _subtit(ws,
        f'Ciclo 10d hábiles (inicio 29-jun, repite c/2 semanas) | Quedan {dias_ciclo}d | '
        f'A Reponer = max(0, CDL×{dias_ciclo}d − (Stock Bod.AA + Stock Farm.AA)), '
        f'redondeado al empaque CENABAST | '
        f'Stock Bod.Fármacos solo decide la acción: pedir traspaso o compra externa | '
        f'Ámbar = compra externa a Bod.Fármacos',
        len(HDRS2))
    hdrs = list(HDRS2)
    hdrs[5] = (f'Req. ciclo ({dias_ciclo}d, ud)', 13)
    _hdr(ws, 3, hdrs)
    for i, vals in enumerate(rows, 4):
        _fila_crit(ws, i, vals, vals[1], {2, 3, 4, 5, 6, 7})
        ac = str(vals[7])
        if 'COMPRA EXTERNA' in ac:
            c = ws.cell(i, 8)
            c.fill = _pfill('FEF08A')
            c.font = Font(name='Arial', size=10, color='854D0E', bold=True)
    ws.freeze_panes = 'A4'
    if rows:
        last = 3 + len(rows)
        ws.auto_filter.ref = f'A3:{get_column_letter(len(HDRS2))}{last}'
        _totals(ws, 4, last, len(HDRS2), {7})
        ws.print_area = f'A1:{get_column_letter(len(HDRS2))}{last + 1}'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = 'landscape'


# ─────────────── Hoja 3: Diálisis (Farm + Bod) ──────────────────────────────

# Fe ICP usado internamente para redondear mensual, no como columna visible
HDRS3 = [
    ('Medicamento',              44),
    ('Cons. Mensual Diál. (ud)', 18),
    ('Cob.Farm (mes)',           12),   # Stock Farm / Cons.Mensual
    ('A Pedir Farm. (ud)',       14),
    ('Cob.Bod (mes)',            12),   # Stock Bod / Cons.Mensual
    ('A Pedir Bod. (ud)',        14),
    ('Obs',                      46),
]

def calc_h3(df_dfarm, df_dbod, fe_map, rep_h2_map=None):
    rep_h2_map = rep_h2_map or {}

    def _idx(df, stock_col):
        d = {}
        for _, r in df.iterrows():
            m = str(r.get('Medicamento', '')).strip()
            d[m] = {
                'c5d'   : _n(r.get('Consumo_5D_Solo_Dialisis', 0)),
                'stock' : int(_n(r.get(stock_col, 0))),
                'fe'    : int(_n(r.get('Factor_Empaque', 1))) or 1,
            }
        return d

    fd = _idx(df_dfarm, 'Stock_Farm_Actual')
    bd = _idx(df_dbod,  'Stock_Bod_Actual')
    meds = sorted(set(list(fd.keys()) + list(bd.keys())))

    rows = []
    for med in meds:
        f = fd.get(med, {})
        b = bd.get(med,  {})
        c5d = f.get('c5d', 0) or b.get('c5d', 0)
        if c5d <= 0:
            continue
        # Unidades_Caja de SGLI_Estres (cenabast_tallas.csv, curado a mano) es la
        # fuente confiable — el Factor_Empaque propio de la hoja Dialisis viene de
        # cenabast_intermediacion.csv con un matching de nombre mas debil y cae a 1
        # (sin redondeo) para muchos medicamentos que si tienen empaque conocido.
        fe = int(fe_map.get(med, 0)) or f.get('fe', 1) or b.get('fe', 1) or 1
        mensual  = _ceil_fe(c5d / 5 * 30, fe)
        # Si el medicamento no aparece en Dialisis_Pedido_Farm/Bod, maestro_aa.py ya
        # determino que ese nivel tiene stock suficiente (Necesidad_Farm/Bod<=0) —
        # no tratar la ausencia como stock=0, o se genera un pedido falso.
        sfarm    = f.get('stock', 0) if med in fd else None
        sbod     = b.get('stock', 0) if med in bd else None
        apfarm   = _ceil_fe(max(0, mensual - sfarm), fe) if sfarm is not None else 0

        # El CDL de Bod_Farmacos ya es COMBINADO (incluye diálisis, ver calc_h2),
        # así que el "A Reponer" de esa hoja ya trae en camino parte (o todo) de lo
        # que diálisis necesita de Bodega AA. Si ambas hojas se piden la misma
        # semana (coincide con S3), sin netear se pediría dos veces a Bod.Fármacos
        # contra el mismo stock base. Se descuenta el rep del ciclo antes de pedir.
        rep_bod  = int(rep_h2_map.get(med, 0))
        sbod_proy = (sbod + rep_bod) if sbod is not None else None
        apbod    = _ceil_fe(max(0, mensual - sbod_proy), fe) if sbod_proy is not None else 0
        cob_farm = round(sfarm / mensual, 1) if (sfarm is not None and mensual > 0) else None
        cob_bod  = round(sbod  / mensual, 1) if (sbod  is not None and mensual > 0) else None

        obs = []
        if apfarm > 0: obs.append(f'Farm: solicitar {apfarm} ud a Bodega AA')
        if apbod  > 0: obs.append(f'Bod:  solicitar {apbod} ud a Bod.Fármacos')
        if rep_bod > 0 and (apbod > 0 or max(0, mensual - (sbod or 0)) > 0):
            obs.append(f'(neteado con {rep_bod} ud que ya trae el ciclo Bod_Farmacos)')
        if not obs:    obs = ['Stock suficiente en ambos niveles']

        # (med, mensual, cob_farm, apfarm, cob_bod, apbod, obs) — Fe invisible, aplicado en mensual
        rows.append((med, mensual, cob_farm, apfarm, cob_bod, apbod, ' / '.join(obs)))

    rows.sort(key=lambda x: x[0])
    return rows


def write_h3(ws, rows, hoy, semana, activa):
    _titulo(ws,
        f'DIÁLISIS MENSUAL — Farm. AA + Bodega AA  ·  {hoy.strftime("%d/%m/%Y")}  ·  S{semana}',
        len(HDRS3))
    if not activa:
        _subtit(ws,
            f'⚠  Pedido de diálisis solo en S3. Semana actual: S{semana}. '
            f'Usa --forzar-dialisis para generar fuera de S3.',
            len(HDRS3))
        ws.merge_cells(f'A4:{get_column_letter(len(HDRS3))}4')
        ws.cell(4, 1, 'Sin datos — ejecutar en S3 o con --forzar-dialisis').font = \
            Font(italic=True, size=10, color='B45309', name='Arial')
        return
    _subtit(ws,
        'Cons.Mensual = C5D_Diál÷5×30, redondeado al empaque CENABAST (días naturales) | '
        'Cob.Farm/Bod = Stock÷Cons.Mensual (en meses, sobre stock actual) | '
        'A Pedir Bod. = max(0, Mensual − (Stock Bod.AA + lo que ya trae el ciclo Bod_Farmacos)) '
        '— evita pedir dos veces a Bod.Fármacos contra el mismo stock',
        len(HDRS3))
    _hdr(ws, 3, HDRS3)
    for i, vals in enumerate(rows, 4):
        # cob_farm@col3 y cob_bod@col5 en meses → 0.0
        _fila_dial(ws, i, vals, {2, 3, 4, 5, 6}, cols_fmt1d={3, 5})
    ws.freeze_panes = 'A4'
    if rows:
        last = 3 + len(rows)
        ws.auto_filter.ref = f'A3:{get_column_letter(len(HDRS3))}{last}'
        _totals(ws, 4, last, len(HDRS3), {2, 4, 6})
        ws.print_area = f'A1:{get_column_letter(len(HDRS3))}{last + 1}'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = 'landscape'


# ─────────────── main ────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(
        description='Planilla simplificada Pedidos AA — 3 hojas, sin IA.')
    ap.add_argument('--forzar-dialisis', action='store_true',
                    help='Genera hoja Diálisis aunque no sea S3')
    ap.add_argument('--todos', action='store_true',
                    help='Incluye meds sin necesidad de pedido en Farm_Bod '
                         '(Bod_Farmacos siempre lista el universo completo)')
    args = ap.parse_args()

    hoy  = dt.date.today()
    fer  = _feriados()
    mae  = _maestro()
    sem  = _semana(hoy)
    def_ = _dias_ef(hoy, fer)
    ndia = ['Lun','Mar','Mié','Jue','Vie','Sáb','Dom'][hoy.weekday()]

    print(f'HOY = {hoy} ({ndia}) | S{sem} | dias hab. restantes: {def_}')
    print(f'Maestro: {os.path.basename(mae)}\n')

    data = _leer(mae)

    # fe_map: Medicamento → Unidades_Caja (ICP CENABAST desde SGLI_Estres)
    fe_map = {}
    if len(data['sgli']):
        for _, r in data['sgli'].iterrows():
            m = str(r.get('Medicamento', '')).strip()
            fe_map[m] = int(_n(r.get('Unidades_Caja', 1))) or 1

    r1 = calc_h1(data['farm'], fe_map, def_, args.todos)
    dc, r2 = calc_h2(data['bod'], fe_map, hoy, fer)
    dial_activa = args.forzar_dialisis or sem == 3
    # rep_h2_map: lo que la hoja Bod_Farmacos ya trae para cada med (CDL combinado,
    # incluye diálisis) — se usa para netear el pedido de diálisis a Bod.Fármacos
    # y no pedir dos veces contra el mismo stock cuando ambas hojas salen la misma semana.
    rep_h2_map = {v[0]: v[6] for v in r2}
    r3 = calc_h3(data['dfarm'], data['dbod'], fe_map, rep_h2_map) if dial_activa else []

    wb = openpyxl.Workbook()
    ws1 = wb.active; ws1.title = 'Farm_Bod'
    write_h1(ws1,                             r1, def_, hoy, sem)
    write_h2(wb.create_sheet('Bod_Farmacos'), r2, hoy, sem, dc)
    write_h3(wb.create_sheet('Dialisis'),     r3, hoy, sem, dial_activa)

    sal = os.path.join(WORK_DIR,
        f'Pedido_Fusion_AA_{hoy.strftime("%Y%m%d_%H%M")}.xlsx')
    wb.save(sal)

    # índices: h1=(med,crit,sfarm,cob_actual,cdl,ud,accion1,accion2)            → ud@5
    #          h2=(med,crit,sbod,sfarm,sbfarm,req_ciclo,rep,accion)            → rep@6
    #          h3=(med,mensual,cob_farm,apfarm,cob_bod,apbod,obs)              → apfarm@3, apbod@5
    n1 = sum(1 for v in r1 if v[5] > 0)
    n2 = sum(1 for v in r2 if v[6] > 0)
    n3 = sum(1 for v in r3 if (v[3] + v[5]) > 0)
    print(f'Farm->Bod       : {len(r1)} meds ({n1} con pedido)')
    print(f'Bod->Farmacos   : {len(r2)} meds ({n2} con reposicion)')
    if dial_activa:
        print(f'Dialisis        : {len(r3)} meds ({n3} con faltante)')
    else:
        print(f'Dialisis        : fuera de S3 (usa --forzar-dialisis)')
    print(f'\nExcel: {os.path.basename(sal)}')


if __name__ == '__main__':
    main()
