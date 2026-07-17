"""sgli.py — Motor SGLI (Sistema de Gestion Logistica) de Farmacia AA.

Fuente UNICA de verdad de la reposicion **basada en demanda** (sin techo de
capacidad fisica), usada por:
  - maestro_aa.py  -> persiste la hoja 'SGLI_Estres' en el Consolidado.
  - app_pedidos.py -> pestana 'SGLI / Capacidad' que recalcula EN VIVO al mover
                      el Factor_Carga.

Reglas (spec Farmacia AA, Hospital Pitrufquen):
  1. Limite fisico (Capacidad_Max), EN UNIDADES:
       Base_cajas = TALLA(S=23 | M=15 | L=9 cajas/gaveta) o ⌊2790/Vol_Caja⌋ (EXTERNO).
       Cap_cajas  = Base_cajas × (3 si Rotacion ALTA | 1 si NORMAL).
       Cap_Max    = Cap_cajas × Unidades_por_Caja.
     (La farmacia ALMACENA Y FRACCIONA POR UNIDAD, no por empaque -> todo el
      modelo trabaja en unidades; las cajas solo dan el volumen ocupado.)
  2. Intervalo de reposicion dinamico (IR):
       IR = max(1, min(5, ⌊5/Factor_Carga⌋))
  3. Nivel objetivo (T) y deficit:
       Demanda = IR × CDL × Factor_Carga
       T       = ⌊Demanda × 1.25⌋
       Deficit = max(0, T − Stock_Farm)
  4. Arbol de decision (solo si Deficit > 0):
       Stock_Bod >= Deficit     -> TRASPASAR [Deficit] DESDE BODEGA   | --
       0 < Stock_Bod < Deficit  -> TRASPASAR [Stock_Bod] DESDE BODEGA | COMPRA URGENTE [Deficit-Stock_Bod]
       Stock_Bod == 0           -> SIN STOCK BODEGA                   | COMPRA URGENTE [Deficit]

Orden de salida: 1) criticidad (1-CRITICO primero); 2) Deficit (desc).

Talla por medicamento: no existe en los datos -> se asigna por FORMA FARMACEUTICA
(promedios de caja del mercado chileno) y se puede sobre-escribir por medicamento
en `cenabast_tallas.csv` (col: Medicamento;Talla;Vol_Caja_cm3;Unidades_Caja).
Para 'EXTERNO' la fila trae Vol_Caja_cm3 (dimensiones investigadas) y/o Unidades_Caja.
Unidades_por_Caja usa el Factor_Empaque CENABAST real cuando se conoce; si no, el
promedio por forma; si la fila de override lo trae, ese manda.
"""
import math
import os
import glob

import numpy as np
import pandas as pd

from aa_colors import crit_nivel

# ── Constantes fisicas / del modelo ───────────────────────────────────────────
VOLUMEN_GAVETA_CM3       = 2790
TALLA_BASE_CAJAS         = {'S': 23, 'M': 15, 'L': 9}     # cajas por 1 gaveta
ROTACION_GAVETAS         = {'ALTA': 3, 'NORMAL': 1}
UMBRAL_ALTA_ROTACION_CDL = 30      # ud/dia habil -> ALTA (igual que maestro_aa)
FACTOR_CARGA_DEFAULT     = 1.15
MARGEN_SEGURIDAD         = 1.25
IR_MAX                   = 5
TALLA_DEFAULT            = 'M'
_EPS                     = 1e-9   # evita que floor() caiga 1 unidad por error de coma flotante

# ── Frecuencia de revisión por medicamento (Stock_Minimo) ─────────────────────
# CDL bucket → dias habiles entre revisiones (1=diario, 5=semanal).
# Umbrales derivados de la distribucion EMPIRICA del AA (404 meds con consumo),
# calibrados para que cada bucket contenga ~20% de los medicamentos:
#
#   CDL >= 150  ud/dia  ->  Bucket 1 (diario)      ~78 meds / 19 %
#   CDL  30-150         ->  Bucket 2 (cada 2 dias)  ~81 meds / 20 %
#   CDL   5-30          ->  Bucket 3 (cada 3 dias)  ~81 meds / 20 %
#   CDL  0.5-5          ->  Bucket 4 (cada 4 dias)  ~75 meds / 19 %
#   CDL  < 0.5          ->  Bucket 5 (semanal)      ~89 meds / 22 %
#
# UMBRAL_ALTA_ROTACION_CDL=30 permanece separado: solo define capacidad fisica
# de gaveta (1 vs 3 gavetas), NO la frecuencia de revision.
FREQ_CDL_UMBRALES = [(150, 1), (30, 2), (5, 3), (0.5, 4), (0, 5)]
BUFFER_FINDE_DIAS = 1   # dias de SS para blindar reapertura del lunes (demanda sab/dom ~0)
EXTRA_SS_CRITICOS = 1   # dias SS adicionales para criticidad nivel 1-2 (no pueden quebrar)

# Promedios por FORMA farmaceutica: (Talla, Unidades_por_Caja).
# Talla = bucket por VOLUMEN tipico de la caja comercial chilena, comparado con
# los volumenes de referencia S~=120 / M~=180 / L~=300 cm3 (-> 23/15/9 cajas por
# gaveta de 2790 cm3). Unidades_por_Caja = pack CENABAST habitual; el
# Factor_Empaque real manda cuando se conoce (ver calcular_sgli). El volumen de
# caja estimado va al final de cada linea para que sea facil ajustar la talla.
FORMA_DEFAULTS = {
    'SOLIDO_ORAL': ('S', 30),   # comprimido/capsula, caja blister     (~100 cm3)
    'OFTALMICO':   ('S', 1),    # colirio/gotas/otico, frasco gotario  (~85 cm3)
    'OVULO':       ('S', 10),   # ovulo/supositorio                    (~110 cm3)
    'TOPICO':      ('M', 1),    # crema/unguento/gel, tubo             (~170 cm3)
    'INYECTABLE':  ('M', 10),   # ampolla/vial/frasco-ampolla          (~175 cm3)
    'SOBRE':       ('M', 20),   # sobre/sachet                         (~190 cm3)
    'PARCHE':      ('M', 5),    # parche transdermico, caja ancha      (~200 cm3)
    'LIQUIDO_ORAL':('L', 1),    # jarabe/suspension/solucion, frasco   (~300 cm3)
    'INHALADOR':   ('L', 1),    # inhalador/aerosol                    (~285 cm3)
    'INSULINA':    ('L', 5),    # insulina/lapicera, caja de pens      (~440 cm3)
    'OTRO':        ('M', 15),
}

# Columnas finales, en el orden del spec (+ contexto al final).
COLUMNAS_SGLI = [
    'Medicamento', 'Criticidad', 'Freq_Revision', 'Stock_Minimo',
    'Dias_Reposicion_IR', 'Cap_Max',
    'Nivel_Objetivo_T', 'Deficit', 'Accion_1_Traspaso', 'Accion_2_Externa',
    'Talla', 'Rotacion', 'Unidades_Caja', 'Vol_Caja_cm3', 'Stock_Farm', 'Stock_Bod', 'CDL',
]

_ETIQUETA_NIVEL = {1: '1-CRITICO', 2: '2-ALTO', 3: '3-MODERADO', 4: '4-BAJO', 5: '5-OK'}


def calcular_freq_revision(cdl: float, criticidad=None, variabilidad: float = 0.0) -> int:
    """Frecuencia de revision en dias habiles por medicamento (1=diario, 5=semanal).

    Algoritmo multicriteria:
      1. CDL -> bucket base (FREQ_CDL_UMBRALES, calibrado con distribucion real AA)
           >= 150 ud/dia -> 1 (diario)      | 30-150  -> 2 | 5-30    -> 3
             0.5-5 ud/dia -> 4 (cada 4 dias) | < 0.5   -> 5 (semanal)
      2. Criticidad nivel 1-2 -> -1 (no pueden quebrar stock)
      3. Variabilidad CDL > 50% vs promedio historico -> -1 (demanda irregular)

    Args:
        cdl:          consumo diario laboral (ud/dia habil), promedio del periodo
        criticidad:   etiqueta '1-CRITICO', '2-ALTO', etc. (o None)
        variabilidad: CV proxy = |CDL_reciente - CDL| / CDL; 0 si no disponible

    Returns:
        int 1-5 (dias habiles entre revisiones)
    """
    if cdl <= 0:
        return 5
    bucket = next((d for u, d in FREQ_CDL_UMBRALES if cdl >= u), 5)
    nivel = crit_nivel(criticidad) if criticidad is not None else 5
    if nivel <= 2:
        bucket -= 1
    if variabilidad > 0.5:
        bucket -= 1
    return max(1, min(5, bucket))


def detectar_forma(nombre: str) -> str:
    """Forma farmaceutica deducida del nombre. Normaliza separadores (. / ( ))
    y evalua de la forma mas especifica a la mas generica. 'SOLIDO_ORAL' es la
    red de seguridad para cualquier med con concentracion (MG/MCG/G/UI) sin
    marcador de liquido/inyectable; 'OTRO' queda solo para insumos/dispositivos
    sin concentracion."""
    s = ' ' + ' '.join(str(nombre).upper()
                       .replace('.', ' ').replace('/', ' ')
                       .replace('(', ' ').replace(')', ' ').split()) + ' '
    # Inhalador / nebulizacion / aerocamara
    if any(k in s for k in (' INHALADOR ', ' INHALA', ' AEROSOL ', ' NEBULIZAR ',
                            ' P NEBULIZAR ', ' POLVO INHAL', ' AEROCAMARA ')):
        return 'INHALADOR'
    # Oftalmico / otico
    if any(k in s for k in (' COLIRIO ', ' COLIRI ', ' OFTAL', ' OFT ', ' OFTALMIC',
                            ' OTICO ', ' OTIC ', ' LAGRIMAS ')):
        return 'OFTALMICO'
    # Insulina / lapicera (pen)
    if any(k in s for k in (' INSULINA ', ' LAPICERA ', ' LAPIZ ')):
        return 'INSULINA'
    # Inyectable: ampolla / vial / jeringa / vias parenterales / diluyentes
    if any(k in s for k in (' AMPOLLA ', ' AMP ', ' AM ', ' INYECTABLE ', ' INY ', ' VIAL ',
                            ' FA ', ' FCO AMP', ' JERINGA ', ' JER ', ' JRP ', ' PC ',
                            ' INTRAMUSCULAR ', ' IM ', ' EV ', ' SC ', ' LIOFIL',
                            ' MATRAZ ', ' MOLAR ')):
        return 'INYECTABLE'
    # Topico
    if any(k in s for k in (' CREMA ', ' UNGUENTO ', ' UNGUEN', ' POMADA ', ' GEL ',
                            ' PASTA ', ' SHAMPOO ', ' LOCION ')):
        return 'TOPICO'
    # Ovulo / supositorio / parche / sobre
    if any(k in s for k in (' OVULO ', ' SUPOSITORIO ')):
        return 'OVULO'
    if ' PARCHE ' in s:
        return 'PARCHE'
    if any(k in s for k in (' SOBRE ', ' SACHET ', ' SCH ')):
        return 'SOBRE'
    # Liquido oral: jarabe / suspension / solucion / gotas / frasco / dosificado en ML
    if any(k in s for k in (' JARABE ', ' SUSPENSION ', ' SUSP ', ' ELIXIR ', ' GOTAS ',
                            ' SOLUCION ', ' SOL ', ' FRASCO ', ' FC ', ' EMULSION ', ' ML ')):
        return 'LIQUIDO_ORAL'
    # Solido oral (lo mas comun): comprimido / capsula y abreviaturas
    if any(k in s for k in (' COMPRIMIDO ', ' COMP ', ' COMPR', ' CAPSULA ', ' CAP ',
                            ' CM ', ' CMR ', ' CR ', ' CP ', ' GRAGEA ', ' GRAG ',
                            ' TABLETA ', ' TAB ', ' UD ')):
        return 'SOLIDO_ORAL'
    # Red de seguridad: concentracion solida sin marcador liquido/inyectable -> comprimido
    if any(k in s for k in (' MG ', ' MCG ', ' G ', ' UI ', ' UNIDAD ')):
        return 'SOLIDO_ORAL'
    return 'OTRO'


def _clave(nombre) -> str:
    return str(nombre).strip().upper()


def cargar_tallas(work_dir):
    """Lee `cenabast_tallas.csv` (overrides por medicamento). Devuelve
    {clave_med: {'Talla':str|None, 'Vol_Caja_cm3':float|None, 'Unidades_Caja':int|None}}.
    Tolerante a separador ';' o ','. Si no existe, devuelve {}."""
    path = os.path.join(work_dir, 'cenabast_tallas.csv')
    if not os.path.exists(path):
        alt = glob.glob(os.path.join(work_dir, 'cenabast_tallas*.csv'))
        if not alt:
            return {}
        path = max(alt, key=os.path.getmtime)
    df = None
    for sep in (';', ','):
        try:
            tmp = pd.read_csv(path, sep=sep, dtype=str, encoding='latin1').dropna(how='all')
        except Exception:
            continue
        if tmp.shape[1] >= 2:
            df = tmp
            break
    if df is None or 'Medicamento' not in df.columns:
        return {}

    def _num(v):
        try:
            v = str(v).replace(',', '.').strip()
            return float(v) if v not in ('', 'nan', 'None') else None
        except Exception:
            return None

    out = {}
    for _, r in df.iterrows():
        med = r.get('Medicamento')
        if not isinstance(med, str) or not med.strip():
            continue
        talla = str(r.get('Talla', '') or '').strip().upper() or None
        out[_clave(med)] = {
            'Talla': talla,
            'Vol_Caja_cm3': _num(r.get('Vol_Caja_cm3')),
            'Unidades_Caja': (int(_num(r.get('Unidades_Caja')))
                              if _num(r.get('Unidades_Caja')) else None),
        }
    return out


def _base_cajas(talla, vol_caja):
    """Cajas que caben en 1 gaveta. Vol_Caja (cm3) manda si viene; si no, la talla."""
    if vol_caja and vol_caja > 0:
        return max(1, math.floor(VOLUMEN_GAVETA_CM3 / vol_caja + _EPS))
    return TALLA_BASE_CAJAS.get(str(talla).upper(), TALLA_BASE_CAJAS[TALLA_DEFAULT])


def _accion_logistica(deficit_i: int, stock_bod_i: int):
    if deficit_i <= 0:
        return '', ''
    if stock_bod_i >= deficit_i:
        return f'TRASPASAR {deficit_i} DESDE BODEGA', ''
    if stock_bod_i > 0:
        return (f'TRASPASAR {stock_bod_i} DESDE BODEGA',
                f'COMPRA URGENTE {deficit_i - stock_bod_i}')
    return 'SIN STOCK BODEGA', f'COMPRA URGENTE {deficit_i}'


def _resolver_fisico(nombre, overrides, factor_empaque,
                     talla_row, rotacion_row, upb_row, vol_row, cdl):
    """Devuelve (talla, rotacion, vol_caja, unidades_caja) para una fila,
    combinando override de archivo > columnas ya persistidas > heuristica."""
    ov = overrides.get(_clave(nombre), {}) if overrides else {}
    forma = detectar_forma(nombre)
    talla_f, upb_f = FORMA_DEFAULTS.get(forma, FORMA_DEFAULTS['OTRO'])

    # Talla: override > columna persistida > forma
    talla_row_ok = talla_row if (isinstance(talla_row, str) and talla_row.strip()) else None
    talla = ov.get('Talla') or talla_row_ok or talla_f
    talla = str(talla).upper()

    # Volumen de caja (para EXTERNO o precision): override > columna persistida
    vol = ov.get('Vol_Caja_cm3')
    if vol is None and vol_row is not None and not (isinstance(vol_row, float) and np.isnan(vol_row)):
        try:
            vol = float(vol_row) or None
        except Exception:
            vol = None

    # Rotacion: columna persistida > derivada de CDL
    if rotacion_row and str(rotacion_row).upper() in ROTACION_GAVETAS:
        rotacion = str(rotacion_row).upper()
    else:
        rotacion = 'ALTA' if (cdl or 0) >= UMBRAL_ALTA_ROTACION_CDL else 'NORMAL'

    # Unidades por caja: override > columna persistida > Factor_Empaque (si >1) > forma
    upb = ov.get('Unidades_Caja')
    if upb is None and upb_row is not None and not (isinstance(upb_row, float) and np.isnan(upb_row)):
        try:
            upb = int(float(upb_row)) or None
        except Exception:
            upb = None
    if upb is None:
        fe = 0
        try:
            fe = int(float(factor_empaque)) if factor_empaque is not None else 0
        except Exception:
            fe = 0
        upb = fe if fe > 1 else upb_f
    return talla, rotacion, vol, max(1, int(upb))


def calcular_sgli(df, factor_carga=FACTOR_CARGA_DEFAULT, *, overrides=None,
                  col_med='Medicamento', col_crit='Criticidad',
                  col_farm='Stock_Farmacia_AA', col_bod='Stock_Bodega_AA',
                  col_cdl='CDL', col_factor_empaque='Factor_Empaque',
                  col_talla='Talla', col_rotacion='Rotacion',
                  col_upb='Unidades_Caja', col_vol='Vol_Caja_cm3',
                  col_cdl_pond=None):
    """Aplica el motor SGLI (capacidad fisica) a ``df`` -> DataFrame COLUMNAS_SGLI.

    ``factor_carga`` es global (un valor para todos). Si ``df`` ya trae Talla/
    Rotacion/Unidades_Caja (hoja persistida), se reutilizan; si no, se derivan
    (forma + CDL + Factor_Empaque + overrides). Cap_Max NO depende de
    factor_carga, asi que el recalculo en vivo de la app es estable.

    ``col_cdl_pond`` (opcional): columna con CDL ponderado reciente (p.ej. 'CDL_Pond'
    de maestro_aa). Si se pasa, se usa para calcular la variabilidad del consumo
    (|CDL_pond - CDL| / CDL) que puede reducir el Freq_Revision en 1 nivel.
    """
    if df is None or len(df) == 0:
        return pd.DataFrame(columns=COLUMNAS_SGLI)

    fc = float(factor_carga)
    ir_factor_cap = max(1, min(IR_MAX, math.floor(IR_MAX / fc + _EPS))) if fc > 0 else IR_MAX
    overrides = overrides or {}

    def _col(c):
        return df[c] if c in df.columns else pd.Series([None] * len(df), index=df.index)

    s_med      = df[col_med].astype(str).values
    s_crit     = _col(col_crit).values
    s_farm     = pd.to_numeric(_col(col_farm), errors='coerce').fillna(0).values
    s_bod      = pd.to_numeric(_col(col_bod),  errors='coerce').fillna(0).values
    s_cdl      = pd.to_numeric(_col(col_cdl),  errors='coerce').fillna(0).values
    s_fe       = _col(col_factor_empaque).values
    s_talla    = _col(col_talla).values
    s_rot      = _col(col_rotacion).values
    s_upb      = _col(col_upb).values
    s_vol      = _col(col_vol).values
    s_cdl_pond = (pd.to_numeric(_col(col_cdl_pond), errors='coerce').values
                  if col_cdl_pond and col_cdl_pond in df.columns else None)

    rows = []
    for i in range(len(df)):
        nombre = s_med[i]
        cdl = float(s_cdl[i])
        talla, rotacion, vol, upb = _resolver_fisico(
            nombre, overrides, s_fe[i], s_talla[i], s_rot[i], s_upb[i], s_vol[i], cdl)

        base_cajas = _base_cajas(talla, vol)
        cap_cajas  = base_cajas * ROTACION_GAVETAS[rotacion]
        cap_max    = int(cap_cajas * upb)                       # EN UNIDADES

        nivel = crit_nivel(s_crit[i])

        # Variabilidad CDL: proxy = |CDL_pond - CDL| / CDL (0 si no disponible)
        variabilidad = 0.0
        if s_cdl_pond is not None and cdl > 0:
            pond_i = s_cdl_pond[i]
            if pond_i == pond_i:  # no NaN
                variabilidad = abs(float(pond_i) - cdl) / cdl

        # Freq_Revision: per-medicamento, no depende del factor de carga
        freq_rev = calcular_freq_revision(cdl, s_crit[i], variabilidad)

        # Stock_Minimo (ROP per-medicamento): CDL * Freq_Rev + Stock_Seguridad
        ss_dias = BUFFER_FINDE_DIAS + (EXTRA_SS_CRITICOS if nivel <= 2 else 0)
        stock_minimo = math.ceil(cdl * freq_rev) + math.ceil(cdl * ss_dias) if cdl > 0 else 0

        if cdl > 0:
            ir = max(1, min(IR_MAX, ir_factor_cap))
            demanda = ir * cdl * fc
            t = math.floor(demanda * MARGEN_SEGURIDAD + _EPS)
        else:
            ir = ir_factor_cap
            t = 0
        stock_farm = int(round(s_farm[i]))
        deficit = max(0, int(t) - stock_farm)
        stock_bod = int(round(s_bod[i]))
        a1, a2 = _accion_logistica(deficit, stock_bod)

        rows.append({
            'Medicamento': nombre,
            '_nivel': nivel,
            'Criticidad': _ETIQUETA_NIVEL.get(nivel, '5-OK'),
            'Freq_Revision': freq_rev,
            'Stock_Minimo': stock_minimo,
            'Dias_Reposicion_IR': int(ir),
            'Cap_Max': int(cap_max),
            'Nivel_Objetivo_T': int(t),
            'Deficit': int(deficit),
            'Accion_1_Traspaso': a1,
            'Accion_2_Externa': a2,
            'Talla': talla,
            'Rotacion': rotacion,
            'Unidades_Caja': int(upb),
            'Vol_Caja_cm3': vol,
            'Stock_Farm': stock_farm,
            'Stock_Bod': stock_bod,
            'CDL': round(cdl, 2),
        })

    out = pd.DataFrame(rows)
    out.sort_values(['_nivel', 'Deficit'], ascending=[True, False], inplace=True)
    out.drop(columns=['_nivel'], inplace=True)
    return out.reset_index(drop=True)[COLUMNAS_SGLI]


def to_markdown(df) -> str:
    """Tabla Markdown con las columnas operativas del SGLI."""
    headers = ['Medicamento', 'Criticidad', 'Rev.(d)', 'Stock Min.', 'IR global',
               'Cap_Max', 'Nivel Obj.(T)', 'Deficit', 'Accion 1: Traspaso', 'Accion 2: Externa']
    src = ['Medicamento', 'Criticidad', 'Freq_Revision', 'Stock_Minimo', 'Dias_Reposicion_IR',
           'Cap_Max', 'Nivel_Objetivo_T', 'Deficit', 'Accion_1_Traspaso', 'Accion_2_Externa']
    lines = ['| ' + ' | '.join(headers) + ' |',
             '|' + '|'.join(['---'] * len(headers)) + '|']
    for _, r in df.iterrows():
        cells = [str(r.get(c, '')).replace('|', '/') for c in src]
        lines.append('| ' + ' | '.join(cells) + ' |')
    return '\n'.join(lines)
