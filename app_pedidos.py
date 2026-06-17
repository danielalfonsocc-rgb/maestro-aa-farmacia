import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import streamlit as st
import pandas as pd
import numpy as np
import os, io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from aa_colors import CRIT_FILL_HEX, crit_fill, crit_nivel, crit_hex

# PDF
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, HRFlowable)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ─────────────────────────────────────────────
# Ruta dinamica — funciona en cualquier PC Windows
WORK_DIR    = os.path.dirname(os.path.abspath(__file__))
XLS_MAESTRO = os.path.join(WORK_DIR, "Consolidado_AA_MAESTRO.xlsx")

st.set_page_config(
    page_title="Pedidos AA — Farmacia AT Abierta",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ═══════════════════════════════════════════════════════════════
   SISTEMA DE DISEÑO — paleta neutra, color solo para acción/estado
   Fondo gris muy claro · tarjetas blancas · texto carbón ·
   amplio espaciado · esquinas redondeadas · sombras suaves difusas
   ═══════════════════════════════════════════════════════════════ */

/* ── Base tipográfica + lienzo neutro ─────────────────────────── */
html, body, [class*="css"] {
    font-family:-apple-system,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;
    color:#1F2937;                       /* gray-800 carbón */
}
.stApp { background:#F9FAFB; }           /* gray-50 lienzo */
.block-container { padding-top:1.4rem; padding-bottom:3rem; max-width:1320px; }
h1,h2,h3,h4 { color:#111827; }           /* gray-900 títulos */
h2 { font-weight:800; letter-spacing:-.01em; }
hr { border:0; border-top:1px solid #EEF0F4; margin:1.4rem 0; }
a { color:#0F766E; }                     /* teal = enlace/acción */

/* ── Encabezado de marca (único bloque con color de marca) ────── */
.app-hero {
    display:flex; align-items:center; gap:18px;
    background:linear-gradient(90deg,#0F766E 0%,#0E7490 100%);
    border-radius:18px; padding:20px 26px; margin-bottom:18px; color:#fff;
    box-shadow:0 10px 30px -14px rgba(15,118,110,.55);
}
.app-hero .ico {
    width:54px; height:54px; border-radius:16px; background:rgba(255,255,255,.18);
    display:flex; align-items:center; justify-content:center; font-size:27px; flex:none;
}
.app-hero .tit { font-size:1.55rem; font-weight:800; line-height:1.1; margin:0; color:#fff; }
.app-hero .sub { font-size:.9rem; opacity:.93; margin:4px 0 0; }
.app-hero .chip {
    margin-left:auto; text-align:right; flex:none;
    background:rgba(255,255,255,.16); border-radius:12px; padding:9px 15px;
}
.app-hero .chip .d { font-size:.78rem; opacity:.9; }
.app-hero .chip .w { font-size:.95rem; font-weight:700; }

/* ── Leyenda de criticidad — badges semánticos suaves ─────────── */
.leyenda { display:flex; flex-wrap:wrap; gap:9px; margin:0 0 22px; }
.leyenda .lg {
    border-radius:999px; padding:5px 13px; font-size:.78rem; font-weight:600;
    background:#F3F4F6; color:#4B5563; border:1px solid #E5E7EB;
}

/* ── Tarjetas de medicamento — blancas, riel de color = estado ── */
.tarjeta {
    background:#fff; border:1px solid #EEF0F4; border-radius:16px;
    padding:20px 24px; margin-bottom:14px;
    box-shadow:0 1px 2px rgba(16,24,40,.04), 0 10px 24px -16px rgba(16,24,40,.18);
    border-left:5px solid #CBD5E1;
}
.t-crit1 { border-left-color:#DC2626; }
.t-crit2 { border-left-color:#EA580C; }
.t-crit3 { border-left-color:#F59E0B; }
.t-ok    { border-left-color:#16A34A; }

/* ── Bloque-med (pestaña Buscar) — blanco, riel de estado ─────── */
.bloque-med {
    background:#fff; border:1px solid #EEF0F4; border-left:5px solid #94A3B8;
    padding:16px 20px; border-radius:14px; margin-bottom:10px;
    box-shadow:0 1px 2px rgba(16,24,40,.04);
}
.crit-1  { border-left-color:#DC2626; }
.crit-2  { border-left-color:#EA580C; }
.crit-3  { border-left-color:#F59E0B; }
.crit-ok { border-left-color:#16A34A; }

/* ── Números y etiquetas ──────────────────────────────────────── */
.num-grande { font-size:2.05rem; font-weight:800; color:#111827; line-height:1; }
.etiqueta   { font-size:.76rem; color:#6B7280; margin-top:4px; }

/* ── Cajas de acción — badges suaves ──────────────────────────── */
.accion-box {
    background:#ECFDF5; border:1px solid #D1FAE5; border-radius:10px;
    padding:8px 14px; margin-top:8px; font-size:.86rem; color:#065F46; font-weight:600;
}
.accion-ext {
    background:#FFF1F2; border:1px solid #FECDD3; border-radius:10px;
    padding:8px 14px; margin-top:6px; font-size:.86rem; color:#9D174D; font-weight:600;
}

/* ── Badges semánticos reutilizables ──────────────────────────── */
.badge { display:inline-block; border-radius:999px; padding:2px 11px;
         font-size:.76rem; font-weight:700; line-height:1.5; }
.badge-red    { background:#FEF2F2; color:#B91C1C; }
.badge-orange { background:#FFF7ED; color:#C2410C; }
.badge-amber  { background:#FFFBEB; color:#B45309; }
.badge-green  { background:#F0FDF4; color:#15803D; }

/* ── Métricas Streamlit — tarjetas blancas espaciosas ─────────── */
div[data-testid="stMetric"], div[data-testid="metric-container"] {
    background:#fff; border:1px solid #EEF0F4; border-radius:16px;
    padding:18px 20px;
    box-shadow:0 1px 2px rgba(16,24,40,.04), 0 8px 20px -16px rgba(16,24,40,.16);
}
div[data-testid="stMetricLabel"] p { color:#6B7280; font-weight:600; }

/* ── Botones — color vibrante reservado para acción ───────────── */
.stButton > button, div[data-testid="stDownloadButton"] > button {
    border-radius:10px; font-weight:600; padding:.55rem 1.1rem;
    border:1px solid #0F766E; background:#0F766E; color:#fff;
    box-shadow:0 1px 2px rgba(16,24,40,.06); transition:filter .12s ease;
}
.stButton > button:hover, div[data-testid="stDownloadButton"] > button:hover {
    filter:brightness(1.08); border-color:#0F766E; color:#fff;
}

/* ── Tablas — limpias, solo separadores horizontales tenues ───── */
div[data-testid="stDataFrame"] {
    border:1px solid #EEF0F4; border-radius:14px; overflow:hidden;
    box-shadow:0 1px 2px rgba(16,24,40,.04);
}

/* ── Pestañas ─────────────────────────────────────────────────── */
button[data-baseweb="tab"] { font-size:.92rem; font-weight:600; color:#6B7280; }
button[data-baseweb="tab"][aria-selected="true"] { color:#0F766E; }
div[data-baseweb="tab-list"] { border-bottom:1px solid #EEF0F4; gap:4px; }
div[data-baseweb="tab-highlight"] { background:#0F766E; }

/* ── Expanders / inputs ───────────────────────────────────────── */
div[data-testid="stExpander"] {
    border:1px solid #EEF0F4; border-radius:14px; background:#fff;
    box-shadow:0 1px 2px rgba(16,24,40,.04);
}
</style>
""", unsafe_allow_html=True)

# ─── Helpers ─────────────────────────────────────────────────────────────────
def semana_mes(d):
    return min((d.day - 1) // 7 + 1, 4)

# Las 4 funciones derivan del nivel unico (aa_colors.crit_nivel), de modo que
# reconocen tanto la escala de pedidos ('1-CRITICO'...) como la de faltantes
# ('[CRITICO]...'). Antes solo entendian el prefijo 'N-', y por eso la pestana
# Faltantes mostraba todo en verde/OK sin importar la criticidad real.
def crit_class(c):
    return {1: 'crit-1', 2: 'crit-2', 3: 'crit-3', 4: 'crit-3'}.get(crit_nivel(c), 'crit-ok')

def crit_emoji(c):
    return {1: '🔴', 2: '🟠', 3: '🟡', 4: '🟡'}.get(crit_nivel(c), '🟢')

def tarjeta_class(c):
    return {1: 't-crit1', 2: 't-crit2', 3: 't-crit3', 4: 't-crit3'}.get(crit_nivel(c), 't-ok')

def orden_crit(c):
    return crit_nivel(c)

# Estilo unificado de tablas: filas blancas + celda 'Criticidad' como badge
# semantico suave (fondo claro + texto del color del estado). Sustituye al
# pintado de fila completa para una lectura limpia tipo "solo lineas".
_BADGE_CRIT = {
    1: 'background-color:#FEF2F2;color:#B91C1C;font-weight:700',
    2: 'background-color:#FFF7ED;color:#C2410C;font-weight:700',
    3: 'background-color:#FFFBEB;color:#B45309;font-weight:700',
    4: 'background-color:#FFFBEB;color:#B45309;font-weight:700',
}
_BADGE_OK = 'background-color:#F0FDF4;color:#15803D;font-weight:700'

def estilo_tabla(df, col_crit='Criticidad'):
    """Devuelve un Styler con filas blancas y la columna de criticidad
    resaltada como badge semantico suave."""
    sty = df.style.set_properties(**{'background-color': '#ffffff', 'color': '#1F2937'})
    if col_crit in df.columns:
        sty = sty.apply(
            lambda col: [_BADGE_CRIT.get(crit_nivel(str(v)), _BADGE_OK) for v in col],
            subset=[col_crit],
        )
    return sty

# ─── Generador PDF reutilizable ──────────────────────────────────────────────
def build_pdf(titulo, subtitulo, df_data, col_config, orientacion='landscape'):
    """
    Genera un PDF carta con tabla formateada y columna 'Solicitado' en blanco.

    col_config: list de dicts con keys:
        name  → nombre columna en df_data
        label → cabecera visible en PDF
        width → ancho en cm
        align → 'left' | 'center' | 'right'
    orientacion: 'landscape' (apaisado) o 'portrait' (vertical)
    """
    buf = io.BytesIO()
    page = landscape(letter) if orientacion == 'landscape' else letter
    pw, ph = page

    doc = SimpleDocTemplate(
        buf,
        pagesize=page,
        leftMargin=1.5*cm, rightMargin=1.5*cm,
        topMargin=1.5*cm,  bottomMargin=1.5*cm,
    )

    styles = getSampleStyleSheet()
    style_titulo = ParagraphStyle('tit', parent=styles['Heading1'],
                                  fontSize=13, textColor=colors.HexColor('#1A237E'),
                                  spaceAfter=2, alignment=TA_LEFT)
    style_sub    = ParagraphStyle('sub', parent=styles['Normal'],
                                  fontSize=8,  textColor=colors.HexColor('#555555'),
                                  spaceAfter=6, alignment=TA_LEFT)
    style_cell   = ParagraphStyle('cel', parent=styles['Normal'],
                                  fontSize=7.5, leading=9, wordWrap='CJK')
    style_hdr    = ParagraphStyle('hdr', parent=styles['Normal'],
                                  fontSize=8, fontName='Helvetica-Bold',
                                  textColor=colors.white, alignment=TA_CENTER)

    # Columnas de datos + columna SOLICITADO al final
    all_cols    = col_config + [{'name': '_solicitado', 'label': 'Solicitado',
                                  'width': 2.8, 'align': 'center'}]
    col_widths  = [c['width']*cm for c in all_cols]

    # Cabecera
    header_row  = [Paragraph(c['label'], style_hdr) for c in all_cols]

    # Filas 1-CRITICO: texto blanco/negrita para que se lea sobre el rojo fuerte B71C1C
    style_cell_critico = ParagraphStyle('cel_crit', parent=style_cell,
                                         textColor=colors.white, fontName='Helvetica-Bold')

    data_rows = []
    row_crits = []
    for _, row in df_data.iterrows():
        crit = str(row.get('Criticidad', ''))
        base_style = style_cell_critico if crit_nivel(crit) == 1 else style_cell
        cells = []
        for c in col_config:
            v = row.get(c['name'], '')
            v = '' if pd.isna(v) else v
            align = TA_CENTER if c.get('align','left') == 'center' else TA_LEFT
            cells.append(Paragraph(str(v),
                         ParagraphStyle('c', parent=base_style, alignment=align)))
        cells.append('')  # columna Solicitado — vacía
        data_rows.append(cells)
        row_crits.append(crit)

    table_data = [header_row] + data_rows

    # Estilos de tabla
    t_style = [
        # Cabecera
        ('BACKGROUND',  (0,0), (-1,0), colors.HexColor('#1F4E78')),
        ('TEXTCOLOR',   (0,0), (-1,0), colors.white),
        ('FONTNAME',    (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',    (0,0), (-1,0), 8),
        ('ALIGN',       (0,0), (-1,0), 'CENTER'),
        ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
        ('ROWBACKGROUND',(0,0),(-1,0), colors.HexColor('#1F4E78')),
        # Bordes
        ('GRID',        (0,0), (-1,-1), 0.3, colors.HexColor('#CCCCCC')),
        ('LINEBELOW',   (0,0), (-1,0),  0.8, colors.HexColor('#AAAAAA')),
        # Columna Solicitado (última) — borde izquierdo más grueso
        ('LINEAFTER',   (-2,0), (-2,-1), 1.2, colors.HexColor('#999999')),
        # Padding
        ('TOPPADDING',  (0,0), (-1,-1), 3),
        ('BOTTOMPADDING',(0,0),(-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING',(0,0), (-1,-1), 4),
    ]
    # Colores por fila — misma paleta que los Excel (aa_colors.CRIT_FILL_HEX)
    for i, crit in enumerate(row_crits, 1):
        # Etiquetas de pedidos estan en CRIT_FILL_HEX; las de faltantes
        # ('[CRITICO]'...) caen al color por nivel (crit_hex) en vez de gris.
        t_style.append(('BACKGROUND', (0,i), (-1,i),
                         colors.HexColor('#' + (CRIT_FILL_HEX.get(crit) or crit_hex(crit)))))

    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle(t_style))

    # Línea separadora pie
    fecha_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    pie = Paragraph(
        f'Farmacia AT Abierta — Pitrufquen  ·  Generado: {fecha_str}  ·  '
        f'Total: {len(df_data)} registros',
        ParagraphStyle('pie', parent=styles['Normal'], fontSize=7,
                       textColor=colors.HexColor('#888888'), alignment=TA_CENTER)
    )

    story = [
        Paragraph(titulo, style_titulo),
        Paragraph(subtitulo, style_sub),
        HRFlowable(width='100%', thickness=1, color=colors.HexColor('#1F4E78'), spaceAfter=6),
        table,
        Spacer(1, 0.4*cm),
        HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#CCCCCC'), spaceBefore=2),
        pie,
    ]

    doc.build(story)
    buf.seek(0)
    return buf


# ─── Carga de datos ───────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def cargar_datos():
    if not os.path.exists(XLS_MAESTRO):
        return None
    try:
        stock    = pd.read_excel(XLS_MAESTRO, sheet_name='Stock_AA',             engine='openpyxl')
        farm     = pd.read_excel(XLS_MAESTRO, sheet_name='Pedido_Farm_Bodega',   engine='openpyxl')
        bod      = pd.read_excel(XLS_MAESTRO, sheet_name='Pedido_Repos_Bodega',  engine='openpyxl')
        falt     = pd.read_excel(XLS_MAESTRO, sheet_name='Faltas_Farmacia_AA',   engine='openpyxl')
        falt_det = pd.read_excel(XLS_MAESTRO, sheet_name='Faltantes_Detalle_AA', engine='openpyxl')
        # Dialisis (solo recetas de nefrologos) — opcional para compatibilidad
        try:
            dial_farm = pd.read_excel(XLS_MAESTRO, sheet_name='Dialisis_Pedido_Farm', engine='openpyxl')
        except Exception:
            dial_farm = pd.DataFrame()
        try:
            dial_bod = pd.read_excel(XLS_MAESTRO, sheet_name='Dialisis_Pedido_Bod', engine='openpyxl')
        except Exception:
            dial_bod = pd.DataFrame()
        mtime    = datetime.fromtimestamp(os.path.getmtime(XLS_MAESTRO))
        return {'stock': stock, 'farm': farm, 'bod': bod,
                'falt': falt, 'falt_det': falt_det,
                'dial_farm': dial_farm, 'dial_bod': dial_bod, 'mtime': mtime}
    except Exception as e:
        st.error(f"Error cargando datos: {e}")
        return None

def _get(row, col, default=0):
    v = row.get(col, default)
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return default
    return v

# ─── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 💊 Pedidos AA")
    st.markdown("**Farmacia AT Abierta**")
    st.markdown("---")

    if st.button("🔄 Recargar datos"):
        st.cache_data.clear()
        st.rerun()

    datos = cargar_datos()
    if datos:
        horas_ant = (datetime.now() - datos['mtime']).total_seconds() / 3600
        if horas_ant < 4:
            st.success(f"✅ Datos al día\n{datos['mtime'].strftime('%d/%m/%Y %H:%M')}")
        elif horas_ant < 24:
            st.warning(f"⚠️ Datos de hace {int(horas_ant)}h\n{datos['mtime'].strftime('%d/%m/%Y %H:%M')}")
        else:
            dias = int(horas_ant // 24)
            st.error(f"🔴 Datos de hace **{dias} día{'s' if dias>1 else ''}**\n"
                     f"{datos['mtime'].strftime('%d/%m/%Y %H:%M')}\n"
                     "Ejecuta AUTO_SSASUR.bat para actualizar.")
    else:
        st.error("No se encontró el archivo maestro.\nEjecuta EJECUTAR_MAESTRO.bat primero.")
        st.stop()

    hoy    = datetime.now()
    semana = semana_mes(hoy)
    st.info(f"📅 {hoy.strftime('%d/%m/%Y')}  ·  Semana **S{semana}**")

    st.markdown("---")
    n_ped = len(st.session_state.get('pedido', {}))
    if n_ped:
        st.markdown(f"🛒 **{n_ped} medicamento{'s' if n_ped>1 else ''} en pedido**")
    else:
        st.caption("Tu pedido aparece en la pestaña 🔍 Buscar")

# ─── Datos ───────────────────────────────────────────────────────────────────
df_stock    = datos['stock']
df_farm     = datos['farm']
df_bod      = datos['bod']
df_falt     = datos['falt']
df_falt_det = datos['falt_det']
df_dial_farm = datos.get('dial_farm', pd.DataFrame())
df_dial_bod  = datos.get('dial_bod',  pd.DataFrame())
todos_meds = sorted(df_stock['Medicamento'].dropna().unique().tolist())

if 'pedido' not in st.session_state:
    st.session_state.pedido = {}

# ═══════════════════════════════════════════════════════════════════════
# ENCABEZADO DE MARCA + LEYENDA  (visible sobre todas las pestañas)
# ═══════════════════════════════════════════════════════════════════════
_dia_es = {'Mon':'lun','Tue':'mar','Wed':'mié','Thu':'jue','Fri':'vie','Sat':'sáb','Sun':'dom'}
_mes_es = {1:'ene',2:'feb',3:'mar',4:'abr',5:'may',6:'jun',7:'jul',8:'ago',9:'sep',10:'oct',11:'nov',12:'dic'}
_fecha_chip = f"{_dia_es.get(hoy.strftime('%a'), '')} {hoy.day} {_mes_es.get(hoy.month, '')}"
st.markdown(f"""
<div class='app-hero'>
  <div class='ico'>💊</div>
  <div>
    <p class='tit'>Pedidos AA</p>
    <p class='sub'>Farmacia Atención Abierta · Hospital de Pitrufquén</p>
  </div>
  <div class='chip'>
    <div class='d'>{_fecha_chip}</div>
    <div class='w'>Semana S{semana}</div>
  </div>
</div>
<div class='leyenda'>
  <span class='lg'>🔴 Crítico — sin stock ni respaldo</span>
  <span class='lg'>🟠 Urgente</span>
  <span class='lg'>🟡 Moderado</span>
  <span class='lg'>🟢 Suficiente</span>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════
# PESTAÑAS PRINCIPALES
# ═══════════════════════════════════════════════════════════════════════
tab_farm, tab_bod, tab_pedido_bod, tab_pedido_farm, tab_dialisis, tab_faltantes, tab_foto, tab_buscar, tab_feedback = st.tabs([
    "📋  TOP 20 — Farmacia AA",
    "📦  TOP 20 — Bodega AA",
    "📝  Pedido a Bodega AA",
    "🏭  Pedido a Bodega Fármacos",
    "💉  Diálisis",
    "🚨  Faltantes",
    "📷  Foto Manuscrita",
    "🔍  Buscar y armar pedido",
    "💬  Diagnóstico y Sugerencias",
])

# ══════════════════════════════════════════════════════════════════════
# PESTAÑA 1 — TOP 20 FARMACIA AA
# ══════════════════════════════════════════════════════════════════════
with tab_farm:
    st.markdown("## 📋 Top 20 — Pedido Farmacia AA → Bodega AA")
    st.markdown(
        "Estos son los **20 medicamentos que más urgentemente necesita pedir la Farmacia AA** "
        "a la Bodega AA, según las estadísticas históricas y la semana actual del mes.\n\n"
        "Se priorizan los medicamentos con mayor criticidad (pacientes sin alternativa) "
        "y mayor necesidad proyectada para los próximos **5 días hábiles**."
    )

    # Alertas faltantes críticos
    if 'Stock_AA_Total' in df_falt.columns:
        falt_reales = df_falt[df_falt['Stock_AA_Total'].fillna(0) == 0]
        if len(falt_reales):
            st.error(
                f"⚠️ **{len(falt_reales)} medicamento(s) con stock = 0 en este momento** — "
                "aparecen primero en la lista."
            )

    # Filtrar y ordenar
    df_f = df_farm.copy()
    if 'Necesidad_5D_Farm' in df_f.columns:
        df_f['_nec'] = pd.to_numeric(df_f['Necesidad_5D_Farm'], errors='coerce').fillna(0)
    else:
        df_f['_nec'] = 0
    if 'Criticidad' in df_f.columns:
        df_f['_ord'] = df_f['Criticidad'].apply(orden_crit)
    else:
        df_f['_ord'] = 5

    # Prioridad: (1) faltantes reales, (2) criticidad, (3) necesidad descendente
    meds_faltantes = set(df_falt[df_falt.get('Stock_AA_Total', pd.Series(0, index=df_falt.index)).fillna(0) == 0]['Medicamento'].tolist()) \
        if 'Stock_AA_Total' in df_falt.columns else set()

    df_f['_faltante'] = df_f['Medicamento'].isin(meds_faltantes).astype(int)
    df_f = df_f.sort_values(['_faltante','_ord','_nec'], ascending=[False, True, False])

    # Solo los que tienen necesidad > 0 o son faltantes reales
    df_top = df_f[(df_f['_nec'] > 0) | (df_f['_faltante'] == 1)].head(20).reset_index(drop=True)

    if df_top.empty:
        st.success("✅ No hay medicamentos urgentes en Farmacia AA en este momento.")
    else:
        st.markdown(f"### Mostrando {len(df_top)} medicamentos prioritarios")
        st.markdown("---")

        for i, row in df_top.iterrows():
            med       = str(_get(row, 'Medicamento', ''))
            crit      = str(_get(row, 'Criticidad', '5-OK'))
            nec       = int(float(_get(row, 'Necesidad_5D_Farm', 0)))
            stock_f   = int(float(_get(row, 'Stock_Farm_Actual', 0)))
            cob       = float(_get(row, 'Cob_Farm_Actual_Dias', 0))
            consumo   = float(_get(row, 'Consumo_5D_Trend', 0))
            accion1   = str(_get(row, 'Accion_1_Traspaso_Bodega', '') or '')
            accion2   = str(_get(row, 'Accion_2_Gestion_Externa', '') or '')
            es_falt   = med in meds_faltantes
            tc        = tarjeta_class(crit)
            emo       = crit_emoji(crit)

            # Número de orden
            num = i + 1
            alerta_falt = " 🚨 FALTANTE REAL" if es_falt else ""

            with st.container():
                st.markdown(f"""
                <div class='tarjeta {tc}'>
                  <div style='display:flex;align-items:center;gap:16px;flex-wrap:wrap'>
                    <div style='min-width:36px;text-align:center'>
                      <span style='font-size:1.6rem;font-weight:900;color:#555'>#{num}</span>
                    </div>
                    <div style='flex:1;min-width:200px'>
                      <div style='font-size:1.05rem;font-weight:700'>{emo} {med}{alerta_falt}</div>
                      <div style='font-size:0.85rem;color:#555;margin-top:2px'>
                        Criticidad: <b>{crit}</b>
                      </div>
                    </div>
                    <div style='text-align:center;min-width:90px'>
                      <div class='num-grande'>{nec:,}</div>
                      <div class='etiqueta'>unidades a pedir</div>
                    </div>
                    <div style='text-align:center;min-width:90px'>
                      <div style='font-size:1.2rem;font-weight:700;color:#333'>{stock_f:,}</div>
                      <div class='etiqueta'>stock farmacia</div>
                    </div>
                    <div style='text-align:center;min-width:90px'>
                      <div style='font-size:1.2rem;font-weight:700;color:#1565C0'>{cob:.1f} dias</div>
                      <div class='etiqueta'>cobertura actual</div>
                    </div>
                    <div style='text-align:center;min-width:90px'>
                      <div style='font-size:1.1rem;font-weight:600;color:#555'>{consumo:.0f}</div>
                      <div class='etiqueta'>consumo proyectado</div>
                    </div>
                  </div>
                  {'<div class=accion-box>✅ Accion 1: ' + accion1 + '</div>' if accion1 else ''}
                  {'<div class=accion-ext>⚠️ Accion 2: ' + accion2 + '</div>' if accion2 else ''}
                </div>
                """, unsafe_allow_html=True)

        # Botón agregar todos al pedido
        st.markdown("---")
        if st.button("➕ Agregar los 20 al pedido", type="primary", key="add_top20_farm"):
            agregados = 0
            for _, row in df_top.iterrows():
                med    = str(_get(row, 'Medicamento', ''))
                nec    = int(float(_get(row, 'Necesidad_5D_Farm', 0)))
                crit   = str(_get(row, 'Criticidad', '5-OK'))
                acc1   = str(_get(row, 'Accion_1_Traspaso_Bodega', '') or '')
                acc2   = str(_get(row, 'Accion_2_Gestion_Externa', '') or '')
                if med and nec > 0:
                    st.session_state.pedido[med] = {
                        'sugerido': nec, 'confirmado': nec,
                        'criticidad': crit, 'accion1': acc1,
                        'accion2': acc2, 'notas': 'Top 20 Farm',
                    }
                    agregados += 1
            st.success(f"✅ {agregados} medicamentos agregados al pedido. Ve a la pestaña **Buscar y armar pedido** para revisar y descargar.")

# ══════════════════════════════════════════════════════════════════════
# PESTAÑA 2 — TOP 20 BODEGA AA
# ══════════════════════════════════════════════════════════════════════
with tab_bod:
    st.markdown("## 📦 Top 20 — Reposición Bodega AA → Hospital / Compra Externa")
    st.markdown(
        "Estos son los **20 medicamentos que más urgentemente necesita reponer la Bodega AA**, "
        "considerando la demanda proyectada para el próximo **ciclo de pedido = 2 semanas (10 días háb.)**.\n\n"
        "La Bodega AA solicita a Bodega Fármacos **cada 2 semanas**. "
        "El objetivo es mantener al menos **1 ciclo completo (10 días háb.) de cobertura** "
        "para abastecer sin interrupciones a la Farmacia AA."
    )

    # Filtrar y ordenar
    df_b = df_bod.copy()
    if 'Reponer_Bodega' in df_b.columns:
        df_b['_nec'] = pd.to_numeric(df_b['Reponer_Bodega'], errors='coerce').fillna(0)
    else:
        df_b['_nec'] = 0
    if 'Criticidad' in df_b.columns:
        df_b['_ord'] = df_b['Criticidad'].apply(orden_crit)
    else:
        df_b['_ord'] = 5

    df_b['_faltante'] = df_b['Medicamento'].isin(meds_faltantes).astype(int)
    df_b = df_b.sort_values(['_faltante','_ord','_nec'], ascending=[False, True, False])
    df_top_b = df_b[(df_b['_nec'] > 0) | (df_b['_faltante'] == 1)].head(20).reset_index(drop=True)

    if df_top_b.empty:
        st.success("✅ La Bodega AA tiene autonomia suficiente para todos los medicamentos.")
    else:
        st.markdown(f"### Mostrando {len(df_top_b)} medicamentos a reponer")
        st.markdown("---")

        # Métricas resumen en la cabecera
        total_compra = int(df_top_b.get('Compra_Externa_Bod', pd.Series(0)).fillna(0).sum()) \
            if 'Compra_Externa_Bod' in df_top_b.columns else 0
        total_hosp   = int(df_top_b.get('Traspaso_Hospital_Bod', pd.Series(0)).fillna(0).sum()) \
            if 'Traspaso_Hospital_Bod' in df_top_b.columns else 0
        n_criticos   = int((df_top_b['_ord'] <= 2).sum()) if '_ord' in df_top_b.columns else 0

        m1, m2, m3 = st.columns(3)
        m1.metric("Medicamentos criticos (1-2)", f"{n_criticos}")
        m2.metric("Unidades desde Hospital",     f"{total_hosp:,}")
        m3.metric("Unidades a comprar externo",  f"{total_compra:,}")
        st.markdown("---")

        for i, row in df_top_b.iterrows():
            med      = str(_get(row, 'Medicamento', ''))
            crit     = str(_get(row, 'Criticidad', '5-OK'))
            nec      = int(float(_get(row, 'Reponer_Bodega', 0)))
            stock_b  = int(float(_get(row, 'Stock_Bod_Post_Traspaso', 0)))
            cob_b    = float(_get(row, 'Cob_Bod_Post_Dias', 0))
            consumo  = float(_get(row, 'Consumo_10D_Trend', 0))
            hosp     = int(float(_get(row, 'Traspaso_Hospital_Bod', 0)))
            compra   = int(float(_get(row, 'Compra_Externa_Bod', 0)))
            accion1  = str(_get(row, 'Accion_1_Traspaso_Hospital', '') or '')
            accion2  = str(_get(row, 'Accion_2_Compra_Externa', '') or '')
            es_falt  = med in meds_faltantes
            tc       = tarjeta_class(crit)
            emo      = crit_emoji(crit)
            num      = i + 1
            alerta   = " 🚨 FALTANTE REAL" if es_falt else ""

            # Barra de cobertura visual
            # Meta: 10 dias hab. = ciclo 2 semanas
            cob_pct   = min(int(cob_b / 10 * 100), 100)
            color_bar = ("#E53935" if cob_b < 4    # critico < 4 dias
                         else "#F57C00" if cob_b < 7   # alerta < 1 semana
                         else "#FBC02D" if cob_b < 10  # precaucion < 2 semanas
                         else "#2E7D32")               # ok: cubre ciclo completo

            with st.container():
                st.markdown(f"""
                <div class='tarjeta {tc}'>
                  <div style='display:flex;align-items:center;gap:16px;flex-wrap:wrap'>
                    <div style='min-width:36px;text-align:center'>
                      <span style='font-size:1.6rem;font-weight:900;color:#555'>#{num}</span>
                    </div>
                    <div style='flex:1;min-width:200px'>
                      <div style='font-size:1.05rem;font-weight:700'>{emo} {med}{alerta}</div>
                      <div style='font-size:0.85rem;color:#555;margin-top:2px'>
                        Criticidad: <b>{crit}</b>
                      </div>
                      <div style='margin-top:6px;background:#ddd;border-radius:4px;height:8px;width:100%'>
                        <div style='background:{color_bar};height:8px;border-radius:4px;width:{cob_pct}%'></div>
                      </div>
                      <div style='font-size:0.75rem;color:{color_bar}'><b>{cob_b:.1f} dias</b> cobertura · meta: 10 dias (ciclo 2 semanas)</div>
                    </div>
                    <div style='text-align:center;min-width:90px'>
                      <div class='num-grande' style='color:#1A237E'>{nec:,}</div>
                      <div class='etiqueta'>total a reponer</div>
                    </div>
                    <div style='text-align:center;min-width:90px'>
                      <div style='font-size:1.2rem;font-weight:700;color:#1565C0'>{hosp:,}</div>
                      <div class='etiqueta'>desde hospital</div>
                    </div>
                    <div style='text-align:center;min-width:90px'>
                      <div style='font-size:1.2rem;font-weight:700;color:#880E4F'>{compra:,}</div>
                      <div class='etiqueta'>compra externa</div>
                    </div>
                    <div style='text-align:center;min-width:90px'>
                      <div style='font-size:1.1rem;font-weight:600;color:#555'>{consumo:.0f}</div>
                      <div class='etiqueta'>consumo 10 dias</div>
                    </div>
                  </div>
                  {'<div class=accion-box>✅ Accion 1: ' + accion1 + '</div>' if accion1 else ''}
                  {'<div class=accion-ext>⚠️ Accion 2: ' + accion2 + '</div>' if accion2 else ''}
                </div>
                """, unsafe_allow_html=True)

        # Botón agregar todos al pedido
        st.markdown("---")
        if st.button("➕ Agregar los 20 al pedido", type="primary", key="add_top20_bod"):
            agregados = 0
            for _, row in df_top_b.iterrows():
                med    = str(_get(row, 'Medicamento', ''))
                nec    = int(float(_get(row, 'Reponer_Bodega', 0)))
                crit   = str(_get(row, 'Criticidad', '5-OK'))
                acc1   = str(_get(row, 'Accion_1_Traspaso_Hospital', '') or '')
                acc2   = str(_get(row, 'Accion_2_Compra_Externa', '') or '')
                if med and nec > 0:
                    st.session_state.pedido[med] = {
                        'sugerido': nec, 'confirmado': nec,
                        'criticidad': crit, 'accion1': acc1,
                        'accion2': acc2, 'notas': 'Top 20 Bod',
                    }
                    agregados += 1
            st.success(f"✅ {agregados} medicamentos agregados. Ve a la pestaña **Buscar y armar pedido** para revisar y descargar.")

# ══════════════════════════════════════════════════════════════════════
# PESTAÑA 3 — PEDIDO A BODEGA
# ══════════════════════════════════════════════════════════════════════
with tab_pedido_bod:
    st.markdown("## 📝 Pedido a Bodega AA")
    st.markdown(
        "Lista completa de medicamentos que **Farmacia AA debe solicitar a Bodega AA**, "
        "ordenados por criticidad. Ajusta las cantidades si es necesario y descarga el pedido formal."
    )
    st.markdown("---")

    # Preparar datos del pedido a bodega
    def _col(df, col):
        """Devuelve la columna como Series numérica; cero si no existe."""
        if col in df.columns:
            return pd.to_numeric(df[col], errors='coerce').fillna(0)
        return pd.Series(0, index=df.index, dtype=float)

    df_pb = df_farm.copy()
    df_pb['_nec']  = _col(df_pb, 'Necesidad_5D_Farm')
    df_pb['_tras'] = _col(df_pb, 'Traspaso_Posible')
    df_pb['_ord']  = df_pb['Criticidad'].apply(orden_crit) if 'Criticidad' in df_pb.columns \
                     else pd.Series(5, index=df_pb.index)

    # Solo los que tienen necesidad > 0
    df_pb = df_pb[df_pb['_nec'] > 0].sort_values(
        ['_ord', '_nec'], ascending=[True, False]
    ).reset_index(drop=True)

    if df_pb.empty:
        st.success("✅ No hay medicamentos pendientes de solicitar a Bodega AA en este momento.")
    else:
        # Métricas resumen
        n_crit    = len(df_pb[df_pb['_ord'] <= 1])
        n_urg     = len(df_pb[df_pb['_ord'] <= 2])
        total_ud  = int(df_pb['_nec'].sum())
        con_stock = len(df_pb[df_pb['_tras'] > 0])
        sin_stock = len(df_pb[df_pb['_tras'] == 0])

        mc1, mc2, mc3, mc4 = st.columns(4)
        mc1.metric("Total medicamentos",   f"{len(df_pb)}")
        mc2.metric("Críticos + Urgentes",  f"{n_urg}")
        mc3.metric("Con stock en Bodega",  f"{con_stock}")
        mc4.metric("Sin stock en Bodega",  f"{sin_stock} ⚠️")

        st.markdown("---")

        # ── Tabla editable del pedido ─────────────────────────────────────────
        st.markdown("### Detalle del pedido — edita las cantidades si necesitas ajustar")

        # Construir dataframe para mostrar/editar
        filas_pb = []
        for _, row in df_pb.iterrows():
            nec   = int(float(_get(row, 'Necesidad_5D_Farm', 0)))
            tras  = int(float(_get(row, 'Traspaso_Posible', 0)))
            def_  = int(float(_get(row, 'Deficit_Farm', 0)))
            filas_pb.append({
                'Medicamento'      : str(_get(row, 'Medicamento', '')),
                'Criticidad'       : str(_get(row, 'Criticidad', '5-OK')),
                'Stock Farmacia'   : int(float(_get(row, 'Stock_Farm_Actual', 0))),
                'Cob. actual (dias)': round(float(_get(row, 'Cob_Farm_Actual_Dias', 0)), 1),
                'Disponible Bodega': tras,
                'Solicitar (ud)'   : nec,
                'Deficit externo'  : def_,
                'Accion'           : str(_get(row, 'Accion_1_Traspaso_Bodega', '') or ''),
            })

        df_tabla = pd.DataFrame(filas_pb)

        st.dataframe(
            estilo_tabla(df_tabla),
            use_container_width=True,
            hide_index=True,
            column_config={
                'Medicamento'       : st.column_config.TextColumn("Medicamento",         width="large"),
                'Criticidad'        : st.column_config.TextColumn("Criticidad",          width="small"),
                'Stock Farmacia'    : st.column_config.NumberColumn("Stock Farm.",        format="%d"),
                'Cob. actual (dias)': st.column_config.NumberColumn("Cob. actual",        format="%.1f d"),
                'Disponible Bodega' : st.column_config.NumberColumn("Disponible Bod.",    format="%d"),
                'Solicitar (ud)'    : st.column_config.NumberColumn("A solicitar",        format="%d ud"),
                'Deficit externo'   : st.column_config.NumberColumn("Deficit externo",    format="%d"),
                'Accion'            : st.column_config.TextColumn("Accion sugerida",      width="medium"),
            },
            height=min(50 + len(df_tabla) * 35, 600),
        )

        # ── Alertas importantes ───────────────────────────────────────────────
        sin_bod = df_tabla[df_tabla['Disponible Bodega'] == 0]
        if len(sin_bod):
            with st.expander(f"⚠️ {len(sin_bod)} medicamento(s) SIN STOCK en Bodega — requieren gestión externa", expanded=True):
                for _, r in sin_bod.iterrows():
                    crit = str(r['Criticidad'])
                    emo  = crit_emoji(crit)
                    st.markdown(
                        f"{emo} **{r['Medicamento']}** — "
                        f"Criticidad: `{crit}` — "
                        f"Necesidad: **{r['Solicitar (ud)']} ud.**"
                    )

        # ── Descarga Excel del pedido ─────────────────────────────────────────
        st.markdown("---")
        st.markdown("### Descargar pedido formal")

        def build_pedido_bodega_excel():
            wb = Workbook()
            ws = wb.active
            ws.title = "Pedido a Bodega AA"

            fills_crit = {k: crit_fill(k) for k in CRIT_FILL_HEX}

            # Fila 1 — Título
            ws['A1'] = f'PEDIDO FARMACIA AA → BODEGA AA   |   S{semana}  {hoy.strftime("%d/%m/%Y")}'
            ws['A1'].fill = PatternFill('solid', fgColor='1A237E')
            ws['A1'].font = Font(bold=True, color='FFFFFF', name='Arial', size=13)
            ws.merge_cells('A1:H1')
            ws.row_dimensions[1].height = 28

            # Fila 2 — Info resumen
            ws['A2'] = f'Total: {len(df_tabla)} medicamentos  |  Con stock bodega: {con_stock}  |  Sin stock: {sin_stock}'
            ws['A2'].font = Font(italic=True, name='Arial', size=10, color='444444')
            ws.merge_cells('A2:H2')
            ws.row_dimensions[2].height = 16

            # Fila 3 — Encabezados
            headers = ['N°','Medicamento','Criticidad','Stock Farmacia',
                       'Cob. actual (dias)','Disponible Bodega','A Solicitar (ud)','Accion sugerida']
            hfill = PatternFill('solid', fgColor='1F4E78')
            hfont = Font(bold=True, color='FFFFFF', name='Arial', size=10)
            for ci, h in enumerate(headers, 1):
                c = ws.cell(row=3, column=ci, value=h)
                c.fill = hfill; c.font = hfont
                c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws.row_dimensions[3].height = 30

            # Datos
            for ri, (_, row) in enumerate(df_tabla.iterrows(), 4):
                crit = str(row.get('Criticidad', ''))
                fill = fills_crit.get(crit, PatternFill('solid', fgColor='FFFFFF'))
                is_crit1 = crit == '1-CRITICO'
                vals = [ri - 3,
                        row.get('Medicamento', ''),
                        crit,
                        int(row.get('Stock Farmacia', 0) or 0),
                        round(float(row.get('Cob. actual (dias)', 0) or 0), 1),
                        int(row.get('Disponible Bodega', 0) or 0),
                        int(row.get('Solicitar (ud)', 0) or 0),
                        str(row.get('Accion', '') or '')]
                for ci, v in enumerate(vals, 1):
                    c = ws.cell(row=ri, column=ci, value=v)
                    c.fill = fill
                    c.font = Font(name='Arial', size=10,
                                  bold=is_crit1,
                                  color='FFFFFF' if is_crit1 else '000000')
                    c.alignment = Alignment(vertical='center',
                                            wrap_text=(ci == 2 or ci == 8))

            # Anchos de columna
            for ci, w in enumerate([5, 50, 13, 13, 14, 15, 14, 35], 1):
                ws.column_dimensions[get_column_letter(ci)].width = w
            ws.freeze_panes = 'A4'

            # Sección de firma / autorización al final
            ultima = len(df_tabla) + 5
            ws.cell(row=ultima, column=1,
                    value='Solicitado por:').font = Font(name='Arial', size=10)
            ws.cell(row=ultima, column=4,
                    value='Firma Farmaceutico:').font = Font(name='Arial', size=10)
            ws.cell(row=ultima, column=7,
                    value='Fecha:').font = Font(name='Arial', size=10)

            buf = io.BytesIO()
            wb.save(buf); buf.seek(0)
            return buf

        col_dl1, col_dl2, col_dl3 = st.columns([2, 2, 2])
        with col_dl1:
            fname_bod = f"Pedido_Bodega_AA_{hoy.strftime('%Y%m%d_%H%M')}.xlsx"
            st.download_button(
                label="📥 Excel",
                data=build_pedido_bodega_excel(),
                file_name=fname_bod,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary", use_container_width=True,
            )
        with col_dl2:
            _cols_pdf_bod = [
                {'name':'Medicamento',    'label':'Medicamento',        'width':7.0, 'align':'left'},
                {'name':'Criticidad',     'label':'Criticidad',         'width':2.8, 'align':'center'},
                {'name':'Stock Farmacia', 'label':'Stock Farm.',         'width':2.2, 'align':'center'},
                {'name':'Cob. actual (dias)', 'label':'Cob.(dias)',     'width':1.8, 'align':'center'},
                {'name':'Disponible Bodega','label':'Disp. Bodega',     'width':2.2, 'align':'center'},
                {'name':'Solicitar (ud)', 'label':'A Solicitar',        'width':2.2, 'align':'center'},
                {'name':'Accion',         'label':'Accion',             'width':5.5, 'align':'left'},
            ]
            st.download_button(
                label="📄 PDF Carta",
                data=build_pdf(
                    f"PEDIDO FARMACIA AA → BODEGA AA  |  S{semana} · {hoy.strftime('%d/%m/%Y')}",
                    f"Periodo: {datos['mtime'].strftime('%d/%m/%Y')}  ·  {len(df_tabla)} medicamentos",
                    df_tabla, _cols_pdf_bod, orientacion='landscape'
                ),
                file_name=f"Pedido_Bodega_AA_{hoy.strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        with col_dl3:
            st.caption(
                f"{len(df_tabla)} medicamentos · análisis: {datos['mtime'].strftime('%d/%m/%Y %H:%M')}"
            )

# ══════════════════════════════════════════════════════════════════════
# HELPER NUMÉRICO — usado por las pestañas de abajo
# ══════════════════════════════════════════════════════════════════════
def _num(df, col):
    if col in df.columns:
        return pd.to_numeric(df[col], errors='coerce').fillna(0)
    return pd.Series(0, index=df.index, dtype=float)

# ══════════════════════════════════════════════════════════════════════
# PESTAÑA 4 — PEDIDO A BODEGA FÁRMACOS
# Bodega AA solicita a Bodega Fármacos (única bodega de respaldo)
# ══════════════════════════════════════════════════════════════════════
with tab_pedido_farm:
    st.markdown("## 🏭 Pedido Bodega AA → Bodega Fármacos")
    st.markdown(
        "Medicamentos que **Bodega AA necesita reponer desde Bodega Fármacos** "
        "(única bodega de respaldo), ordenados por criticidad. "
        "Incluye la cantidad disponible en Bodega Fármacos y lo que requiere compra externa."
    )
    st.markdown("---")

    df_pf = df_bod.copy()
    df_pf['_nec']   = _num(df_pf, 'Reponer_Bodega')
    # Stock REAL en Bodega Fármacos (único respaldo válido). Fallback a la
    # columna antigua (suma de todas las bodegas) si el Excel no se regeneró.
    _col_bf = 'Stock_BODEGA_FARMACOS' if 'Stock_BODEGA_FARMACOS' in df_pf.columns else 'Stock_Hospital_Total'
    df_pf['_hosp']  = _num(df_pf, _col_bf)
    df_pf['_ord']   = (df_pf['Criticidad'].apply(orden_crit)
                       if 'Criticidad' in df_pf.columns
                       else pd.Series(5, index=df_pf.index))

    df_pf = df_pf[df_pf['_nec'] > 0].sort_values(
        ['_ord', '_nec'], ascending=[True, False]
    ).reset_index(drop=True)

    if df_pf.empty:
        st.success("✅ Bodega AA no necesita reponer desde Bodega Fármacos en este momento.")
    else:
        # Métricas resumen
        n_con_stk  = len(df_pf[df_pf['_hosp'] > 0])
        n_sin_stk  = len(df_pf[df_pf['_hosp'] == 0])
        total_pedir = int(df_pf['_nec'].sum())
        total_disp  = int(df_pf['_hosp'].sum())

        mp1, mp2, mp3, mp4 = st.columns(4)
        mp1.metric("Total medicamentos",           f"{len(df_pf)}")
        mp2.metric("Con stock en Bod. Fármacos",   f"{n_con_stk} ✅")
        mp3.metric("Sin stock en Bod. Fármacos",   f"{n_sin_stk} ⚠️")
        mp4.metric("Unidades totales a reponer",   f"{total_pedir:,}")
        st.markdown("---")

        # Tabla principal
        filas_pf = []
        for _, row in df_pf.iterrows():
            nec    = int(_get(row, 'Reponer_Bodega', 0))
            hosp   = int(_get(row, _col_bf, 0))
            a_pedir = min(nec, hosp)
            compra  = max(nec - hosp, 0)
            filas_pf.append({
                'Medicamento'            : str(_get(row, 'Medicamento', '')),
                'Criticidad'             : str(_get(row, 'Criticidad', '')),
                'Stock Bodega AA actual' : int(float(_get(row, 'Stock_Bod_Actual', 0))),
                'Cob. actual (dias)'     : round(float(_get(row, 'Cob_Bod_Actual_Dias', 0)), 1),
                'Req. 2 semanas (10d)'   : int(float(_get(row, 'Req_2_Semanas', 0))),
                'Consumo 10D proyectado' : int(float(_get(row, 'Consumo_10D_Trend', 0))),
                'Total a reponer'        : nec,
                'Disponible Bod. Farm.'  : hosp,
                'Solicitar a Bod. Farm.' : a_pedir,
                'Compra externa'         : compra,
                'Accion 1'               : str(_get(row, 'Accion_1_Traspaso_Hospital', '') or ''),
                'Accion 2'               : str(_get(row, 'Accion_2_Compra_Externa', '') or ''),
            })

        df_tabla_pf = pd.DataFrame(filas_pf)

        st.dataframe(
            estilo_tabla(df_tabla_pf),
            use_container_width=True, hide_index=True,
            column_config={
                'Medicamento'           : st.column_config.TextColumn("Medicamento",             width="large"),
                'Criticidad'            : st.column_config.TextColumn("Criticidad",              width="small"),
                'Stock Bodega AA actual': st.column_config.NumberColumn("Stock Bod. AA",         format="%d"),
                'Cob. actual (dias)'    : st.column_config.NumberColumn("Cob. actual",           format="%.1f d"),
                'Req. 2 semanas (10d)'  : st.column_config.NumberColumn("Req. 2 sem. (10d)",     format="%d"),
                'Consumo 10D proyectado': st.column_config.NumberColumn("Consumo 10D tend.",     format="%d"),
                'Total a reponer'       : st.column_config.NumberColumn("Total a reponer",       format="%d ud"),
                'Disponible Bod. Farm.' : st.column_config.NumberColumn("Disp. Bod. Farm.",      format="%d"),
                'Solicitar a Bod. Farm.': st.column_config.NumberColumn("Solicitar",             format="%d ud"),
                'Compra externa'        : st.column_config.NumberColumn("Compra externa",        format="%d"),
                'Accion 1'              : st.column_config.TextColumn("Accion 1",                width="medium"),
                'Accion 2'              : st.column_config.TextColumn("Accion 2",                width="medium"),
            },
            height=min(50 + len(df_tabla_pf) * 35, 620),
        )

        # Alertas: sin stock en Bodega Fármacos
        sin_farm = df_tabla_pf[df_tabla_pf['Disponible Bod. Farm.'] == 0]
        if len(sin_farm):
            with st.expander(f"⚠️ {len(sin_farm)} medicamento(s) sin stock en Bodega Fármacos — requieren compra externa", expanded=False):
                for _, r in sin_farm.iterrows():
                    emo = crit_emoji(str(r['Criticidad']))
                    st.markdown(f"{emo} **{r['Medicamento']}** — Necesidad: **{r['Total a reponer']} ud.** → Compra externa")

        # Descarga Excel
        st.markdown("---")
        def build_pedido_farm_excel():
            wb = Workbook()
            ws = wb.active
            ws.title = "Pedido a Bodega Farmacos"

            ws['A1'] = f'PEDIDO BODEGA AA → BODEGA FARMACOS  (CICLO 2 SEMANAS)  |   S{semana}  {hoy.strftime("%d/%m/%Y")}'
            ws['A1'].fill = PatternFill('solid', fgColor='1B5E20')
            ws['A1'].font = Font(bold=True, color='FFFFFF', name='Arial', size=13)
            ws.merge_cells('A1:J1')
            ws.row_dimensions[1].height = 28

            ws['A2'] = f'Ciclo pedido: cada 2 semanas (10 días háb.)  |  Total: {len(df_tabla_pf)} meds  |  Con stock Bod.Farm.: {n_con_stk}  |  Requieren compra: {n_sin_stk}'
            ws['A2'].font = Font(italic=True, name='Arial', size=10, color='444444')
            ws.merge_cells('A2:J2')

            hdrs = ['N°','Medicamento','Criticidad','Stock Bod.AA','Cob.(días)',
                    'Req. 2 Semanas','Consumo 10D tend.','Disp.Bod.Farm.','Solicitar','Compra Externa']
            hfill = PatternFill('solid', fgColor='2E7D32')
            hfont = Font(bold=True, color='FFFFFF', name='Arial', size=10)
            for ci, h in enumerate(hdrs, 1):
                c = ws.cell(row=3, column=ci, value=h)
                c.fill = hfill; c.font = hfont
                c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws.row_dimensions[3].height = 30

            fills_c = {k: crit_fill(k) for k in CRIT_FILL_HEX}
            for ri, (_, row) in enumerate(df_tabla_pf.iterrows(), 4):
                crit = str(row.get('Criticidad', ''))
                fill = fills_c.get(crit, PatternFill('solid', fgColor='F1F8F1'))
                is_crit1 = crit == '1-CRITICO'
                vals = [ri-3,
                        row.get('Medicamento', ''),
                        crit,
                        int(row.get('Stock Bodega AA actual', 0) or 0),
                        round(float(row.get('Cob. actual (dias)', 0) or 0), 1),
                        int(row.get('Req. 2 semanas (10d)', 0) or 0),
                        int(row.get('Consumo 10D proyectado', 0) or 0),
                        int(row.get('Disponible Bod. Farm.', 0) or 0),
                        int(row.get('Solicitar a Bod. Farm.', 0) or 0),
                        int(row.get('Compra externa', 0) or 0)]
                for ci, v in enumerate(vals, 1):
                    c = ws.cell(row=ri, column=ci, value=v)
                    c.fill = fill
                    c.font = Font(name='Arial', size=10,
                                  bold=is_crit1,
                                  color='FFFFFF' if is_crit1 else '000000')
                    c.alignment = Alignment(vertical='center', wrap_text=(ci==2))

            for ci, w in enumerate([4,48,13,10,12,13,13,14,12,13], 1):
                ws.column_dimensions[get_column_letter(ci)].width = w
            ws.freeze_panes = 'A4'

            # Firma
            ult = len(df_tabla_pf) + 6
            ws.cell(row=ult, column=1, value='Solicitado por (Bodega AA):').font = Font(name='Arial',size=10)
            ws.cell(row=ult, column=5, value='Recibido por (Bod. Fármacos):').font = Font(name='Arial',size=10)
            ws.cell(row=ult, column=10, value=f'Fecha: {hoy.strftime("%d/%m/%Y")}').font = Font(name='Arial',size=10)

            buf = io.BytesIO(); wb.save(buf); buf.seek(0)
            return buf

        cpf1, cpf2, cpf3 = st.columns([2, 2, 2])
        with cpf1:
            st.download_button(
                label="📥 Excel",
                data=build_pedido_farm_excel(),
                file_name=f"Pedido_BodFarmacos_{hoy.strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary", use_container_width=True,
            )
        with cpf2:
            _cols_pdf_farm = [
                {'name':'Medicamento',            'label':'Medicamento',         'width':6.8, 'align':'left'},
                {'name':'Criticidad',             'label':'Criticidad',          'width':2.8, 'align':'center'},
                {'name':'Stock Bodega AA actual', 'label':'Stock Bod.AA',        'width':2.2, 'align':'center'},
                {'name':'Consumo 10D proyectado', 'label':'Consumo 10D',         'width':2.2, 'align':'center'},
                {'name':'Total a reponer',        'label':'A Reponer',           'width':2.2, 'align':'center'},
                {'name':'Disponible Bod. Farm.',  'label':'Disp.Bod.Farm.',      'width':2.4, 'align':'center'},
                {'name':'Solicitar a Bod. Farm.', 'label':'Solicitar',           'width':2.2, 'align':'center'},
                {'name':'Compra externa',         'label':'Compra Externa',      'width':2.2, 'align':'center'},
            ]
            st.download_button(
                label="📄 PDF Carta",
                data=build_pdf(
                    f"PEDIDO BODEGA AA → BODEGA FÁRMACOS  |  S{semana} · {hoy.strftime('%d/%m/%Y')}",
                    f"Ciclo 2 semanas (10 días hábiles)  ·  {len(df_tabla_pf)} medicamentos",
                    df_tabla_pf, _cols_pdf_farm, orientacion='landscape'
                ),
                file_name=f"Pedido_BodFarmacos_{hoy.strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        with cpf3:
            st.caption(f"Datos al: {datos['mtime'].strftime('%d/%m/%Y %H:%M')}.")

# ══════════════════════════════════════════════════════════════════════
# PESTAÑA 5 — DIÁLISIS  (solo recetas de los nefrólogos)
# Mismo motor de cálculo y misma escala de criticidad que las demás pestañas,
# pero el consumo se restringe a las recetas de:
#   Dr. Yasmani Ortiz Amador · Dra. Martha Peralta · Dra. Mónica Amaya
# ══════════════════════════════════════════════════════════════════════
with tab_dialisis:
    st.markdown("## 💉 Pedidos de Diálisis")
    st.markdown(
        "Medicamentos para pacientes en **diálisis / nefrología**, identificados por el "
        "**consumo de las recetas prescritas por** Dr. *Yasmani Ortiz Amador*, "
        "Dra. *Martha Peralta* y Dra. *Mónica Amaya*. "
        "El stock es el compartido de la farmacia.\n\n"
        "🛡️ **Para no quedar corto sobre el stock compartido, la demanda de este pedido suma "
        "el consumo de 5 días de la farmacia (excluyendo diálisis) + el consumo de diálisis** "
        "— es decir, cubre el consumo total real de esos medicamentos. Las columnas *Solo diálisis* "
        "y *Farmacia (no diálisis)* muestran el desglose.\n\n"
        "Se muestran los dos flujos con la **misma categorización** del resto de la app: "
        "qué pedir desde **Farmacia AA → Bodega AA** y desde **Bodega AA → Bodega Fármacos**."
    )

    def _coln(df, col):
        if col in df.columns:
            return pd.to_numeric(df[col], errors='coerce').fillna(0)
        return pd.Series(0, index=df.index, dtype=float)

    sub_farm, sub_bod = st.tabs([
        "📝  Farmacia AA → Bodega AA",
        "🏭  Bodega AA → Bodega Fármacos",
    ])

    # ── 5a) Farmacia AA → Bodega AA (diálisis) ────────────────────────────────
    with sub_farm:
        if df_dial_farm is None or df_dial_farm.empty:
            st.info("No hay pedidos de diálisis Farmacia → Bodega AA en este momento "
                    "(o el archivo maestro es una versión previa: ejecuta EJECUTAR_MAESTRO.bat).")
        else:
            dfa = df_dial_farm.copy()
            dfa['_nec'] = _coln(dfa, 'Necesidad_5D_Farm')
            dfa['_ord'] = dfa['Criticidad'].apply(orden_crit) if 'Criticidad' in dfa.columns \
                          else pd.Series(5, index=dfa.index)
            dfa = dfa[dfa['_nec'] > 0].sort_values(['_ord', '_nec'],
                                                    ascending=[True, False]).reset_index(drop=True)
            if dfa.empty:
                st.success("✅ Farmacia AA no necesita pedir medicamentos de diálisis a Bodega AA.")
            else:
                m1, m2, m3 = st.columns(3)
                m1.metric("Medicamentos a pedir", f"{len(dfa)}")
                m2.metric("Críticos + Urgentes",  f"{len(dfa[dfa['_ord'] <= 2])}")
                m3.metric("Unidades totales",     f"{int(dfa['_nec'].sum()):,}")
                st.markdown("---")

                filas = []
                for _, row in dfa.iterrows():
                    _cons5d_dial = float(_get(row, 'Consumo_5D_Solo_Dialisis', 0))
                    filas.append({
                        'Medicamento'       : str(_get(row, 'Medicamento', '')),
                        'Criticidad'        : str(_get(row, 'Criticidad', '5-OK')),
                        'Stock Farmacia'    : int(float(_get(row, 'Stock_Farm_Actual', 0))),
                        'Cob. actual (dias)': round(float(_get(row, 'Cob_Farm_Actual_Dias', 0)), 1),
                        'Cons. mens. diál.' : int(round(_cons5d_dial / 5 * 30)),
                        'Solo dialisis'     : int(round(_cons5d_dial)),
                        'Farmacia no dial'  : int(round(float(_get(row, 'Consumo_5D_Farm_NoDial', 0)))),
                        'Disponible Bodega' : int(float(_get(row, 'A_Traspasar', 0))),
                        'Solicitar (ud)'    : int(float(_get(row, 'Necesidad_5D_Farm', 0))),
                        'Deficit externo'   : int(float(_get(row, 'Deficit_Post_Traspaso', 0))),
                        'Accion 1'          : str(_get(row, 'Accion_1_Traspaso_Bodega', '') or ''),
                        'Accion 2'          : str(_get(row, 'Accion_2_Gestion_Externa', '') or ''),
                    })
                df_tabla_da = pd.DataFrame(filas)
                st.dataframe(
                    estilo_tabla(df_tabla_da),
                    use_container_width=True, hide_index=True,
                    column_config={
                        'Medicamento'       : st.column_config.TextColumn("Medicamento",           width="large"),
                        'Criticidad'        : st.column_config.TextColumn("Criticidad",            width="small"),
                        'Stock Farmacia'    : st.column_config.NumberColumn("Stock Farm.",          format="%d"),
                        'Cob. actual (dias)': st.column_config.NumberColumn("Cob. actual",          format="%.1f d"),
                        'Cons. mens. diál.' : st.column_config.NumberColumn("Cons. mensual diál.", format="%d ud",
                                              help="Consumo mensual estimado de diálisis (consumo 5D × 6)"),
                        'Solo dialisis'     : st.column_config.NumberColumn("Solo diálisis 5D",    format="%d",
                                              help="Consumo de 5 días sólo de las recetas de diálisis"),
                        'Farmacia no dial'  : st.column_config.NumberColumn("Farmacia (no diál.) 5D", format="%d",
                                              help="Consumo de 5 días de la farmacia general, excluyendo diálisis"),
                        'Disponible Bodega' : st.column_config.NumberColumn("Disponible Bod.",     format="%d"),
                        'Solicitar (ud)'    : st.column_config.NumberColumn("A solicitar",          format="%d ud"),
                        'Deficit externo'   : st.column_config.NumberColumn("Deficit externo",     format="%d"),
                        'Accion 1'          : st.column_config.TextColumn("Acción 1 (traspaso)",   width="medium"),
                        'Accion 2'          : st.column_config.TextColumn("Acción 2 (gestión externa)", width="medium"),
                    },
                    height=min(50 + len(df_tabla_da) * 35, 600),
                )

                def build_dial_farm_excel():
                    wb = Workbook(); ws = wb.active; ws.title = "Dialisis Farmacia AA"
                    fills_c = {k: crit_fill(k) for k in CRIT_FILL_HEX}
                    ws['A1'] = f'DIÁLISIS · PEDIDO FARMACIA AA → BODEGA AA   |   S{semana}  {hoy.strftime("%d/%m/%Y")}'
                    ws['A1'].fill = PatternFill('solid', fgColor='0F766E')
                    ws['A1'].font = Font(bold=True, color='FFFFFF', name='Arial', size=13)
                    ws.merge_cells('A1:H1'); ws.row_dimensions[1].height = 28
                    ws['A2'] = 'Consumo según recetas de: Dr. Yasmani Ortiz Amador · Dra. Martha Peralta · Dra. Mónica Amaya'
                    ws['A2'].font = Font(italic=True, name='Arial', size=10, color='444444')
                    ws.merge_cells('A2:H2')
                    hdrs = ['N°','Medicamento','Criticidad','Stock Farmacia','Cob.(días)',
                            'Disponible Bodega','Solicitar','Déficit externo']
                    hfill = PatternFill('solid', fgColor='0F766E')
                    hfont = Font(bold=True, color='FFFFFF', name='Arial', size=10)
                    for ci, h in enumerate(hdrs, 1):
                        c = ws.cell(row=3, column=ci, value=h)
                        c.fill = hfill; c.font = hfont
                        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    ws.row_dimensions[3].height = 30
                    for ri, (_, r) in enumerate(df_tabla_da.iterrows(), 4):
                        crit = str(r.get('Criticidad', ''))
                        fill = fills_c.get(crit, PatternFill('solid', fgColor='F1F8F1'))
                        is_c1 = crit == '1-CRITICO'
                        vals = [ri-3, r['Medicamento'], crit,
                                int(r['Stock Farmacia']), round(float(r['Cob. actual (dias)']), 1),
                                int(r['Disponible Bodega']), int(r['Solicitar (ud)']), int(r['Deficit externo'])]
                        for ci, v in enumerate(vals, 1):
                            c = ws.cell(row=ri, column=ci, value=v)
                            c.fill = fill
                            c.font = Font(name='Arial', size=10, bold=is_c1,
                                          color='FFFFFF' if is_c1 else '000000')
                            c.alignment = Alignment(vertical='center', wrap_text=(ci == 2))
                    for ci, w in enumerate([4,48,13,13,11,15,11,13], 1):
                        ws.column_dimensions[get_column_letter(ci)].width = w
                    ws.freeze_panes = 'A4'
                    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

                st.markdown("---")
                cda1, cda2 = st.columns([2, 4])
                with cda1:
                    st.download_button(
                        "📥 Excel", data=build_dial_farm_excel(),
                        file_name=f"Dialisis_Farmacia_{hoy.strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary", use_container_width=True,
                    )
                with cda2:
                    st.caption(f"Datos al: {datos['mtime'].strftime('%d/%m/%Y %H:%M')}.")

    # ── 5b) Bodega AA → Bodega Fármacos (diálisis) ────────────────────────────
    with sub_bod:
        if df_dial_bod is None or df_dial_bod.empty:
            st.info("No hay pedidos de diálisis Bodega AA → Bodega Fármacos en este momento "
                    "(o el archivo maestro es una versión previa: ejecuta EJECUTAR_MAESTRO.bat).")
        else:
            dfb = df_dial_bod.copy()
            dfb['_nec']  = _coln(dfb, 'Reponer_Bodega')
            # Stock REAL en Bodega Fármacos (único respaldo válido). Fallback a la
            # columna antigua (suma de todas las bodegas) si el Excel no se regeneró.
            _col_bf_d = 'Stock_BODEGA_FARMACOS' if 'Stock_BODEGA_FARMACOS' in dfb.columns else 'Stock_Hospital_Total'
            dfb['_hosp'] = _coln(dfb, _col_bf_d)
            dfb['_ord']  = dfb['Criticidad'].apply(orden_crit) if 'Criticidad' in dfb.columns \
                           else pd.Series(5, index=dfb.index)
            dfb = dfb[dfb['_nec'] > 0].sort_values(['_ord', '_nec'],
                                                   ascending=[True, False]).reset_index(drop=True)
            if dfb.empty:
                st.success("✅ Bodega AA no necesita reponer medicamentos de diálisis desde Bodega Fármacos.")
            else:
                n_con = len(dfb[dfb['_hosp'] > 0]); n_sin = len(dfb[dfb['_hosp'] == 0])
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Total medicamentos",        f"{len(dfb)}")
                m2.metric("Con stock Bod. Fármacos",    f"{n_con} ✅")
                m3.metric("Sin stock Bod. Fármacos",    f"{n_sin} ⚠️")
                m4.metric("Unidades a reponer",         f"{int(dfb['_nec'].sum()):,}")
                st.markdown("---")

                filas = []
                for _, row in dfb.iterrows():
                    nec  = int(float(_get(row, 'Reponer_Bodega', 0)))
                    hosp = int(float(_get(row, _col_bf_d, 0)))
                    filas.append({
                        'Medicamento'           : str(_get(row, 'Medicamento', '')),
                        'Criticidad'            : str(_get(row, 'Criticidad', '')),
                        'Stock Bodega AA actual': int(float(_get(row, 'Stock_Bod_Actual', 0))),
                        'Cob. actual (dias)'    : round(float(_get(row, 'Cob_Bod_Actual_Dias', 0)), 1),
                        'Solo dialisis'         : int(round(float(_get(row, 'Consumo_5D_Solo_Dialisis', 0)))),
                        'Farmacia no dial'      : int(round(float(_get(row, 'Consumo_5D_Farm_NoDial', 0)))),
                        'Consumo 10D proyectado': int(float(_get(row, 'Consumo_10D_Trend', 0))),
                        'Total a reponer'       : nec,
                        'Disponible Bod. Farm.' : hosp,
                        'Solicitar a Bod. Farm.': min(nec, hosp),
                        'Compra externa'        : max(nec - hosp, 0),
                        'Accion 1'              : str(_get(row, 'Accion_1_Traspaso_Hospital', '') or ''),
                        'Accion 2'              : str(_get(row, 'Accion_2_Compra_Externa', '') or ''),
                    })
                df_tabla_db = pd.DataFrame(filas)
                st.dataframe(
                    estilo_tabla(df_tabla_db),
                    use_container_width=True, hide_index=True,
                    column_config={
                        'Medicamento'           : st.column_config.TextColumn("Medicamento",       width="large"),
                        'Criticidad'            : st.column_config.TextColumn("Criticidad",        width="small"),
                        'Stock Bodega AA actual': st.column_config.NumberColumn("Stock Bod. AA",    format="%d"),
                        'Cob. actual (dias)'    : st.column_config.NumberColumn("Cob. actual",      format="%.1f d"),
                        'Solo dialisis'         : st.column_config.NumberColumn("Solo diálisis 5D",  format="%d", help="Consumo de 5 días sólo de las recetas de diálisis"),
                        'Farmacia no dial'      : st.column_config.NumberColumn("Farmacia (no diál.) 5D", format="%d", help="Consumo de 5 días de la farmacia general, excluyendo diálisis"),
                        'Consumo 10D proyectado': st.column_config.NumberColumn("Consumo 10D tend.", format="%d"),
                        'Total a reponer'       : st.column_config.NumberColumn("Total a reponer",  format="%d ud"),
                        'Disponible Bod. Farm.' : st.column_config.NumberColumn("Disp. Bod. Farm.", format="%d"),
                        'Solicitar a Bod. Farm.': st.column_config.NumberColumn("Solicitar",        format="%d ud"),
                        'Compra externa'        : st.column_config.NumberColumn("Compra externa",   format="%d"),
                        'Accion 1'              : st.column_config.TextColumn("Acción 1",           width="medium"),
                        'Accion 2'              : st.column_config.TextColumn("Acción 2",           width="medium"),
                    },
                    height=min(50 + len(df_tabla_db) * 35, 620),
                )

                def build_dial_bod_excel():
                    wb = Workbook(); ws = wb.active; ws.title = "Dialisis Bodega Farmacos"
                    fills_c = {k: crit_fill(k) for k in CRIT_FILL_HEX}
                    ws['A1'] = f'DIÁLISIS · PEDIDO BODEGA AA → BODEGA FÁRMACOS   |   S{semana}  {hoy.strftime("%d/%m/%Y")}'
                    ws['A1'].fill = PatternFill('solid', fgColor='0E7490')
                    ws['A1'].font = Font(bold=True, color='FFFFFF', name='Arial', size=13)
                    ws.merge_cells('A1:J1'); ws.row_dimensions[1].height = 28
                    ws['A2'] = 'Consumo según recetas de: Dr. Yasmani Ortiz Amador · Dra. Martha Peralta · Dra. Mónica Amaya'
                    ws['A2'].font = Font(italic=True, name='Arial', size=10, color='444444')
                    ws.merge_cells('A2:J2')
                    hdrs = ['N°','Medicamento','Criticidad','Stock Bod.AA','Cob.(días)',
                            'Consumo 10D tend.','Disp.Bod.Farm.','Solicitar','Compra Externa']
                    hfill = PatternFill('solid', fgColor='0E7490')
                    hfont = Font(bold=True, color='FFFFFF', name='Arial', size=10)
                    for ci, h in enumerate(hdrs, 1):
                        c = ws.cell(row=3, column=ci, value=h)
                        c.fill = hfill; c.font = hfont
                        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    ws.row_dimensions[3].height = 30
                    for ri, (_, r) in enumerate(df_tabla_db.iterrows(), 4):
                        crit = str(r.get('Criticidad', ''))
                        fill = fills_c.get(crit, PatternFill('solid', fgColor='F1F8F1'))
                        is_c1 = crit == '1-CRITICO'
                        vals = [ri-3, r['Medicamento'], crit,
                                int(r['Stock Bodega AA actual']), round(float(r['Cob. actual (dias)']), 1),
                                int(r['Consumo 10D proyectado']), int(r['Disponible Bod. Farm.']),
                                int(r['Solicitar a Bod. Farm.']), int(r['Compra externa'])]
                        for ci, v in enumerate(vals, 1):
                            c = ws.cell(row=ri, column=ci, value=v)
                            c.fill = fill
                            c.font = Font(name='Arial', size=10, bold=is_c1,
                                          color='FFFFFF' if is_c1 else '000000')
                            c.alignment = Alignment(vertical='center', wrap_text=(ci == 2))
                    for ci, w in enumerate([4,48,13,12,11,15,14,12,13], 1):
                        ws.column_dimensions[get_column_letter(ci)].width = w
                    ws.freeze_panes = 'A4'
                    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

                st.markdown("---")
                cdb1, cdb2 = st.columns([2, 4])
                with cdb1:
                    st.download_button(
                        "📥 Excel", data=build_dial_bod_excel(),
                        file_name=f"Dialisis_BodFarmacos_{hoy.strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary", use_container_width=True,
                    )
                with cdb2:
                    st.caption(f"Datos al: {datos['mtime'].strftime('%d/%m/%Y %H:%M')}.")

# ══════════════════════════════════════════════════════════════════════
# PESTAÑA 6 — FALTANTES
# Solo usa Bodega Fármacos como respaldo (única bodega de respaldo real)
# ══════════════════════════════════════════════════════════════════════
with tab_faltantes:
    st.markdown("## 🚨 Faltantes Farmacia AT Abierta")
    st.markdown(
        "Medicamentos prescritos con procedencia **AT Abierta**, sin stock en Farmacia AA ni en Bodega AA. "
        "Se separan según si **Bodega Fármacos** (única bodega de respaldo) tiene o no existencias."
    )
    st.markdown("---")

    # Usar Faltantes_Detalle_AA que tiene Stock_BODEGA_FARMACOS
    df_fd = df_falt_det.copy()
    df_fd['_stk_aa']   = _num(df_fd, 'Stock_AA_Total')
    df_fd['_stk_bf']   = _num(df_fd, 'Stock_BODEGA_FARMACOS')   # único respaldo real
    df_fd['_faltante'] = _num(df_fd, 'Faltante_Neto')
    df_fd['_pac']      = _num(df_fd, 'Pacientes_Afectados')
    df_fd['_rec']      = _num(df_fd, 'N_Recetas')
    df_fd['_ord']      = (df_fd['Criticidad'].apply(orden_crit)
                          if 'Criticidad' in df_fd.columns
                          else pd.Series(5, index=df_fd.index))

    # Solo los que tienen stock 0 en AA
    df_fd = df_fd[df_fd['_stk_aa'] <= 0].copy()

    # Separar: sin respaldo (Bod. Fármacos = 0) vs con respaldo (Bod. Fármacos > 0)
    df_sin = df_fd[df_fd['_stk_bf'] <= 0].sort_values(['_ord','_faltante'], ascending=[True,False])
    df_con = df_fd[df_fd['_stk_bf'] > 0].sort_values(['_ord','_faltante'], ascending=[True,False])

    n_sin  = len(df_sin)
    n_con  = len(df_con)
    pac_tot = int(df_fd['_pac'].sum())
    ud_tot  = int(df_fd['_faltante'].sum())

    mf1, mf2, mf3, mf4 = st.columns(4)
    mf1.metric("Sin respaldo — Compra urgente",      f"{n_sin} 🔴")
    mf2.metric("Con respaldo en Bod. Fármacos",      f"{n_con} 🟡")
    mf3.metric("Pacientes afectados",                f"{pac_tot}")
    mf4.metric("Unidades totales faltantes",         f"{ud_tot:,}")

    # ── BLOQUE 1: SIN RESPALDO (Bodega Fármacos = 0) ──────────────────────────
    st.markdown("---")
    st.markdown(f"### 🔴 Sin respaldo — {n_sin} medicamento(s) — Compra Urgente")
    st.caption("Stock AA = 0 y Bodega Fármacos = 0. Requieren compra externa inmediata.")

    if df_sin.empty:
        st.success("✅ No hay faltantes sin respaldo ahora.")
    else:
        for _, row in df_sin.iterrows():
            med  = str(row.get('Medicamento',''))
            crit = str(row.get('Criticidad',''))
            falt = int(row['_faltante'])
            pac  = int(row['_pac'])
            rec  = int(row['_rec'])
            acc  = str(row.get('Accion_Sugerida','') or '')
            emo  = crit_emoji(crit)
            tc   = tarjeta_class(crit)
            st.markdown(f"""
            <div class='tarjeta {tc}'>
              <div style='display:flex;align-items:center;gap:20px;flex-wrap:wrap'>
                <div style='flex:1;min-width:220px'>
                  <div style='font-size:1.05rem;font-weight:700'>{emo} {med}</div>
                  <div style='font-size:0.82rem;color:#666;margin-top:2px'>Criticidad: <b>{crit}</b></div>
                </div>
                <div style='text-align:center;min-width:85px'>
                  <div class='num-grande' style='color:#E53935'>{falt:,}</div>
                  <div class='etiqueta'>unidades faltantes</div>
                </div>
                <div style='text-align:center;min-width:85px'>
                  <div style='font-size:1.3rem;font-weight:700'>{pac}</div>
                  <div class='etiqueta'>pacientes afectados</div>
                </div>
                <div style='text-align:center;min-width:75px'>
                  <div style='font-size:1.1rem;font-weight:600;color:#555'>{rec}</div>
                  <div class='etiqueta'>recetas pendientes</div>
                </div>
              </div>
              {'<div class=accion-ext>🚨 ' + acc + '</div>' if acc else '<div class=accion-ext>🚨 COMPRA URGENTE</div>'}
            </div>
            """, unsafe_allow_html=True)

    # ── BLOQUE 2: CON RESPALDO EN BODEGA FÁRMACOS ─────────────────────────────
    st.markdown("---")
    st.markdown(f"### 🟡 Con respaldo en Bodega Fármacos — {n_con} medicamento(s)")
    st.caption("Stock AA = 0 pero Bodega Fármacos tiene existencias. Gestionar traspaso urgente.")

    if df_con.empty:
        st.info("No hay faltantes con respaldo en Bodega Fármacos.")
    else:
        filas_con = []
        for _, row in df_con.iterrows():
            filas_con.append({
                'Medicamento'          : str(row.get('Medicamento','')),
                'Criticidad'           : str(row.get('Criticidad','')),
                'Unidades faltantes'   : int(row['_faltante']),
                'Stock Bod. Fármacos'  : int(row['_stk_bf']),
                'Pacientes afectados'  : int(row['_pac']),
                'Recetas pendientes'   : int(row['_rec']),
                'Accion'               : str(row.get('Accion_Sugerida','') or 'TRASPASAR DESDE BODEGA FARMACOS'),
            })
        st.dataframe(
            estilo_tabla(pd.DataFrame(filas_con)),
            use_container_width=True, hide_index=True,
            column_config={
                'Medicamento'         : st.column_config.TextColumn("Medicamento",          width="large"),
                'Criticidad'          : st.column_config.TextColumn("Criticidad",           width="small"),
                'Unidades faltantes'  : st.column_config.NumberColumn("Faltantes",          format="%d ud"),
                'Stock Bod. Fármacos' : st.column_config.NumberColumn("Stock Bod. Farm.",   format="%d"),
                'Pacientes afectados' : st.column_config.NumberColumn("Pacientes",          format="%d"),
                'Recetas pendientes'  : st.column_config.NumberColumn("Recetas",            format="%d"),
                'Accion'              : st.column_config.TextColumn("Accion",               width="medium"),
            },
        )

    # ── Descarga Excel ─────────────────────────────────────────────────────────
    st.markdown("---")

    def build_faltantes_excel():
        wb   = Workbook()
        fh   = Font(bold=True, color='FFFFFF', name='Arial', size=11)
        fc   = {k: crit_fill(k) for k in CRIT_FILL_HEX}

        def _hoja(ws, titulo, color_tit, color_h, df_rows, hdrs, extractor):
            ws['A1'] = titulo
            ws['A1'].fill = PatternFill('solid', fgColor=color_tit)
            ws['A1'].font = Font(bold=True, color='FFFFFF', name='Arial', size=13)
            ws.merge_cells(f'A1:{get_column_letter(len(hdrs))}1')
            ws.row_dimensions[1].height = 26
            for ci, h in enumerate(hdrs, 1):
                c = ws.cell(row=2, column=ci, value=h)
                c.fill = PatternFill('solid', fgColor=color_h)
                c.font = fh
                c.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[2].height = 22
            for ri, (_, row) in enumerate(df_rows.iterrows(), 3):
                crit = str(row.get('Criticidad', ''))
                fill = fc.get(crit, PatternFill('solid', fgColor='FFF3F3'))
                vals = extractor(row)
                for ci, v in enumerate(vals, 1):
                    c = ws.cell(row=ri, column=ci, value=v)
                    c.fill = fill
                    c.font = Font(name='Arial', size=11)

        # Hoja 1 — Sin respaldo
        ws1 = wb.active
        ws1.title = "SIN RESPALDO - URGENTE"
        _hoja(ws1,
              f'FALTANTES SIN RESPALDO — {hoy.strftime("%d/%m/%Y %H:%M")}',
              'E53935', 'C62828', df_sin,
              ['Medicamento','Criticidad','Unidades Faltantes','Pacientes','Recetas','Accion'],
              lambda r: [r.get('Medicamento',''),
                         str(r.get('Criticidad','')),
                         int(r.get('_faltante',0) or 0),
                         int(r.get('_pac',0) or 0),
                         int(r.get('_rec',0) or 0),
                         str(r.get('Accion_Sugerida','') or 'COMPRA URGENTE')])
        for ci,w in enumerate([50,15,18,12,10,25],1):
            ws1.column_dimensions[get_column_letter(ci)].width = w
        ws1.freeze_panes = 'A3'

        # Hoja 2 — Con respaldo Bod. Fármacos
        ws2 = wb.create_sheet("CON RESPALDO BOD.FARMACOS")
        _hoja(ws2,
              f'FALTANTES CON RESPALDO BODEGA FARMACOS — {hoy.strftime("%d/%m/%Y %H:%M")}',
              'E65100', 'BF360C', df_con,
              ['Medicamento','Criticidad','Unidades Faltantes','Stock Bod. Farmacos','Pacientes','Accion'],
              lambda r: [r.get('Medicamento',''),
                         str(r.get('Criticidad','')),
                         int(r.get('_faltante',0) or 0),
                         int(r.get('_stk_bf',0) or 0),
                         int(r.get('_pac',0) or 0),
                         str(r.get('Accion_Sugerida','') or 'TRASPASAR DESDE BODEGA FARMACOS')])
        for ci,w in enumerate([50,15,18,16,12,30],1):
            ws2.column_dimensions[get_column_letter(ci)].width = w
        ws2.freeze_panes = 'A3'

        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        return buf

    cff1, cff2, cff3 = st.columns([2, 2, 2])
    with cff1:
        st.download_button(
            label="📥 Excel",
            data=build_faltantes_excel(),
            file_name=f"Faltantes_AA_{hoy.strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True,
        )
    with cff2:
        # PDF faltantes — une sin respaldo y con respaldo
        _df_falt_pdf = pd.concat([df_sin, df_con], ignore_index=True)
        _cols_pdf_falt = [
            {'name':'Medicamento',        'label':'Medicamento',       'width':8.0, 'align':'left'},
            {'name':'Criticidad',         'label':'Criticidad',        'width':3.2, 'align':'center'},
            {'name':'_faltante',          'label':'Unid. Faltantes',   'width':2.4, 'align':'center'},
            {'name':'Stock_BODEGA_FARMACOS','label':'Stock Bod.Farm.', 'width':2.4, 'align':'center'},
            {'name':'_pac',               'label':'Pacientes',         'width':2.0, 'align':'center'},
            {'name':'Accion_Sugerida',    'label':'Accion Sugerida',   'width':5.0, 'align':'left'},
        ]
        st.download_button(
            label="📄 PDF Carta",
            data=build_pdf(
                f"FALTANTES FARMACIA AT ABIERTA  |  {hoy.strftime('%d/%m/%Y')}",
                f"Sin respaldo: {n_sin}  ·  Con respaldo Bod. Fármacos: {n_con}  ·  Total: {len(_df_falt_pdf)}",
                _df_falt_pdf, _cols_pdf_falt, orientacion='landscape'
            ),
            file_name=f"Faltantes_AA_{hoy.strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    with cff3:
        st.caption(
            f"Sin Respaldo: {n_sin} · Con Respaldo Bod. Farm.: {n_con} · "
            f"Datos al {datos['mtime'].strftime('%d/%m/%Y %H:%M')}."
        )

# ══════════════════════════════════════════════════════════════════════
# PESTAÑA 6 — HOJA MANUSCRITA
# Foto O texto → homologación automática → sugerido de cantidades
# ══════════════════════════════════════════════════════════════════════
with tab_foto:
    import unicodedata
    from rapidfuzz import fuzz, process as rfprocess

    st.markdown("## 📋 Hoja de Pedido → Homologación y Sugerido")
    st.markdown(
        "Sube **una foto** de tu hoja manuscrita o pega el texto directamente. "
        "La app homologa cada medicamento con el sistema y te dice cuánto pedir."
    )
    st.markdown("---")

    # ── Helpers (usados en ambos modos) ──────────────────────────────────────
    def _norm(s):
        s = unicodedata.normalize('NFD', str(s).upper())
        return ''.join(c for c in s if not unicodedata.combining(c)).strip()

    @st.cache_data
    def _idx_uni(meds):
        return {_norm(m): m for m in meds}
    idx_uni = _idx_uni(tuple(todos_meds))

    # ── Tipo de pedido ────────────────────────────────────────────────────────
    tipo_ped = st.radio(
        "Tipo de pedido",
        ["Farmacia AA → Bodega AA  (5 días)",
         "Bodega AA → Bodega Fármacos  (10 días)"],
        horizontal=True, key="tipo_ped_hoja"
    )
    es_fh    = "Farmacia" in tipo_ped
    df_ped_h = df_farm if es_fh else df_bod
    col_sug  = 'Necesidad_5D_Farm' if es_fh else 'Reponer_Bodega'

    st.markdown("---")

    # ── Modo de entrada: FOTO  o  TEXTO ──────────────────────────────────────
    modo_entrada = st.radio(
        "¿Cómo ingresarás los medicamentos?",
        ["📷  Foto de la hoja", "✏️  Escribir / Pegar texto"],
        horizontal=True, key="modo_entrada"
    )
    st.markdown("")

    texto_para_procesar = ""  # se llena según el modo

    # ══ MODO FOTO ══════════════════════════════════════════════════════════════
    if "Foto" in modo_entrada:

        # API key — guardada en archivo oculto, invisible si ya existe
        CFG_KEY = os.path.join(WORK_DIR, ".apikey")
        if 'ak' not in st.session_state:
            try:
                with open(CFG_KEY, encoding='utf-8') as _f:
                    st.session_state.ak = _f.read().strip()
            except OSError:
                st.session_state.ak = ''

        if not st.session_state.ak:
            st.info("Para leer fotos manuscritas necesitas una clave API de Claude (gratis en console.anthropic.com).")
            ak_inp = st.text_input("Pega tu API Key aquí", type="password", key="ak_input",
                                   placeholder="sk-ant-api03-...")
            if st.button("Guardar clave", key="save_ak"):
                if ak_inp.startswith("sk-"):
                    st.session_state.ak = ak_inp
                    with open(CFG_KEY, 'w', encoding='utf-8') as _f:
                        _f.write(ak_inp)
                    st.success("✅ Clave guardada. Ya puedes subir fotos.")
                    st.rerun()
                else:
                    st.error("La clave debe empezar con 'sk-'")
        else:
            st.caption("✅ API Key configurada")
            if st.button("🔑 Cambiar clave API", key="change_ak"):
                st.session_state.ak = ''
                if os.path.exists(CFG_KEY): os.remove(CFG_KEY)
                st.rerun()

        if st.session_state.ak:
            img_up = st.file_uploader(
                "Sube la foto — JPG, PNG o foto del celular",
                type=['jpg','jpeg','png','webp'], key="img_hoja"
            )

            if img_up:
                col_img, col_btn = st.columns([2, 1])
                with col_img:
                    st.image(img_up, use_container_width=True)
                with col_btn:
                    st.markdown(f"**{img_up.name}**  \n{img_up.size/1024:.0f} KB")
                    leer_foto = st.button("📖 Leer medicamentos", type="primary",
                                         key="leer_foto", use_container_width=True)

                if leer_foto:
                    with st.spinner("Claude está leyendo la hoja..."):
                        try:
                            import anthropic, base64 as _b64
                            img_up.seek(0)
                            datos_img  = _b64.standard_b64encode(img_up.read()).decode()
                            media_type = img_up.type or 'image/jpeg'

                            cli  = anthropic.Anthropic(api_key=st.session_state.ak)
                            resp = cli.messages.create(
                                model="claude-opus-4-8",   # Opus actual; vision alta-res mejora lectura de manuscritos
                                max_tokens=2048,            # margen para hojas largas (evita corte silencioso de la lista)
                                messages=[{"role":"user","content":[
                                    {"type":"image","source":{"type":"base64",
                                     "media_type":media_type,"data":datos_img}},
                                    {"type":"text","text":(
                                        "Esta es una hoja de pedido de farmacia. "
                                        "Extrae TODOS los medicamentos que veas escritos, uno por línea. "
                                        "Si hay una cantidad anotada junto al nombre, ponla después de '|'. "
                                        "Solo la lista, sin explicaciones. Ejemplo:\n"
                                        "PARACETAMOL 500 MG | 200\n"
                                        "LOSARTAN 50 MG\n"
                                        "METFORMINA 850 MG | 150"
                                    )}
                                ]}]
                            )
                            texto_leido = next((b.text for b in resp.content
                                                if b.type == "text"), "").strip()
                            st.session_state['hoja_texto_raw'] = texto_leido
                            st.success(f"✅ Claude detectó {len([l for l in texto_leido.split(chr(10)) if l.strip()])} medicamentos.")
                        except Exception as e:
                            if 'auth' in str(e).lower():
                                st.error("❌ API Key inválida. Cámbiala arriba.")
                            else:
                                st.error(f"❌ Error: {e}")

                # Mostrar texto leído (editable por si Claude se equivocó)
                if 'hoja_texto_raw' in st.session_state:
                    st.markdown("**Texto detectado** _(puedes corregir antes de homologar)_:")
                    texto_para_procesar = st.text_area(
                        "texto_foto", value=st.session_state.hoja_texto_raw,
                        height=180, label_visibility="collapsed", key="txt_foto_edit"
                    )

    # ══ MODO TEXTO ═════════════════════════════════════════════════════════════
    else:
        st.caption("Escribe o pega los medicamentos — uno por línea. "
                   "Puedes abreviar, sin acento, en mayúsculas o minúsculas.")
        texto_para_procesar = st.text_area(
            "Lista de medicamentos",
            height=200,
            placeholder=(
                "PARACETAMOL 500\n"
                "losartan 50mg\n"
                "Omeprazol\n"
                "Metformina 850\n"
                "insulina glargina\n"
                "Enalapril 10"
            ),
            label_visibility="collapsed",
            key="txt_manual"
        )

    # ══ HOMOLOGAR (común a ambos modos) ════════════════════════════════════════
    st.markdown("")
    col_hom, col_lim = st.columns([3, 1])
    with col_hom:
        homologar = st.button("🔗 Homologar y generar sugerido", type="primary",
                              key="btn_homologar", use_container_width=True,
                              disabled=not bool(texto_para_procesar and texto_para_procesar.strip()))
    with col_lim:
        if st.button("🗑️ Limpiar todo", key="btn_limpiar_todo", use_container_width=True):
            for k in ['hoja_resultados','hoja_texto_raw']:
                st.session_state.pop(k, None)
            st.rerun()

    # ── Motor de homologación ─────────────────────────────────────────────────
    if homologar and texto_para_procesar.strip():
        lineas = [l.strip() for l in texto_para_procesar.strip().split('\n') if l.strip()]
        resultados_h = []

        for linea in lineas:
            # Separar cantidad manuscrita si viene con |
            cant_ms = None
            nombre_raw = linea
            if '|' in linea:
                partes = linea.split('|', 1)
                nombre_raw = partes[0].strip()
                try:
                    cant_ms = int(''.join(c for c in partes[1] if c.isdigit()))
                except Exception:
                    pass

            if not nombre_raw:
                continue

            # Fuzzy match contra universo AA
            matches = rfprocess.extract(
                _norm(nombre_raw), list(idx_uni.keys()),
                scorer=fuzz.token_set_ratio, limit=1
            )
            if not matches or matches[0][1] < 50:
                resultados_h.append({'escrito': linea, 'med': '— sin coincidencia —',
                                      'score': 0, 'ok': False,
                                      'stock':0,'cob':0.0,'sugerido':0,
                                      'criticidad':'','accion':'','cant_ms': cant_ms})
                continue

            mejor_norm, score, _ = matches[0]
            mejor_med = idx_uni[mejor_norm]

            # Datos del maestro
            fila = df_ped_h[df_ped_h['Medicamento'] == mejor_med]
            if len(fila):
                p        = fila.iloc[0]
                sugerido = int(float(p.get(col_sug, 0) or 0))
                crit     = str(p.get('Criticidad', '') or '')
                stock    = int(float(p.get('Stock_Farm_Actual' if es_fh
                                           else 'Stock_Bod_Actual', 0) or 0))
                cob      = round(float(p.get('Cob_Farm_Actual_Dias' if es_fh
                                             else 'Cob_Bod_Actual_Dias', 0) or 0), 1)
                accion1  = str(p.get('Accion_1_Traspaso_Bodega' if es_fh
                                     else 'Accion_1_Traspaso_Hospital', '') or '')
            else:
                sugerido = 0; crit = ''; stock = 0; cob = 0.0; accion1 = ''

            resultados_h.append({
                'escrito'   : nombre_raw,
                'med'       : mejor_med,
                'score'     : score,
                'stock'     : stock,
                'cob'       : cob,
                'sugerido'  : sugerido,
                'criticidad': crit,
                'accion'    : accion1,
                'cant_ms'   : cant_ms,
                'ok'        : True,
            })

        st.session_state['hoja_resultados'] = resultados_h

    # ── Tabla de resultados ───────────────────────────────────────────────────
    if 'hoja_resultados' in st.session_state:
        resultados_h = st.session_state['hoja_resultados']
        ok  = [r for r in resultados_h if r['ok']]
        nok = [r for r in resultados_h if not r['ok']]

        st.markdown("---")
        st.markdown(f"### Resultado homologación — {len(ok)} coincidencias · {len(nok)} sin resultado")

        if nok:
            with st.expander(f"⚠️ {len(nok)} sin coincidencia — revisa ortografía", expanded=False):
                for r in nok:
                    st.markdown(f"- `{r['escrito']}`")

        if ok:
            # Cabecera
            hdr = st.columns([4, 1, 1, 2, 2])
            for col, lbl in zip(hdr, ["Medicamento del sistema","Stock","Cob.(d)","Sugerido maestro","Cantidad a pedir"]):
                col.markdown(f"**{lbl}**")
            st.divider()

            for i, r in enumerate(ok):
                emo   = crit_emoji(r['criticidad'])
                score = r['score']
                badge = "🟢" if score >= 88 else ("🟡" if score >= 70 else "🟠")

                c1, c2, c3, c4, c5 = st.columns([4, 1, 1, 2, 2])
                c1.markdown(
                    f"{emo} **{r['med']}**  \n"
                    f"<span style='font-size:.76rem;color:#999'>"
                    f"Leído: *{r['escrito']}*  {badge} {score}%"
                    f"{'  ·  manuscrita: ' + str(r['cant_ms']) if r['cant_ms'] else ''}"
                    f"</span>",
                    unsafe_allow_html=True
                )
                c2.markdown(f"{r['stock']:,}")
                c3.markdown(f"{r['cob']:.1f}")
                c4.markdown(
                    f"<div style='text-align:center'>"
                    f"<span style='font-size:1.15rem;font-weight:700;color:#1A237E'>"
                    f"{r['sugerido']:,}</span><br>"
                    f"<span style='font-size:.72rem;color:#888'>por maestro</span></div>",
                    unsafe_allow_html=True
                )
                val_ini = r['cant_ms'] if r['cant_ms'] else r['sugerido']
                c5.number_input("", min_value=0, value=int(val_ini),
                                step=1, key=f"hqty_{i}",
                                label_visibility="collapsed")

            st.markdown("---")
            ca, cb = st.columns([2, 1])
            with ca:
                if st.button("➕ Agregar todos al pedido", type="primary",
                             key="add_hoja_ok", use_container_width=True):
                    agg = 0
                    for i, r in enumerate(ok):
                        qty = int(st.session_state.get(f"hqty_{i}", 0) or 0)
                        if qty > 0:
                            st.session_state.pedido[r['med']] = {
                                'sugerido'  : r['sugerido'],
                                'confirmado': qty,
                                'criticidad': r['criticidad'] or '-',
                                'accion1'   : r['accion'],
                                'accion2'   : '',
                                'notas'     : f"Hoja: {r['escrito'][:25]}",
                            }
                            agg += 1
                    st.success(f"✅ {agg} medicamentos agregados → pestaña **Buscar y armar pedido**")
            with cb:
                if st.button("🗑️ Limpiar", key="limpiar_ok_hoja", use_container_width=True):
                    st.session_state.pop('hoja_resultados', None)
                    st.rerun()

# PESTAÑA 5 — BUSCAR Y ARMAR PEDIDO
# ══════════════════════════════════════════════════════════════════════
with tab_buscar:
    st.markdown("## 🔍 Buscar medicamento y armar pedido")

    # Sub-modo
    modo = st.radio(
        "¿Para qué estás pidiendo?",
        ["Farmacia AA → Bodega AA  (5 dias hábiles)",
         "Bodega AA → Hospital / Externo  (10 dias hábiles)"],
        horizontal=True,
    )

    if "Farmacia" in modo:
        df_pedidos   = df_farm
        col_sugerido = 'Necesidad_5D_Farm'
        col_stock    = 'Stock_Farm_Actual'
        col_cob      = 'Cob_Farm_Actual_Dias'
        col_consumo  = 'Consumo_5D_Trend'
        col_accion1  = 'Accion_1_Traspaso_Bodega'
        col_accion2  = 'Accion_2_Gestion_Externa'
        titulo_modo  = "FARMACIA AA → BODEGA AA"
        color_hex    = "880E4F"
        dias_label   = "5 dias habiles"
    else:
        df_pedidos   = df_bod
        col_sugerido = 'Reponer_Bodega'
        col_stock    = 'Stock_Bod_Post_Traspaso'
        col_cob      = 'Cob_Bod_Post_Dias'
        col_consumo  = 'Consumo_10D_Trend'
        col_accion1  = 'Accion_1_Traspaso_Hospital'
        col_accion2  = 'Accion_2_Compra_Externa'
        titulo_modo  = "BODEGA AA → EXTERNO"
        color_hex    = "1A237E"
        dias_label   = "2 semanas (10 dias hab.)"

    st.markdown("---")
    st.markdown("### Buscar medicamento")
    texto = st.text_input(
        "Escribe parte del nombre (minimo 3 letras)",
        placeholder="Ej: PARACETAMOL  /  LOSARTAN  /  INSULINA  /  OMEPRAZOL",
        label_visibility="collapsed",
    )

    if texto and len(texto.strip()) >= 3:
        terminos   = texto.upper().strip().split()
        resultados = [m for m in todos_meds if all(t in m for t in terminos)]

        if not resultados:
            st.warning(f"No se encontro '{texto}' entre los {len(todos_meds)} medicamentos.")
        else:
            med = st.selectbox(
                f"{len(resultados)} resultado(s) — selecciona:",
                resultados,
            )

            rs = df_stock[df_stock['Medicamento'] == med]
            rp = df_pedidos[df_pedidos['Medicamento'] == med]

            st.markdown("---")

            if len(rs):
                s = rs.iloc[0]
                cdl = float(_get(s, 'CDL_DiasHab', 0))
                cmp = float(_get(s, 'CMP_Mensual_22d', 0))
                sf  = float(_get(s, 'Stock_Farmacia_AA', 0))
                sb  = float(_get(s, 'Stock_Bodega_AA', 0))
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("CDL (ud/dia habil)",  f"{cdl:,.2f}")
                c2.metric("CMP mensual (x22d)",  f"{cmp:,.0f}")
                c3.metric("Stock Farmacia",       f"{sf:,.0f} ud")
                c4.metric("Stock Bodega AA",      f"{sb:,.0f} ud")

            sugerido = 0
            accion1 = accion2 = criticidad = ''

            if len(rp):
                p          = rp.iloc[0]
                sugerido   = int(float(_get(p, col_sugerido, 0)))
                consumo_t  = float(_get(p, col_consumo, 0))
                cob_act    = float(_get(p, col_cob, 0))
                factor     = float(_get(p, 'Factor_Carga_5D', 1))
                criticidad = str(_get(p, 'Criticidad', '5-OK'))
                accion1    = str(_get(p, col_accion1, '') or '')
                accion2    = str(_get(p, col_accion2, '') or '')
                pico       = str(_get(p, 'Semana_Pico_Hist', '') or '')
                css_c      = crit_class(criticidad)
                emo_c      = crit_emoji(criticidad)

                st.markdown(f"""
                <div class='bloque-med {css_c}'>
                  <b>{emo_c} {criticidad}</b> &nbsp;·&nbsp;
                  Semana pico historica: <b>{pico}</b> &nbsp;·&nbsp;
                  Factor carga esta semana: <b>{factor:.0%}</b> &nbsp;·&nbsp;
                  Cobertura actual: <b>{cob_act:.1f} dias</b>
                </div>""", unsafe_allow_html=True)

                col_a, col_b = st.columns(2)
                with col_a:
                    st.info(f"📦 Consumo proyectado ({dias_label}): **{consumo_t:,.0f} ud.**")
                    if accion1:
                        st.success(f"✅ **Accion 1:** {accion1}")
                with col_b:
                    st.warning(f"🛒 **Cantidad sugerida: {sugerido:,} ud.**")
                    if accion2:
                        st.error(f"⚠️ **Accion 2:** {accion2}")

            elif len(rs):
                s   = rs.iloc[0]
                cdl = float(_get(s, 'CDL_DiasHab', 0))
                stk = float(_get(s, 'Stock_Farmacia_AA' if "Farmacia" in modo else 'Stock_Bodega_AA', 0))
                cob = round(stk / cdl, 1) if cdl > 0 else 9999
                st.success(f"✅ Stock suficiente — cobertura {cob:.1f} dias. No requiere pedido en este ciclo.")

            st.markdown("")
            col_qty, col_notas, col_btn = st.columns([2, 4, 2])
            with col_qty:
                qty = st.number_input("Cantidad a pedir", min_value=0,
                                      value=sugerido, step=1, key=f"qty_{med}")
            with col_notas:
                notas = st.text_input("Notas (opcional)", key=f"notas_{med}")
            with col_btn:
                st.markdown("<div style='margin-top:28px'>", unsafe_allow_html=True)
                if st.button("➕ Agregar al pedido", type="primary", key=f"add_{med}"):
                    st.session_state.pedido[med] = {
                        'sugerido': sugerido, 'confirmado': int(qty),
                        'criticidad': criticidad or '-',
                        'accion1': accion1, 'accion2': accion2, 'notas': notas,
                    }
                    st.success(f"✅ {med[:45]}... agregado al pedido")
                st.markdown("</div>", unsafe_allow_html=True)

    # ── Pedido acumulado ──────────────────────────────────────────────────────
    st.markdown("---")
    n_ped = len(st.session_state.pedido)
    st.markdown(f"### 🛒 Mi pedido  ({n_ped} medicamento{'s' if n_ped != 1 else ''})")

    if st.session_state.pedido:
        filas = []
        for m, info in st.session_state.pedido.items():
            filas.append({
                'Medicamento': m,
                'Criticidad':  info['criticidad'],
                'Sugerido':    info['sugerido'],
                'A Pedir':     info['confirmado'],
                'Accion 1':    info['accion1'],
                'Accion 2':    info['accion2'],
                'Notas':       info['notas'],
            })
        df_ped = pd.DataFrame(filas)

        st.dataframe(
            estilo_tabla(df_ped),
            use_container_width=True,
            hide_index=True,
            column_config={
                'Medicamento': st.column_config.TextColumn("Medicamento", width="large"),
                'Criticidad':  st.column_config.TextColumn("Criticidad",  width="small"),
                'Sugerido':    st.column_config.NumberColumn("Sugerido",  format="%d ud"),
                'A Pedir':     st.column_config.NumberColumn("A Pedir",   format="%d ud"),
                'Accion 1':    st.column_config.TextColumn("Accion 1",    width="medium"),
                'Accion 2':    st.column_config.TextColumn("Accion 2",    width="medium"),
                'Notas':       st.column_config.TextColumn("Notas",       width="medium"),
            },
        )

        # ── Exportar Excel ────────────────────────────────────────────────────
        def build_excel():
            wb  = Workbook()
            ws  = wb.active
            ws.title = "Pedido"
            fills = {k: crit_fill(k) for k in CRIT_FILL_HEX}
            ws['A1'] = f'PEDIDO {titulo_modo} — S{semana} {hoy.strftime("%d/%m/%Y")}'
            ws['A1'].fill = PatternFill('solid', fgColor=color_hex)
            ws['A1'].font = Font(bold=True, color='FFFFFF', name='Arial', size=13)
            ws.merge_cells('A1:G1')
            ws.row_dimensions[1].height = 26
            headers = ['Medicamento','Criticidad','Sugerido','A Pedir',
                       'Accion 1','Accion 2','Notas']
            hf    = PatternFill('solid', fgColor='1F4E78')
            hfont = Font(bold=True, color='FFFFFF', name='Arial', size=11)
            for ci, h in enumerate(headers, 1):
                c = ws.cell(row=2, column=ci, value=h)
                c.fill = hf; c.font = hfont
                c.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[2].height = 22
            for ri, (_, row) in enumerate(df_ped.iterrows(), 3):
                crit = str(row.get('Criticidad', ''))
                fill = fills.get(crit, PatternFill('solid', fgColor='FFFFFF'))
                vals = [row.get('Medicamento',''), crit,
                        int(row.get('Sugerido',0) or 0),
                        int(row.get('A Pedir',0) or 0),
                        str(row.get('Accion 1','') or ''),
                        str(row.get('Accion 2','') or ''),
                        str(row.get('Notas','') or '')]
                for ci, v in enumerate(vals, 1):
                    c = ws.cell(row=ri, column=ci, value=v)
                    c.fill = fill
                    c.font = Font(name='Arial', size=11,
                                  bold=(crit == '1-CRITICO'),
                                  color='FFFFFF' if crit == '1-CRITICO' else '000000')
                    c.alignment = Alignment(vertical='center')
            for ci, w in enumerate([52,14,11,11,42,30,22], 1):
                ws.column_dimensions[get_column_letter(ci)].width = w
            ws.freeze_panes = 'A3'
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            return buf

        col_dl, col_cl = st.columns([3, 1])
        with col_dl:
            fname = f"Pedido_AA_{hoy.strftime('%Y%m%d_%H%M')}.xlsx"
            st.download_button(
                label="📥 Descargar pedido Excel",
                data=build_excel(),
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        with col_cl:
            if st.button("🗑️ Limpiar pedido", use_container_width=True):
                st.session_state.pedido = {}
                st.rerun()

    else:
        st.info("☝️ Busca un medicamento arriba, o usa el botón **Agregar los 20 al pedido** en alguna de las otras pestañas.")

# ══════════════════════════════════════════════════════════════════════
# PESTAÑA FEEDBACK — Diagnóstico automático + Sugerencias del operador
# ══════════════════════════════════════════════════════════════════════
with tab_feedback:
    import json as _json

    st.markdown("## 💬 Diagnóstico y Sugerencias")
    st.markdown(
        "Esta pestaña muestra el **estado actual de los datos** y te permite "
        "registrar sugerencias o reportar problemas para mejorar la aplicación."
    )

    # ── Diagnóstico automático ────────────────────────────────────────────────
    st.markdown("### 🔍 Diagnóstico automático")

    horas_ant_fb = (datetime.now() - datos['mtime']).total_seconds() / 3600
    dias_ant_fb  = horas_ant_fb / 24

    col_d1, col_d2, col_d3, col_d4 = st.columns(4)
    col_d1.metric("Medicamentos universo",  f"{len(todos_meds)}")
    col_d2.metric("Antigüedad del dato",
                  f"{int(horas_ant_fb)}h" if horas_ant_fb < 48 else f"{dias_ant_fb:.1f} días",
                  delta="OK" if horas_ant_fb < 24 else "DESACTUALIZADO",
                  delta_color="normal" if horas_ant_fb < 24 else "inverse")

    n_crit_farm = 0
    n_sin_stock = 0
    if 'Criticidad' in df_farm.columns:
        n_crit_farm = int((df_farm['Criticidad'].isin(['1-CRITICO','2-URGENTE'])).sum())
    if 'Stock_Farm_Actual' in df_farm.columns:
        n_sin_stock = int((pd.to_numeric(df_farm['Stock_Farm_Actual'], errors='coerce').fillna(0) == 0).sum())

    col_d3.metric("Críticos + Urgentes (farm.)", f"{n_crit_farm}",
                  delta="revisar" if n_crit_farm > 5 else "OK",
                  delta_color="inverse" if n_crit_farm > 5 else "normal")
    col_d4.metric("Sin stock en farmacia", f"{n_sin_stock}",
                  delta="alerta" if n_sin_stock > 0 else "OK",
                  delta_color="inverse" if n_sin_stock > 0 else "normal")

    st.markdown("---")

    # Chequeos automáticos
    issues   = []
    ok_items = []

    if horas_ant_fb > 24:
        issues.append(f"🔴 **Datos desactualizados** — el archivo tiene {dias_ant_fb:.1f} días. "
                      "Ejecuta `AUTO_SSASUR.bat` para obtener stock fresco.")
    else:
        ok_items.append(f"✅ Datos actualizados hace {int(horas_ant_fb)}h.")

    if n_sin_stock > 0:
        meds_0 = df_farm[pd.to_numeric(df_farm.get('Stock_Farm_Actual', pd.Series()), errors='coerce').fillna(0) == 0]['Medicamento'].tolist()
        issues.append(f"🔴 **{n_sin_stock} medicamento(s) con stock = 0** en farmacia: "
                      + ", ".join(meds_0[:5]) + ("..." if len(meds_0) > 5 else ""))
    else:
        ok_items.append("✅ Ningún medicamento con stock cero en farmacia.")

    if n_crit_farm > 10:
        issues.append(f"🟠 **{n_crit_farm} medicamentos Crítico/Urgente** — revisar pestaña 📋 TOP 20.")
    elif n_crit_farm > 0:
        ok_items.append(f"🟡 {n_crit_farm} medicamentos en nivel Crítico/Urgente (dentro de rango).")
    else:
        ok_items.append("✅ Ningún medicamento en estado Crítico/Urgente.")

    if not df_dial_farm.empty:
        ok_items.append(f"✅ Datos de diálisis cargados: {len(df_dial_farm)} medicamentos.")
    else:
        issues.append("🟡 **Sin datos de diálisis** — regenera el Consolidado con `EJECUTAR_MAESTRO.bat`.")

    if issues:
        st.markdown("**Problemas detectados:**")
        for iss in issues:
            st.markdown(f"- {iss}")
    if ok_items:
        with st.expander("Ver chequeos OK", expanded=False):
            for ok in ok_items:
                st.markdown(f"- {ok}")

    st.markdown("---")

    # ── Formulario de sugerencias ─────────────────────────────────────────────
    st.markdown("### ✍️ Registrar sugerencia o problema")
    st.caption("Tus comentarios quedan guardados en `feedback.json` dentro de la carpeta de la app.")

    _FEEDBACK_FILE = os.path.join(WORK_DIR, "feedback.json")

    def _load_feedback():
        if os.path.exists(_FEEDBACK_FILE):
            try:
                with open(_FEEDBACK_FILE, encoding='utf-8') as _f:
                    return _json.load(_f)
            except Exception:
                return []
        return []

    def _save_feedback(entries):
        with open(_FEEDBACK_FILE, 'w', encoding='utf-8') as _f:
            _json.dump(entries, _f, ensure_ascii=False, indent=2)

    cat_opts = ["🐛 Error / dato incorrecto", "💡 Mejora sugerida",
                "❓ Consulta / duda", "✅ Funciona bien — comentario positivo"]

    with st.form("form_feedback", clear_on_submit=True):
        fb_cat  = st.selectbox("Categoría", cat_opts)
        fb_pest = st.selectbox("Pestaña afectada (opcional)",
                               ["(ninguna / general)", "TOP 20 Farmacia", "TOP 20 Bodega",
                                "Pedido a Bodega AA", "Pedido a Bodega Fármacos",
                                "Diálisis", "Faltantes", "Foto Manuscrita", "Buscar"])
        fb_txt  = st.text_area("Describe el problema o sugerencia", height=120,
                               placeholder="Ej: El stock de AMLODIPINO aparece diferente al sistema SSASUR...")
        fb_user = st.text_input("Tu nombre (opcional)", placeholder="Ej: María / Turno mañana")
        enviado = st.form_submit_button("💾 Guardar", type="primary", use_container_width=True)

    if enviado and fb_txt.strip():
        entries = _load_feedback()
        entries.append({
            "fecha"    : datetime.now().strftime("%Y-%m-%d %H:%M"),
            "categoria": fb_cat,
            "pestana"  : fb_pest,
            "texto"    : fb_txt.strip(),
            "usuario"  : fb_user.strip() or "Anónimo",
        })
        _save_feedback(entries)
        st.success("✅ Sugerencia guardada. ¡Gracias por el feedback!")
    elif enviado:
        st.warning("Escribe algo en el campo de descripción.")

    # ── Historial de feedback ────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📋 Historial de sugerencias")
    entries_hist = _load_feedback()
    if entries_hist:
        for e in reversed(entries_hist[-20:]):
            with st.expander(f"{e['fecha']}  ·  {e['categoria']}  ·  {e['usuario']}"):
                if e.get('pestana') and e['pestana'] != "(ninguna / general)":
                    st.caption(f"Pestaña: {e['pestana']}")
                st.markdown(e['texto'])
    else:
        st.info("Aún no hay sugerencias registradas. ¡Sé el primero!")

# ─── Footer ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(
    f"Maestro AA  ·  Universo: {len(todos_meds)} medicamentos  ·  "
    f"Datos: {datos['mtime'].strftime('%d/%m/%Y %H:%M')}"
)
