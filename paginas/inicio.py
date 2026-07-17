"""Inicio — hub del Centro de Operaciones Maestro AA."""

import os
import sys
import glob
from datetime import datetime

import streamlit as st

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from estilo_maestro import inject_css, hero, fecha_relativa, status_badge

st.set_page_config(
    page_title="Maestro AA — Centro de Operaciones",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)
inject_css()

hoy = datetime.now()
_dia_es = {'Mon': 'lun', 'Tue': 'mar', 'Wed': 'mié', 'Thu': 'jue', 'Fri': 'vie', 'Sat': 'sáb', 'Sun': 'dom'}
_mes_es = {1: 'ene', 2: 'feb', 3: 'mar', 4: 'abr', 5: 'may', 6: 'jun',
           7: 'jul', 8: 'ago', 9: 'sep', 10: 'oct', 11: 'nov', 12: 'dic'}
_fecha_chip = f"{_dia_es.get(hoy.strftime('%a'), '')} {hoy.day} {_mes_es.get(hoy.month, '')}"

hero("🏥", "Maestro AA", "Centro de Operaciones · Farmacia Atención Abierta · Hospital de Pitrufquén",
     chip_top=_fecha_chip, chip_bottom=hoy.strftime('%H:%M'))


# ─── Helpers de estado ────────────────────────────────────────────────────
def _latest(patron):
    cand = [f for f in glob.glob(patron) if not os.path.basename(f).startswith("~$")]
    return max(cand, key=os.path.getmtime) if cand else None


def _card_status(patron, sin_dato="Sin datos aún"):
    f = _latest(patron)
    if not f:
        return status_badge(f"⚪ {sin_dato}", "st-idle")
    txt, cls = fecha_relativa(datetime.fromtimestamp(os.path.getmtime(f)))
    icono = {"st-ok": "🟢", "st-warn": "🟡", "st-crit": "🔴"}.get(cls, "⚪")
    return status_badge(f"{icono} {txt}", cls)


# Pedidos AA
_status_pedidos = _card_status(os.path.join(_ROOT, "Consolidado_AA_MAESTRO*.xlsx"),
                                "Ejecuta EJECUTAR_MAESTRO.bat")

# Pedidos Fusionados
_status_fusion = _card_status(os.path.join(_ROOT, "Pedido_Fusion_AA_*.xlsx"),
                               "Aún no se ha generado")

# Centinela
_status_centinela = _card_status(os.path.join(_ROOT, "Centinela_Reportes", "S*", "*.pdf"),
                                  "Sin reportes generados")

# Recetas Cheque — ¿existe formulario del mes en curso?
try:
    from recetas_cheque import listar_formularios, RCH_DIR
    _formularios = listar_formularios(RCH_DIR)
    _mes_actual = any(a == hoy.year and m == hoy.month for _, a, m in _formularios)
    if _mes_actual:
        _status_recetas = status_badge(f"🟢 Formulario de {_mes_es.get(hoy.month,'')} listo", "st-ok")
    elif _formularios:
        _status_recetas = status_badge("🟡 Falta formulario del mes en curso", "st-warn")
    else:
        _status_recetas = status_badge("⚪ Sin formularios en la carpeta RCh", "st-idle")
except Exception:
    _status_recetas = status_badge("⚪ Carpeta RCh no disponible en este equipo", "st-idle")

# Gestión Territorial
_gt_dir = os.path.join(_ROOT, os.pardir, "04_Farmacia_Gestion_Territorial")
_status_gt = _card_status(os.path.join(_gt_dir, "reporteGestionTerritorial_*.xlsx"),
                           "Sin descargas GT aún")

st.markdown("### Resumen rápido")
mcol1, mcol2, mcol3 = st.columns(3)
mcol1.metric("Universo AA", "378 medicamentos")
_n_fusion_mes = len([f for f in glob.glob(os.path.join(_ROOT, "Pedido_Fusion_AA_*.xlsx"))
                      if datetime.fromtimestamp(os.path.getmtime(f)).month == hoy.month])
mcol2.metric("Fusiones generadas este mes", f"{_n_fusion_mes}")
_maestro_f = _latest(os.path.join(_ROOT, "Consolidado_AA_MAESTRO*.xlsx"))
if _maestro_f:
    _h = (datetime.now() - datetime.fromtimestamp(os.path.getmtime(_maestro_f))).total_seconds() / 3600
    mcol3.metric("Datos SSASUR", "Al día" if _h < 4 else f"Hace {int(_h)} h")
else:
    mcol3.metric("Datos SSASUR", "—")

st.markdown("### Módulos")
st.caption(
    "Elige un módulo para revisar su estado y trabajar. Cada uno conserva su propio flujo "
    "y también se puede seguir abriendo por separado con su .bat."
)

MODULOS = [
    ("app_pedidos.py", "c-teal", "💊", "Pedidos AA",
     "Pedido a Bodega, faltantes, SGLI y diálisis en un dashboard.",
     _status_pedidos),
    ("paginas/pedidos_fusionados.py", "c-indigo", "🔗", "Pedidos Fusionados",
     "Pedido consolidado del día: Farm↔Bodega + Bodega↔Fármacos + diálisis en un solo Excel.",
     _status_fusion),
    ("paginas/centinela.py", "c-rose", "🩺", "Centinela",
     "Reporte semanal de medicamentos centinela — campaña de invierno, PDF listo para el MINSAL.",
     _status_centinela),
    ("paginas/recetas_cheque_page.py", "c-amber", "🧾", "Recetas Cheque",
     "Registro ISP de recetas cheque (estupefacientes/psicotrópicos) — obligación legal mensual.",
     _status_recetas),
    ("paginas/gestion_territorial.py", "c-sky", "🗺️", "Gestión Territorial",
     "Clasificación y planillas por establecimiento destino de la red Araucanía Sur.",
     _status_gt),
]

cols = st.columns(3)
for i, (target, clase, icono, titulo, sub, status_html) in enumerate(MODULOS):
    with cols[i % 3]:
        st.markdown(f"""
        <div class='mod-card {clase}'>
          <div class='mc-top'><div class='mc-ico'>{icono}</div></div>
          <div class='mc-tit'>{titulo}</div>
          <div class='mc-sub'>{sub}</div>
          {status_html}
        </div>
        """, unsafe_allow_html=True)
        st.page_link(target, label="Abrir →", icon=None, use_container_width=True)
        st.write("")
