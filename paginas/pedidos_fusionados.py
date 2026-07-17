"""Pedidos Fusionados — Pedido_Fusion_AA_<fecha>.xlsx (Farm_Bod + Bod_Farmacos + Diálisis)."""

import os
import sys
import glob
import subprocess
from datetime import datetime

import pandas as pd
import streamlit as st

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from estilo_maestro import inject_css, hero
from aa_colors import crit_nivel

st.set_page_config(page_title="Pedidos Fusionados — Maestro AA", page_icon="🔗",
                    layout="wide", initial_sidebar_state="expanded")
inject_css()
hero("🔗", "Pedidos Fusionados", "Pedido consolidado del día: Farm↔Bodega · Bodega↔Fármacos · Diálisis · Faltantes 30d")

SHEETS = [
    ("Farm_Bod", "📝 Farmacia AA → Bodega AA"),
    ("Bod_Farmacos", "🏭 Bodega AA → Bodega Fármacos"),
    ("Dialisis", "💉 Diálisis (solo S3)"),
    ("Faltantes_AA", "⛔ Faltantes Absolutos (30d)"),
]

# Columna a usar para la métrica "con pedido/reposición > 0" de cada hoja (por
# posición, 0-based) — Faltantes_AA no tiene columna de pedido, así que usa
# "Faltante (ud)" y una etiqueta propia en vez de la genérica.
_COL_METRICA = {"Farm_Bod": 5, "Bod_Farmacos": 5, "Dialisis": 5, "Faltantes_AA": 4}
_LABEL_METRICA = {"Faltantes_AA": "Con faltante > 0"}

_BADGE_CRIT = {
    1: 'background-color:#FEF2F2;color:#B91C1C;font-weight:700',
    2: 'background-color:#FFF7ED;color:#C2410C;font-weight:700',
    3: 'background-color:#FFFBEB;color:#B45309;font-weight:700',
    4: 'background-color:#FFFBEB;color:#B45309;font-weight:700',
}
_BADGE_OK = 'background-color:#F0FDF4;color:#15803D;font-weight:700'


def _estilo(df):
    sty = df.style.set_properties(**{'background-color': '#ffffff', 'color': '#1F2937'})
    if 'Criticidad' in df.columns:
        sty = sty.apply(
            lambda col: [_BADGE_CRIT.get(crit_nivel(str(v)), _BADGE_OK) for v in col],
            subset=['Criticidad'],
        )
    return sty


with st.sidebar:
    st.markdown("## 🔗 Pedidos Fusionados")
    st.caption("4 hojas en un solo Excel: pedido del día + reposición de bodega + diálisis + faltantes absolutos 30d.")
    st.markdown("---")
    st.markdown("### Generar nuevo")
    forzar_dialisis = st.checkbox("Forzar hoja Diálisis", help="Incluir aunque no sea la 3ª semana del mes")
    todos = st.checkbox("Incluir todos los meds", help="Farm_Bod: listar también los sin necesidad de pedido hoy")
    if st.button("⚙️ Generar ahora", type="primary", use_container_width=True):
        cmd = ["py", "pedido_fusion.py"]
        if forzar_dialisis:
            cmd.append("--forzar-dialisis")
        if todos:
            cmd.append("--todos")
        with st.spinner("Generando Pedido_Fusion_AA..."):
            res = subprocess.run(cmd, cwd=_ROOT, capture_output=True, text=True,
                                  encoding="utf-8", errors="replace")
        if res.returncode == 0:
            st.success("Pedido fusión generado.")
            with st.expander("Detalle"):
                st.code(res.stdout or "(sin salida)")
            st.cache_data.clear()
            st.rerun()
        else:
            st.error("Falló la generación.")
            st.code(res.stderr or res.stdout or "(sin salida)")

archivos = sorted(
    [f for f in glob.glob(os.path.join(_ROOT, "Pedido_Fusion_AA_*.xlsx"))
     if not os.path.basename(f).startswith("~$")],
    key=os.path.getmtime, reverse=True,
)

if not archivos:
    st.info("Todavía no se ha generado ningún Pedido_Fusion_AA. Usa **⚙️ Generar ahora** en la barra lateral.")
    st.stop()

nombres = [os.path.basename(f) for f in archivos]
idx = st.selectbox("Archivo", options=range(len(archivos)), format_func=lambda i: nombres[i])
ruta = archivos[idx]
mtime = datetime.fromtimestamp(os.path.getmtime(ruta))

c1, c2 = st.columns([3, 1])
with c1:
    st.caption(f"Generado: {mtime.strftime('%d/%m/%Y %H:%M')}")
with c2:
    with open(ruta, "rb") as fh:
        st.download_button("📥 Descargar Excel", data=fh.read(), file_name=os.path.basename(ruta),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True)

st.markdown("---")

tabs = st.tabs([label for _, label in SHEETS])
for (sheet, label), tab in zip(SHEETS, tabs):
    with tab:
        try:
            df = pd.read_excel(ruta, sheet_name=sheet, header=2, engine="openpyxl")
        except Exception as e:
            st.warning(f"No se pudo leer la hoja «{sheet}»: {e}")
            continue
        df = df.dropna(how="all")
        if df.empty:
            st.success("✅ Sin filas en esta hoja (nada pendiente).")
            continue
        n_pedido = 0
        col_idx = _COL_METRICA.get(sheet, 5)
        if len(df.columns) > col_idx:
            ultima_col = df.columns[col_idx]
            n_pedido = int(pd.to_numeric(df[ultima_col], errors="coerce").fillna(0).gt(0).sum())
        mcol1, mcol2 = st.columns(2)
        mcol1.metric("Medicamentos en la hoja", f"{len(df)}")
        mcol2.metric(_LABEL_METRICA.get(sheet, "Con pedido/reposición > 0"), f"{n_pedido}")
        st.dataframe(_estilo(df), use_container_width=True, hide_index=True,
                     height=min(50 + len(df) * 35, 620))
