"""Centinela — reporte semanal de medicamentos centinela, campaña de invierno."""

import os
import sys
import json
import glob
import subprocess
from datetime import datetime

import pandas as pd
import streamlit as st

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from estilo_maestro import inject_css, hero, guard_badge
from utils_aa import UMBRAL_DIAS_STALE
from centinela_reporte import _encontrar_archivos, leer_stock

st.set_page_config(page_title="Centinela — Maestro AA", page_icon="🩺",
                    layout="wide", initial_sidebar_state="expanded")
inject_css()
hero("🩺", "Centinela", "Campaña de invierno · Reporte semanal de medicamentos centinela (MINSAL)")


@st.cache_data(ttl=300, show_spinner=False)
def _fecha_max_recetas(paths_mtimes):
    maxima = None
    for p, _ in paths_mtimes:
        try:
            col = pd.read_csv(p, encoding="latin-1", sep=None, engine="python",
                               usecols=["Fecha Entrega Receta"])
            m = pd.to_datetime(col["Fecha Entrega Receta"], dayfirst=True, errors="coerce").max()
            if pd.notna(m) and (maxima is None or m > maxima):
                maxima = m
        except Exception:
            continue
    return maxima.date() if maxima is not None else None


@st.cache_data(ttl=300, show_spinner=False)
def _fecha_stock(xlsx_path_mtime):
    xlsx_path, _ = xlsx_path_mtime
    try:
        _, xlsx_fecha = leer_stock(xlsx_path)
        return datetime.strptime(xlsx_fecha, "%d/%m/%Y").date() if xlsx_fecha else None
    except Exception:
        return None

_ALERTA = {
    "q":  ("🔴 Quiebre de stock",        "st-crit"),
    "a":  ("🟠 Rotación < 4 semanas",    "st-warn"),
    "m":  ("🟡 Rotación < 8 semanas",    "st-warn"),
    "w":  ("⚪ Sin movimiento",          "st-idle"),
    "s":  ("🟢 Estable",                 "st-ok"),
    "na": ("⚪ No aplica",               "st-idle"),
}

with st.sidebar:
    st.markdown("## 🩺 Centinela")
    st.caption("Reporte semanal — se genera habitualmente los **lunes** con la semana epidemiológica completa anterior.")
    st.markdown("---")
    st.markdown("### 🛡️ Guard de frescura")
    st.caption(f"Si el dato auto-detectado tiene más de **{UMBRAL_DIAS_STALE} días**, la generación se bloquea "
               f"(evita repetir el incidente S.52: reportar la semana equivocada por datos viejos).")
    try:
        csvs, xlsx_path = _encontrar_archivos()
        fecha_rec = _fecha_max_recetas(tuple((str(p), os.path.getmtime(p)) for p in csvs))
        fecha_stk = _fecha_stock((str(xlsx_path), os.path.getmtime(xlsx_path)))
        st.markdown(guard_badge(fecha_rec, UMBRAL_DIAS_STALE, "Sábana recetas"), unsafe_allow_html=True)
        st.markdown(guard_badge(fecha_stk, UMBRAL_DIAS_STALE, "Stock"), unsafe_allow_html=True)
    except FileNotFoundError as e:
        st.warning(f"No se pudo evaluar la frescura: {e}")
    st.markdown("---")
    st.markdown("### Generar reporte")
    semana_manual = st.number_input("Forzar semana epidemiológica (opcional)", min_value=0, max_value=53,
                                     value=0, help="0 = automático (última semana completa)")
    if st.button("⚙️ Generar ahora", type="primary", use_container_width=True):
        cmd = ["py", "centinela_reporte.py", "--no-pause"]
        if semana_manual:
            cmd += ["--semana", str(int(semana_manual))]
        with st.spinner("Generando reporte centinela..."):
            res = subprocess.run(cmd, cwd=_ROOT, capture_output=True, text=True,
                                  encoding="utf-8", errors="replace")
        salida = res.stdout or ""
        if res.returncode == 0:
            st.success("Reporte generado.")
            with st.expander("Detalle"):
                st.code(salida or "(sin salida)")
            st.cache_data.clear()
            st.rerun()
        elif "desactualizad" in salida:
            st.warning("🛡️ **Bloqueado por el guard de frescura** — los datos auto-detectados son viejos, "
                       "no se generó un reporte con la semana equivocada.")
            st.code(salida)
            st.caption("Ejecuta AUTO_SSASUR.bat (o el botón de Gestión Territorial) para refrescar los datos.")
        else:
            st.error("Falló la generación.")
            st.code(res.stderr or salida or "(sin salida)")

carpetas = sorted(
    glob.glob(os.path.join(_ROOT, "Centinela_Reportes", "S*")),
    key=lambda p: int(os.path.basename(p).lstrip("S")) if os.path.basename(p).lstrip("S").isdigit() else -1,
    reverse=True,
)

if not carpetas:
    st.info("Todavía no se ha generado ningún reporte centinela. Usa **⚙️ Generar ahora** en la barra lateral.")
    st.stop()

nombres = [os.path.basename(c) for c in carpetas]
idx = st.selectbox("Semana epidemiológica", options=range(len(carpetas)), format_func=lambda i: nombres[i])
carpeta = carpetas[idx]
srep = nombres[idx]

jsons = glob.glob(os.path.join(carpeta, "*.json"))
pdfs = glob.glob(os.path.join(carpeta, "*.pdf"))

if not jsons:
    st.warning(f"No se encontró el JSON de resultados en {carpeta}.")
    st.stop()

with open(jsons[0], "r", encoding="utf-8") as fh:
    data = json.load(fh)

resultados = data.get("resultados", [])
mtime = datetime.fromtimestamp(os.path.getmtime(jsons[0]))

c1, c2, c3, c4 = st.columns(4)
c1.metric("Semana reportada", f"S.{data.get('srep', '?')}")
c2.metric("Stock al", data.get("xlsx_fecha", "—"))
c3.metric("Recetas procesadas", f"{data.get('num_recetas', 0):,}")
n_alerta = sum(1 for r in resultados if r.get("alerta") in ("q", "a"))
c4.metric("Medicamentos en alerta", f"{n_alerta} ⚠️" if n_alerta else "0 ✅")

st.caption(f"Generado: {mtime.strftime('%d/%m/%Y %H:%M')} · Rango histórico S.{data.get('sem_min')}–S.{data.get('sem_max')}")

if pdfs:
    with open(pdfs[0], "rb") as fh:
        st.download_button("📄 Descargar PDF (para MINSAL)", data=fh.read(),
                            file_name=os.path.basename(pdfs[0]), mime="application/pdf",
                            type="primary")
else:
    st.warning("No se generó PDF para esta semana (¿faltó reportlab?).")

st.markdown("---")
st.markdown("### Detalle por medicamento")

filas = []
for r in resultados:
    if r.get("na"):
        continue
    label, _ = _ALERTA.get(r.get("alerta"), ("—", "st-idle"))
    filas.append({
        "Medicamento": r.get("minsal", ""),
        "Stock": r.get("stk", 0),
        "Consumo S.actual": r.get("consumo", 0),
        "Proyección": r.get("proy", 0),
        "Rotación (sem.)": r.get("rot"),
        "Var. vs sem. ant. (%)": r.get("var_sem"),
        "Var. vs histórico (%)": r.get("var_hist"),
        "Alerta": label,
    })

df = pd.DataFrame(filas)
st.dataframe(df, use_container_width=True, hide_index=True, height=min(50 + len(df) * 35, 500))
