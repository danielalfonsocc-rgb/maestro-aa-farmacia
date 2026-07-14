"""Gestión Territorial — clasificación y planillas por establecimiento destino."""

import os
import sys
import json
import glob
import subprocess
from collections import defaultdict
from datetime import datetime

import pandas as pd
import streamlit as st

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from estilo_maestro import inject_css, hero

st.set_page_config(page_title="Gestión Territorial — Maestro AA", page_icon="🗺️",
                    layout="wide", initial_sidebar_state="expanded")
inject_css()
hero("🗺️", "Gestión Territorial", "Clasificación y planillas por establecimiento destino · Red Araucanía Sur")

GT_DIR = os.path.join(_ROOT, os.pardir, "04_Farmacia_Gestion_Territorial")
OUT_GT_DIR = os.path.join(_ROOT, "out_gt")

with st.sidebar:
    st.markdown("## 🗺️ Gestión Territorial")
    st.caption("Descarga desde SSASUR, cruza con el histórico y genera planillas + letreros por destino.")
    st.markdown("---")
    if st.button("🌐 Descargar GT desde SSASUR", type="primary", use_container_width=True,
                 help="Abre un navegador para que inicies sesión en SSASUR (igual que GT.bat)"):
        subprocess.Popen(["py", "AUTO_SSASUR.py", "--gt"], cwd=_ROOT,
                          creationflags=subprocess.CREATE_NEW_CONSOLE)
        st.info("Se abrió una ventana nueva. Inicia sesión en SSASUR ahí; cuando termine, recarga esta página.")
    st.caption("Equivale a **GT.bat**: descarga solo GT (ayer y hoy), sin stock ni maestro.")
    st.markdown("---")
    st.caption(f"Descargas GT: `{os.path.abspath(GT_DIR)}`")
    st.caption(f"Planillas generadas: `{os.path.abspath(OUT_GT_DIR)}`")

# ─── Últimas descargas GT ──────────────────────────────────────────────────
reportes = sorted(
    [f for f in glob.glob(os.path.join(GT_DIR, "reporteGestionTerritorial_*.xlsx"))
     if not os.path.basename(f).startswith("~$")],
    key=os.path.getmtime, reverse=True,
)

if not reportes:
    st.info("Todavía no hay descargas de Gestión Territorial en este equipo. "
            "Usa **🌐 Descargar GT desde SSASUR** en la barra lateral.")
    st.stop()

ultimo_reporte = reportes[0]
mtime_rep = datetime.fromtimestamp(os.path.getmtime(ultimo_reporte))
st.caption(f"Última descarga GT: `{os.path.basename(ultimo_reporte)}` · {mtime_rep.strftime('%d/%m/%Y %H:%M')} "
           f"· {len(reportes)} descarga(s) en total")

# ─── Carpetas de salida (una por rango de fechas) ──────────────────────────
carpetas_out = sorted(
    [d for d in glob.glob(os.path.join(OUT_GT_DIR, "*")) if os.path.isdir(d)],
    key=os.path.getmtime, reverse=True,
)

if not carpetas_out:
    st.warning("Hay descargas GT pero aún no se generó la clasificación/planillas (out_gt/).")
    st.stop()

nombres_out = [os.path.basename(c) for c in carpetas_out]
idx = st.selectbox("Rango procesado", options=range(len(carpetas_out)), format_func=lambda i: nombres_out[i])
carpeta = carpetas_out[idx]

# ─── Resumen desde gt_enriquecido.json ─────────────────────────────────────
json_path = os.path.join(carpeta, "gt_enriquecido.json")
if os.path.exists(json_path):
    with open(json_path, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    regs = data.get("registros", [])

    n_ref = sum(1 for g in regs if g.get("refrigerado"))
    n_con = sum(1 for g in regs if g.get("controlado"))
    n_pen = sum(1 for g in regs if g.get("pendiente"))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Recetas en el rango", f"{len(regs)}")
    c2.metric("❄️ Refrigerados", f"{n_ref}")
    c3.metric("⚠️ Controlados", f"{n_con}")
    c4.metric("⏳ Pendientes", f"{n_pen}")

    st.markdown("### Por establecimiento destino")
    por_dest = defaultdict(lambda: [0, 0, 0, 0])
    for g in regs:
        d = por_dest[g.get("estab_destino", "—")]
        d[0] += 1
        d[1] += bool(g.get("refrigerado"))
        d[2] += bool(g.get("controlado"))
        d[3] += bool(g.get("pendiente"))
    df_dest = pd.DataFrame(
        [{"Destino": k, "Recetas": v[0], "Refrigerados": v[1], "Controlados": v[2], "Pendientes": v[3]}
         for k, v in sorted(por_dest.items())]
    )
    st.dataframe(df_dest, use_container_width=True, hide_index=True)

    with st.expander("Ver detalle por receta (uso interno)"):
        df_det = pd.DataFrame(regs)
        cols_mostrar = [c for c in ["estab_destino", "receta", "paciente", "especialidad",
                                     "periodo", "refrigerado", "controlado", "pendiente"]
                         if c in df_det.columns]
        st.dataframe(df_det[cols_mostrar], use_container_width=True, hide_index=True,
                     height=min(50 + len(df_det) * 32, 500))
else:
    st.warning(f"No se encontró gt_enriquecido.json en {carpeta}.")

# ─── Archivos generados (planillas / letreros / PDFs) ──────────────────────
st.markdown("---")
st.markdown("### Planillas y letreros por destino")

archivos = sorted(glob.glob(os.path.join(carpeta, "*")))
archivos = [f for f in archivos if os.path.isfile(f) and not os.path.basename(f).startswith("~$")]

if not archivos:
    st.info("Aún no se generaron planillas en esta carpeta.")
else:
    por_destino_files = defaultdict(list)
    for f in archivos:
        nombre = os.path.basename(f)
        destino = nombre.split("_")[0] if "_" in nombre else "Otros"
        por_destino_files[destino].append(f)

    for destino, files in sorted(por_destino_files.items()):
        with st.expander(f"📁 {destino}  ({len(files)} archivo(s))"):
            for f in sorted(files):
                colf1, colf2 = st.columns([4, 1])
                colf1.markdown(f"`{os.path.basename(f)}`")
                with open(f, "rb") as fh:
                    colf2.download_button("📥", data=fh.read(), file_name=os.path.basename(f),
                                           key=f"dl_{f}", use_container_width=True)
