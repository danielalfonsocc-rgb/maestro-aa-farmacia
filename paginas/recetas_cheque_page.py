"""Recetas Cheque — Registro ISP de recetas cheque (estupefacientes/psicotrópicos).

No muestra datos de pacientes (RUT): solo folios registrados y metadata del
formulario, en línea con la Ley 19.628 (ver CLAUDE.md del proyecto).
"""

import os
import sys
import glob
import subprocess
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from estilo_maestro import inject_css, hero, guard_badge
from utils_aa import UMBRAL_DIAS_STALE

st.set_page_config(page_title="Recetas Cheque — Maestro AA", page_icon="🧾",
                    layout="wide", initial_sidebar_state="expanded")
inject_css()
hero("🧾", "Recetas Cheque", "Registro ISP — estupefacientes y psicotrópicos (obligación legal)")

try:
    from recetas_cheque import (listar_formularios, leer_folios_existentes,
                                 leer_periodo_form, autodescubrir_csv,
                                 RCH_DIR, MAESTRO_DIR, HOJA_RCH, MESES_ES,
                                 PREFIJO_CSV)
except Exception as e:
    st.error(f"No se pudo cargar recetas_cheque.py: {e}")
    st.stop()

_MES_NOMBRE = {v: k.capitalize() for k, v in MESES_ES.items()}


@st.cache_data(ttl=300, show_spinner=False)
def _fecha_max_sabana(path_mtime):
    path, _ = path_mtime
    try:
        col = pd.read_csv(path, encoding="latin-1", sep=None, engine="python",
                           usecols=["Fecha Entrega Receta"])
        m = pd.to_datetime(col["Fecha Entrega Receta"], dayfirst=True, errors="coerce").max()
        return m.date() if pd.notna(m) else None
    except Exception:
        return None


with st.sidebar:
    st.markdown("## 🧾 Recetas Cheque")
    st.caption("Agrega SOLO los folios nuevos de la sábana más reciente. No toca filas existentes.")
    st.markdown("---")
    try:
        csv_actual = autodescubrir_csv(MAESTRO_DIR)
        st.success(f"Sábana detectada:\n`{os.path.basename(csv_actual)}`")
        st.markdown("### 🛡️ Guard de frescura")
        st.caption(f"Si la sábana tiene más de **{UMBRAL_DIAS_STALE} días**, la actualización se bloquea "
                   f"en vez de reportar en silencio \"0 folios nuevos\" con datos viejos.")
        fecha_sab = _fecha_max_sabana((csv_actual, os.path.getmtime(csv_actual)))
        st.markdown(guard_badge(fecha_sab, UMBRAL_DIAS_STALE, "Sábana"), unsafe_allow_html=True)
    except FileNotFoundError:
        st.warning(f"No hay sábana `{PREFIJO_CSV}*.csv` en la carpeta del proyecto.")
    st.markdown("---")
    st.markdown("### Actualizar")
    if st.button("⚙️ Actualizar formularios ahora", type="primary", use_container_width=True):
        cmd = ["py", "recetas_cheque.py", "--no-pause"]
        with st.spinner("Procesando sábana y actualizando formularios ISP..."):
            res = subprocess.run(cmd, cwd=_ROOT, capture_output=True, text=True,
                                  encoding="utf-8", errors="replace")
        salida = res.stdout or ""
        if res.returncode == 0:
            st.success("Formularios actualizados.")
            with st.expander("Detalle", expanded=True):
                st.code(salida or "(sin salida)")
            st.cache_data.clear()
            st.rerun()
        elif "desactualizad" in salida:
            st.warning("🛡️ **Bloqueado por el guard de frescura** — la sábana auto-detectada es vieja, "
                       "no se procesó para evitar dar por al día un registro ISP desactualizado.")
            st.code(salida)
            st.caption("Ejecuta AUTO_SSASUR.bat para refrescar la sábana de recetas.")
        else:
            st.error("Falló la actualización.")
            st.code(res.stderr or salida or "(sin salida)")
    st.caption(f"Carpeta RCh: `{RCH_DIR}`")

if not os.path.isdir(RCH_DIR):
    st.warning(f"La carpeta de formularios ISP no existe en este equipo:\n\n`{RCH_DIR}`")
    st.stop()

formularios = listar_formularios(RCH_DIR)
if not formularios:
    st.info("No se encontraron formularios ISP con periodo declarado (B7/B8) en la carpeta RCh.")
    st.stop()

filas = []
for path, anio, mes in sorted(formularios, key=lambda t: (t[1], t[2]), reverse=True):
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        n_folios = len(leer_folios_existentes(wb[HOJA_RCH]))
        wb.close()
    except Exception:
        n_folios = None
    mtime = datetime.fromtimestamp(os.path.getmtime(path))
    filas.append({
        "Periodo": f"{_MES_NOMBRE.get(mes, mes)} {anio}",
        "Archivo": os.path.basename(path),
        "Folios registrados": n_folios if n_folios is not None else "—",
        "Última modificación": mtime.strftime("%d/%m/%Y %H:%M"),
        "_path": path,
    })

df = pd.DataFrame(filas)

hoy = datetime.now()
c1, c2, c3 = st.columns(3)
c1.metric("Formularios en carpeta RCh", f"{len(df)}")
mes_actual = any(f["Periodo"].lower().startswith(_MES_NOMBRE.get(hoy.month, "").lower()) and
                  str(hoy.year) in f["Periodo"] for f in filas)
c2.metric("Mes en curso", "✅ Listo" if mes_actual else "⚠️ Falta")
total_folios = sum(f["Folios registrados"] for f in filas if isinstance(f["Folios registrados"], int))
c3.metric("Total folios (histórico)", f"{total_folios:,}")

st.markdown("---")
st.dataframe(df.drop(columns=["_path"]), use_container_width=True, hide_index=True,
             height=min(50 + len(df) * 35, 450))

st.markdown("### Descargar formulario")
nombres = df["Periodo"] + " — " + df["Archivo"]
idx = st.selectbox("Elegir formulario", options=range(len(df)), format_func=lambda i: nombres.iloc[i])
ruta_sel = df.iloc[idx]["_path"]
with open(ruta_sel, "rb") as fh:
    st.download_button("📥 Descargar Excel", data=fh.read(), file_name=os.path.basename(ruta_sel),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("Recuerda completar a mano RUN QF, DV QF, Nombre QF y Posología en el formulario tras cada actualización.")
