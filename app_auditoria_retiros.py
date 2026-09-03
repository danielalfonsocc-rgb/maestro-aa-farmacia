"""Auditoría de Retiros por Medicamento — app Streamlit (Hospital Pitrufquén).

Responde la pregunta operativa: "¿cuántos pacientes distintos retiraron el
medicamento X en el periodo Y, y cuántas veces cada uno?" — sin tener que
pedirlo por chat cada vez.

Usa cargar_recetas_historico() (utils_aa.py), que combina el repo con las
carpetas de respaldo fuera del repo (RECETAS_HIST_DIRS) para tener toda la
profundidad histórica disponible, no solo la rotación reciente del repo.

Privacidad (Ley 19.628): la tabla con RUT/nombre por paciente está oculta por
defecto. Solo se muestra al marcar explícitamente la casilla correspondiente,
y el Excel exportable exige lo mismo.

Correr:  py -m streamlit run app_auditoria_retiros.py --server.port 8505
O:       AUDITORIA_RETIROS.bat
"""
import io
import os

import pandas as pd
import streamlit as st

from utils_aa import cargar_recetas_historico, setup_stdout

setup_stdout()

WORK = os.path.dirname(os.path.abspath(__file__))

COLS = [
    "ID Receta Detalle", "Prescripción", "RUN", "Nombre",
    "Apellido Paterno", "Apellido Materno",
    "Cantidad Recetada", "Cantidad Entregada",
    "Fecha Atención", "Fecha Entrega Receta", "Número Receta",
]

st.set_page_config(page_title="Auditoría de Retiros — Maestro AA", page_icon="💊", layout="wide")


@st.cache_data(show_spinner="Cargando histórico de recetas (puede tardar ~1 min)...")
def cargar_datos() -> pd.DataFrame:
    rec = cargar_recetas_historico(WORK, cols=COLS)
    rec["_ent"] = pd.to_numeric(rec["Cantidad Entregada"], errors="coerce").fillna(0)
    rec["_fent"] = pd.to_datetime(rec["Fecha Entrega Receta"], dayfirst=True, errors="coerce")
    rec["_fat"] = pd.to_datetime(rec["Fecha Atención"], dayfirst=True, errors="coerce")
    rec["_fecha"] = rec["_fent"].fillna(rec["_fat"])
    rec["_presc"] = rec["Prescripción"].fillna("").str.upper()
    return rec


st.title("💊 Auditoría de Retiros por Medicamento")
st.caption(
    "Cuenta pacientes distintos y N° de retiros efectivos (Cantidad Entregada > 0) "
    "de un medicamento en un periodo, usando todo el histórico de recetas disponible."
)

rec = cargar_datos()
fecha_min = rec["_fecha"].min()
fecha_max = rec["_fecha"].max()
st.caption(
    f"Histórico cargado: {len(rec):,} líneas de receta, del "
    f"{fecha_min:%d-%m-%Y} al {fecha_max:%d-%m-%Y}."
)

with st.form("busqueda"):
    col1, col2, col3, col4 = st.columns([2, 1, 1, 1])
    with col1:
        medicamento = st.text_input(
            "Medicamento (contiene, ej: LISDEXANFETAMINA)", ""
        ).strip().upper()
    with col2:
        dosis = st.text_input("Dosis (opcional, ej: 50 MG)", "").strip().upper()
    with col3:
        desde = st.date_input("Desde", value=fecha_min.date())
    with col4:
        hasta = st.date_input("Hasta", value=fecha_max.date())
    buscar = st.form_submit_button("Buscar", type="primary")

if buscar:
    if not medicamento:
        st.warning("Ingresa al menos el nombre (o parte del nombre) del medicamento.")
        st.stop()
    st.session_state["ultima_busqueda"] = (medicamento, dosis, desde, hasta)

if "ultima_busqueda" not in st.session_state:
    st.info("Completa el medicamento y el periodo, luego presiona Buscar.")
    st.stop()

medicamento, dosis, desde, hasta = st.session_state["ultima_busqueda"]

mask = rec["_presc"].str.contains(medicamento, regex=False)
if dosis:
    mask &= rec["_presc"].str.contains(dosis, regex=False)
mask &= rec["_fecha"].between(pd.Timestamp(desde), pd.Timestamp(hasta))
sub = rec[mask].copy()

if sub.empty:
    st.warning("Sin resultados para esos filtros.")
    st.stop()

st.subheader("Variantes de nombre encontradas")
st.dataframe(
    sub["Prescripción"].value_counts().rename_axis("Prescripción").reset_index(name="Líneas"),
    use_container_width=True, hide_index=True,
)

retiros = sub[sub["_ent"] > 0].copy()

c1, c2, c3 = st.columns(3)
c1.metric("Pacientes distintos (con retiro)", f"{retiros['RUN'].nunique():,}")
c2.metric("Retiros efectivos (líneas)", f"{len(retiros):,}")
c3.metric("Unidades entregadas", f"{int(retiros['_ent'].sum()):,}")

retiros["_nombre"] = (
    retiros["Nombre"].fillna("") + " " +
    retiros["Apellido Paterno"].fillna("") + " " +
    retiros["Apellido Materno"].fillna("")
).str.strip().str.upper()
retiros["Presentación"] = retiros["_presc"].str.extract(r"(\d+[\.,]?\d*\s*(?:MG|MCG|ML|G))")[0].fillna("(sin dosis detectada)")

detalle = (
    retiros.groupby(["RUN", "_nombre", "Presentación"])
    .agg(
        N_Retiros=("ID Receta Detalle", "count"),
        Unidades_Entregadas=("_ent", "sum"),
        N_Recetas_Distintas=("Número Receta", "nunique"),
        Primer_Retiro=("_fecha", "min"),
        Ultimo_Retiro=("_fecha", "max"),
    )
    .reset_index()
    .rename(columns={"_nombre": "Nombre"})
    .sort_values(["RUN", "Presentación"])
)

st.divider()
mostrar_rut = st.checkbox(
    "Mostrar RUT y nombres de pacientes (Ley 19.628 — solo uso interno)",
    value=False,
)
if mostrar_rut:
    st.subheader("Detalle por paciente")
    st.dataframe(detalle, use_container_width=True, hide_index=True)

    buf = io.BytesIO()
    detalle.to_excel(buf, index=False)
    st.download_button(
        "⬇️ Descargar Excel (con RUT)",
        data=buf.getvalue(),
        file_name=f"Retiros_{medicamento.replace(' ', '_')}_{desde}_{hasta}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.caption(f"{detalle.shape[0]} filas paciente-presentación ocultas. Marca la casilla para verlas.")
