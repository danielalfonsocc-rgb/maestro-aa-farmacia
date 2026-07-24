"""Conteo de Controlados — app Streamlit (Hospital Pitrufquén).

Sirve el formulario de conteo de stock de medicamentos controlados
(conteo_controlados.html) DENTRO de Streamlit, inyectándole:

  · controlados_config.json          → lista de medicamentos + mapeo SSASUR
                                        (fuente única; ver ese archivo).
  · Conteo_Controlados/stock_sistema.json → stock del DÍA ANTERIOR, que baja
                                        AUTO_SSASUR del "Reporte Stock en Fecha".
                                        Pre-llena la columna "Sistema".

El HTML es el mismo mockup que ya validó la QF (tabs por farmacia, resumen,
tabla de conteo físico, historial, consolidado, exportar JSON/Excel/PDF). Esta
app solo lo envuelve para:
  1. correrlo vía Streamlit (mismo ecosistema que app_maestro.py),
  2. inyectar el stock del día anterior automáticamente cada día,
  3. mostrar el estado del stock en la barra lateral.

El historial se guarda en el navegador (localStorage), igual que el HTML suelto,
y NUNCA se toca la columna Acta/Vencimiento (la maneja la QF a mano). El botón
"Exportar JSON" del formulario genera el archivo que se comparte y se conecta
con las demás apps del Maestro AA.

Correr:  py -m streamlit run conteo_controlados_app.py --server.port 8503
O:       ABRIR_CONTEO_CONTROLADOS.bat
"""
import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import streamlit as st
import streamlit.components.v1 as components

MAESTRO_DIR = Path(__file__).parent
HTML_FILE   = MAESTRO_DIR / "conteo_controlados.html"
CONFIG_FILE = MAESTRO_DIR / "controlados_config.json"
STOCK_FILE  = MAESTRO_DIR / "Conteo_Controlados" / "stock_sistema.json"

st.set_page_config(page_title="Conteo de Controlados", page_icon="💊", layout="wide")

# Quitar el padding superior de Streamlit para que el header del HTML quede al ras.
st.markdown(
    "<style>.block-container{padding:0.4rem 0.6rem 0 0.6rem;max-width:1200px;}"
    "header[data-testid='stHeader']{height:0;}</style>",
    unsafe_allow_html=True,
)


def _cargar_json(path):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except FileNotFoundError:
        return None
    except Exception as e:  # JSON corrupto → avisar, no reventar
        st.sidebar.error(f"No se pudo leer {Path(path).name}: {e}")
        return None


config = _cargar_json(CONFIG_FILE) or {}
stock  = _cargar_json(STOCK_FILE)   # None si aún no se ha bajado el stock

# ── Barra lateral: estado del stock del sistema ────────────────────────────────
with st.sidebar:
    st.markdown("### 💊 Conteo de Controlados")
    st.caption("Hospital Pitrufquén — pauta de supervisión")
    st.divider()
    st.markdown("**Stock del sistema**")
    if stock and stock.get("porFarmacia"):
        st.success("Stock cargado")
        nombres = {f["id"]: f["nombre"] for f in config.get("farmacias", [])}
        for fid, entry in stock["porFarmacia"].items():
            if isinstance(entry, dict) and "stock" in entry:
                fecha, n, dia = entry.get("fecha", "?"), len(entry["stock"]), entry.get("dia")
            else:  # formato plano (compat)
                fecha, n, dia = "?", len(entry or {}), None
            etq = " · día actual" if dia == "actual" else " · día anterior" if dia == "anterior" else ""
            st.caption(f"**{nombres.get(fid, fid)}**: {fecha}{etq} — {n} ítems")
        gen = stock.get("generadoEn")
        if gen:
            st.caption(f"Descargado: {str(gen)[:16].replace('T', ' ')}")
        st.caption(
            "La columna **Sistema** se pre-llena con este stock. "
            "Urgencia no tiene bodega en SSASUR → se ingresa a mano."
        )
    else:
        st.warning(
            "Aún no hay stock del sistema.\n\n"
            "Se genera solo cuando corras **AUTO_SSASUR** "
            "(baja el *Reporte Stock en Fecha*: Cerrada = día actual, Abierta = día anterior)."
        )
    st.divider()
    st.caption(
        "El historial se guarda en **este navegador**. Para compartir un conteo "
        "con las otras apps, usa **📤 Exportar JSON** dentro del formulario."
    )
    if not config:
        st.error("Falta controlados_config.json — el formulario usará su lista de respaldo.")
    if not HTML_FILE.exists():
        st.error("No se encontró conteo_controlados.html en la carpeta del proyecto.")

# ── Inyectar config + stock y servir el HTML ───────────────────────────────────
if HTML_FILE.exists():
    html = HTML_FILE.read_text(encoding="utf-8")
    inyeccion = (
        "<script>\n"
        f"window.CONTROLADOS_CONFIG = {json.dumps(config, ensure_ascii=False)};\n"
        f"window.STOCK_SISTEMA_INYECTADO = "
        f"{json.dumps(stock, ensure_ascii=False) if stock else 'null'};\n"
        "</script>\n"
    )
    # La inyección va JUSTO tras <body>, antes del <script> principal, para que
    # window.CONTROLADOS_CONFIG ya exista cuando corran las const del HTML.
    html = html.replace("<body>", "<body>\n" + inyeccion, 1)
    # OJO: se usa components.html (iframe srcdoc = MISMO origen que la app) a
    # propósito, NO st.iframe(src=...) ni st.html: el srcdoc es lo que permite
    # que el HTML corra su JS Y que localStorage persista el historial. st.iframe
    # solo acepta una URL y st.html no ejecuta <script>, así que ninguno sirve
    # para este formulario (probado en vivo: localStorage OK dentro del srcdoc).
    components.html(html, height=1500, scrolling=True)
else:
    st.stop()
