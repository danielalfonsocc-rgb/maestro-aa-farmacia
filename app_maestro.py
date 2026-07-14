"""Maestro AA — Centro de Operaciones.

Hub de navegación que unifica los módulos de la Farmacia AT Abierta en una
sola app: Pedidos AA, Pedidos Fusionados, Centinela, Recetas Cheque y
Gestión Territorial. Cada módulo sigue siendo un script independiente
(se puede seguir abriendo por separado con su propio .bat); este hub solo
los agrupa bajo una navegación y un lenguaje visual común
(ver estilo_maestro.py).
"""

import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import streamlit as st

paginas = {
    "Centro de Operaciones": [
        st.Page("paginas/inicio.py", title="Inicio", icon="🏠", default=True),
    ],
    "Módulos": [
        st.Page("app_pedidos.py", title="Pedidos AA", icon="💊"),
        st.Page("paginas/pedidos_fusionados.py", title="Pedidos Fusionados", icon="🔗"),
        st.Page("paginas/centinela.py", title="Centinela", icon="🩺"),
        st.Page("paginas/recetas_cheque_page.py", title="Recetas Cheque", icon="🧾"),
        st.Page("paginas/gestion_territorial.py", title="Gestión Territorial", icon="🗺️"),
    ],
}

pg = st.navigation(paginas, position="top")
pg.run()
