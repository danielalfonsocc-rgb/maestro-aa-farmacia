"""Sistema de diseño compartido para el hub Maestro AA (app_maestro.py y paginas/*).

Reutiliza la paleta de aa_colors.py y el lenguaje visual ya validado en
app_pedidos.py (lienzo gris-50, tarjetas blancas, acento teal, sombras
difusas) para que el hub y cada modulo se sientan parte de la misma app.
"""

import streamlit as st
from datetime import datetime

TEAL_GRAD = "linear-gradient(90deg,#0F766E 0%,#0E7490 100%)"


def inject_css():
    st.markdown("""
    <style>
    html, body, [class*="css"] {
        font-family:-apple-system,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;
        color:#1F2937;
    }
    .stApp { background:#F9FAFB; }
    .block-container { padding-top:1.4rem; padding-bottom:3rem; max-width:1320px; }
    h1,h2,h3,h4 { color:#111827; }
    h2 { font-weight:800; letter-spacing:-.01em; }
    hr { border:0; border-top:1px solid #EEF0F4; margin:1.4rem 0; }
    a { color:#0F766E; }

    /* ── Hero de marca ─────────────────────────────────────────── */
    .app-hero {
        display:flex; align-items:center; gap:18px;
        background:""" + TEAL_GRAD + """;
        border-radius:18px; padding:22px 28px; margin-bottom:22px; color:#fff;
        box-shadow:0 10px 30px -14px rgba(15,118,110,.55);
    }
    .app-hero .ico {
        width:56px; height:56px; border-radius:16px; background:rgba(255,255,255,.18);
        display:flex; align-items:center; justify-content:center; font-size:28px; flex:none;
    }
    .app-hero .tit { font-size:1.65rem; font-weight:800; line-height:1.1; margin:0; color:#fff; }
    .app-hero .sub { font-size:.92rem; opacity:.93; margin:4px 0 0; }
    .app-hero .chip {
        margin-left:auto; text-align:right; flex:none;
        background:rgba(255,255,255,.16); border-radius:12px; padding:9px 15px;
    }
    .app-hero .chip .d { font-size:.78rem; opacity:.9; }
    .app-hero .chip .w { font-size:.95rem; font-weight:700; }

    /* ── Grid de tarjetas de modulo ────────────────────────────── */
    .mod-grid {
        display:grid; grid-template-columns:repeat(auto-fit, minmax(270px, 1fr));
        gap:18px; margin:6px 0 26px;
    }
    .mod-card {
        display:block; background:#fff; border:1px solid #EEF0F4; border-radius:18px;
        padding:22px 24px; text-decoration:none !important; color:inherit;
        box-shadow:0 1px 2px rgba(16,24,40,.04), 0 10px 24px -16px rgba(16,24,40,.18);
        border-top:4px solid #CBD5E1; transition:transform .12s ease, box-shadow .12s ease;
        height:100%;
    }
    .mod-card:hover {
        transform:translateY(-2px);
        box-shadow:0 1px 2px rgba(16,24,40,.06), 0 16px 32px -16px rgba(16,24,40,.28);
    }
    .mod-card .mc-top { display:flex; align-items:flex-start; justify-content:space-between; gap:10px; }
    .mod-card .mc-ico {
        width:44px; height:44px; border-radius:13px; display:flex; align-items:center;
        justify-content:center; font-size:22px; flex:none; background:#F0FDFA;
    }
    .mod-card .mc-tit { font-size:1.08rem; font-weight:800; color:#111827; margin:12px 0 4px; }
    .mod-card .mc-sub { font-size:.85rem; color:#6B7280; line-height:1.45; min-height:2.6em; }
    .mod-card .mc-status {
        margin-top:14px; display:inline-flex; align-items:center; gap:6px;
        font-size:.76rem; font-weight:700; border-radius:999px; padding:4px 11px;
    }
    .st-ok    { background:#F0FDF4; color:#15803D; }
    .st-warn  { background:#FFFBEB; color:#B45309; }
    .st-crit  { background:#FEF2F2; color:#B91C1C; }
    .st-idle  { background:#F3F4F6; color:#6B7280; }

    .c-teal    { border-top-color:#0F766E; } .c-teal .mc-ico    { background:#F0FDFA; }
    .c-indigo  { border-top-color:#4F46E5; } .c-indigo .mc-ico  { background:#EEF2FF; }
    .c-rose    { border-top-color:#E11D48; } .c-rose .mc-ico    { background:#FFF1F2; }
    .c-amber   { border-top-color:#D97706; } .c-amber .mc-ico   { background:#FFFBEB; }
    .c-sky     { border-top-color:#0369A1; } .c-sky .mc-ico     { background:#F0F9FF; }

    /* ── Metricas / botones / tablas (mismo lenguaje que app_pedidos) ── */
    div[data-testid="stMetric"], div[data-testid="metric-container"] {
        background:#fff; border:1px solid #EEF0F4; border-radius:16px; padding:18px 20px;
        box-shadow:0 1px 2px rgba(16,24,40,.04), 0 8px 20px -16px rgba(16,24,40,.16);
    }
    div[data-testid="stMetricLabel"] p { color:#6B7280; font-weight:600; }
    .stButton > button, div[data-testid="stDownloadButton"] > button {
        border-radius:10px; font-weight:600; padding:.55rem 1.1rem;
        border:1px solid #0F766E; background:#0F766E; color:#fff;
        box-shadow:0 1px 2px rgba(16,24,40,.06); transition:filter .12s ease;
    }
    .stButton > button:hover, div[data-testid="stDownloadButton"] > button:hover {
        filter:brightness(1.08); border-color:#0F766E; color:#fff;
    }
    div[data-testid="stDataFrame"] {
        border:1px solid #EEF0F4; border-radius:14px; overflow:hidden;
        box-shadow:0 1px 2px rgba(16,24,40,.04);
    }
    div[data-testid="stExpander"] {
        border:1px solid #EEF0F4; border-radius:14px; background:#fff;
        box-shadow:0 1px 2px rgba(16,24,40,.04);
    }
    section[data-testid="stSidebar"] .stPageLink-NavLink { font-weight:600; }
    </style>
    """, unsafe_allow_html=True)


def hero(icono, titulo, subtitulo, chip_top=None, chip_bottom=None):
    chip_html = ""
    if chip_top or chip_bottom:
        chip_html = f"""<div class='chip'>
            <div class='d'>{chip_top or ''}</div>
            <div class='w'>{chip_bottom or ''}</div>
        </div>"""
    st.markdown(f"""
    <div class='app-hero'>
      <div class='ico'>{icono}</div>
      <div>
        <p class='tit'>{titulo}</p>
        <p class='sub'>{subtitulo}</p>
      </div>
      {chip_html}
    </div>
    """, unsafe_allow_html=True)


def fecha_relativa(mtime):
    """Texto humano de antigüedad + clase de estado (ok/warn/crit) para un mtime."""
    horas = (datetime.now() - mtime).total_seconds() / 3600
    fecha_str = mtime.strftime('%d/%m/%Y %H:%M')
    if horas < 4:
        return f"Al día · {fecha_str}", "st-ok"
    if horas < 24:
        return f"Hace {int(horas)} h · {fecha_str}", "st-warn"
    dias = int(horas // 24)
    return f"Hace {dias} día{'s' if dias > 1 else ''} · {fecha_str}", "st-crit"


def status_badge(texto, clase):
    return f"<span class='mc-status {clase}'>{texto}</span>"


def guard_badge(fecha, umbral_dias, etiqueta):
    """Badge para el guard de frescura compartido (ver utils_aa.verificar_frescura):
    ok si `fecha` está dentro de `umbral_dias`, crítico (bloquearía la generación)
    si los supera, e idle si no se pudo determinar la fecha."""
    if fecha is None:
        return status_badge(f"⚪ {etiqueta}: sin dato", "st-idle")
    dias = (datetime.now().date() - fecha).days
    if dias > umbral_dias:
        return status_badge(f"🔴 {etiqueta}: hace {dias} d — bloqueará la generación", "st-crit")
    return status_badge(f"🟢 {etiqueta}: hace {dias} d", "st-ok")
