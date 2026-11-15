import io
import json
import re
import time
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ==============================================================================
# ICONO VECTORIAL DE PESTAÑA DEL NAVEGADOR — IDENTIDAD MEGAMAN X / HIVIMAR S.A.
# ==============================================================================
PAGE_FAVICON_SVG = (
    "data:image/svg+xml;utf8,"
    "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'>"
    "<circle cx='32' cy='32' r='28' fill='%230057FF' stroke='%2300E5FF' stroke-width='5'/>"
    "<circle cx='32' cy='32' r='14' fill='%23FFFFFF' stroke='%23FF3B5C' stroke-width='4'/>"
    "<circle cx='32' cy='32' r='5' fill='%230057FF'/>"
    "</svg>"
)

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA — STOCK&YA! (HIVIMAR S.A.)
# ==============================================================================
st.set_page_config(
    page_title="STOCK&YA! | Inteligencia y Consolidación de Inventarios - HIVIMAR S.A.",
    page_icon=PAGE_FAVICON_SVG,
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==============================================================================
# PALETA DE IDENTIDAD OSCURA DE ALTO CONTRASTE + ACENTOS VIBRANTES
# ==============================================================================
PALETTE = {
    "bg_dark": "#061121",
    "card_dark": "#0D1E36",
    "sidebar_dark": "#050E1D",
    "white": "#FFFFFF",
    "text_soft": "#CBD5E1",
    "electric_blue": "#0057FF",
    "deep_blue": "#06193E",
    "cyan": "#00C8FF",
    "cyan_bright": "#00E5FF",
    "red": "#FF3B5C",
    "red_deep": "#E60026",
    "border_cyan": "rgba(0, 229, 255, 0.28)",
}

CORPORATE_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

:root {
    --x-bg: #061121;
    --x-card: #0D1E36;
    --x-sidebar: #050E1D;
    --x-white: #FFFFFF;
    --x-blue: #0057FF;
    --x-cyan: #00E5FF;
    --x-red: #FF3B5C;
    --x-muted: #CBD5E1;
}

html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    color: #FFFFFF;
}

/* Fondo general Oscuro Profundo para máximo contraste con letras blancas y colores vibrantes */
.stApp,
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
section.main {
    background:
        radial-gradient(circle at 92% 4%, rgba(0, 229, 255, 0.10) 0%, transparent 38%),
        radial-gradient(circle at 8% 95%, rgba(0, 87, 255, 0.14) 0%, transparent 42%),
        linear-gradient(180deg, #061121 0%, #08172E 55%, #050E1D 100%) !important;
    color: #FFFFFF !important;
}

/* Barra superior oscura */
[data-testid="stHeader"] {
    background-color: rgba(6, 17, 33, 0.92) !important;
    border-bottom: 1px solid rgba(0, 229, 255, 0.15) !important;
}

/* Barra lateral Armadura X (Azul Noche Profundo + Cian + Rojo) */
section[data-testid="stSidebar"],
section[data-testid="stSidebar"] > div:first-child {
    background: linear-gradient(180deg, #040B17 0%, #071630 55%, #050E1D 100%) !important;
    border-right: 2px solid rgba(0, 229, 255, 0.35) !important;
}

section[data-testid="stSidebar"] * {
    color: #FFFFFF !important;
}

section[data-testid="stSidebar"] .stRadio > div {
    gap: 0.45rem;
}

section[data-testid="stSidebar"] .stRadio label {
    background-color: #0D1E36 !important;
    border-radius: 10px !important;
    padding: 0.65rem 0.9rem !important;
    transition: all 0.2s ease;
    border: 1px solid rgba(0, 229, 255, 0.22) !important;
    color: #FFFFFF !important;
    font-weight: 600 !important;
}

section[data-testid="stSidebar"] .stRadio label:hover {
    background: linear-gradient(90deg, rgba(0, 87, 255, 0.65), rgba(0, 200, 255, 0.30)) !important;
    border-color: #00E5FF !important;
    transform: translateX(2px);
}

section[data-testid="stSidebar"] .stButton > button {
    background: linear-gradient(90deg, #0057FF 0%, #00B4D8 100%) !important;
    color: #FFFFFF !important;
    border: 1px solid #00E5FF !important;
    box-shadow: 0 4px 16px rgba(0, 180, 216, 0.35) !important;
}

/* Legibilidad garantizada para todos los labels, radios, selectboxes e inputs en modo oscuro */
[data-testid="stMain"] [data-testid="stWidgetLabel"] p,
[data-testid="stMain"] [data-testid="stWidgetLabel"] label,
[data-testid="stMain"] [data-testid="stRadio"] label p,
[data-testid="stMain"] [data-testid="stRadio"] label span,
[data-testid="stMain"] [data-testid="stCheckbox"] label span,
[data-testid="stMain"] label,
[data-testid="stMain"] p,
[data-testid="stMain"] .stMarkdown {
    color: #F8FAFC !important;
}

/* Inputs, Selectbox y NumberInput en oscuro */
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input,
[data-baseweb="select"] > div {
    background-color: #0D1E36 !important;
    color: #FFFFFF !important;
    border: 1px solid rgba(0, 229, 255, 0.35) !important;
    border-radius: 10px !important;
}
[data-baseweb="select"] span,
[data-baseweb="select"] div {
    color: #FFFFFF !important;
}
ul[data-baseweb="menu"] {
    background-color: #0D1E36 !important;
    border: 1px solid #00E5FF !important;
}
ul[data-baseweb="menu"] li {
    color: #FFFFFF !important;
}
ul[data-baseweb="menu"] li:hover {
    background-color: #0057FF !important;
}

/* Zona de carga de archivos (st.file_uploader) en oscuro con borde cian brillante */
[data-testid="stFileUploader"] section {
    background-color: #0D1E36 !important;
    border: 1.5px dashed #00E5FF !important;
    border-radius: 12px !important;
    padding: 16px !important;
}
[data-testid="stFileUploader"] section:hover {
    background-color: #122847 !important;
    border-color: #FFFFFF !important;
}
[data-testid="stFileUploader"] section span,
[data-testid="stFileUploader"] section small,
[data-testid="stFileUploader"] [data-testid="stFileUploaderFileName"] {
    color: #E2E8F0 !important;
}
[data-testid="stFileUploader"] button {
    background: linear-gradient(90deg, #0057FF 0%, #00A8E8 100%) !important;
    color: #FFFFFF !important;
    border: 1px solid #00E5FF !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
}

/* Expander en oscuro */
[data-testid="stExpander"] {
    background-color: #0D1E36 !important;
    border: 1px solid rgba(0, 229, 255, 0.25) !important;
    border-radius: 12px !important;
}
[data-testid="stExpander"] summary,
[data-testid="stExpander"] summary span,
[data-testid="stExpander"] summary p {
    color: #FFFFFF !important;
    font-weight: 700 !important;
}

/* Encabezados de sección en tarjeta oscura con borde vibrante */
.sec-header {
    display: flex;
    align-items: center;
    gap: 0.85rem;
    margin: 0.6rem 0 1.1rem 0;
    padding: 0.8rem 1rem;
    background: #0D1E36;
    border-left: 5px solid #00E5FF;
    border-top: 1px solid rgba(0, 229, 255, 0.22);
    border-right: 1px solid rgba(0, 229, 255, 0.22);
    border-bottom: 1px solid rgba(0, 229, 255, 0.22);
    border-radius: 12px;
    box-shadow: 0 6px 20px rgba(0, 0, 0, 0.35);
}
.sec-icon-wrap {
    width: 42px;
    height: 42px;
    border-radius: 10px;
    background: linear-gradient(135deg, #0057FF 0%, #00C8FF 100%);
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
    box-shadow: 0 4px 14px rgba(0, 200, 255, 0.35);
}
.sec-title {
    font-size: 1.22rem;
    font-weight: 800;
    color: #FFFFFF;
    letter-spacing: -0.01em;
    margin: 0;
}
.sec-sub {
    font-size: 0.84rem;
    color: #CBD5E1;
    margin: 2px 0 0 0;
}

/* Tarjeta de Fuente 01 y 02 en modo oscuro */
.source-header {
    display: flex;
    align-items: center;
    gap: 0.7rem;
    margin-bottom: 0.65rem;
    padding: 0.65rem 0.85rem;
    background: #0D1E36;
    border-radius: 12px;
    border: 1px solid rgba(0, 229, 255, 0.28);
    box-shadow: 0 4px 14px rgba(0, 0, 0, 0.28);
}
.num-badge {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 32px;
    height: 32px;
    border-radius: 8px;
    background: linear-gradient(135deg, #0057FF 0%, #00C8FF 100%);
    color: #FFFFFF;
    font-size: 0.86rem;
    font-weight: 800;
    letter-spacing: 0.04em;
    flex-shrink: 0;
}
.num-badge-red {
    background: linear-gradient(135deg, #E60026 0%, #FF3B5C 100%);
}
.source-title-text {
    font-size: 1rem;
    font-weight: 800;
    color: #FFFFFF;
}

/* Tarjetas KPI Ejecutivas Oscuras con Borde Superior Vibrante */
.kpi-card {
    background: #0D1E36;
    border: 1px solid rgba(0, 229, 255, 0.25);
    border-top: 4px solid #0057FF;
    border-radius: 14px;
    padding: 1.15rem 1.25rem;
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.38);
    transition: transform 0.2s ease, border-color 0.2s ease;
    height: 100%;
}
.kpi-card.cyan-top { border-top-color: #00E5FF; }
.kpi-card.red-top { border-top-color: #FF3B5C; }
.kpi-card:hover {
    transform: translateY(-2px);
    border-color: #00E5FF;
}
.kpi-label {
    font-size: 0.76rem;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    color: #94A3B8;
    font-weight: 800;
    margin-bottom: 0.35rem;
}
.kpi-value {
    font-size: 1.75rem;
    font-weight: 800;
    color: #FFFFFF;
    line-height: 1.15;
}
.kpi-sub {
    font-size: 0.82rem;
    margin-top: 0.42rem;
    font-weight: 700;
}
.badge-ok { color: #38BDF8; }
.badge-cyan { color: #00E5FF; }
.badge-warn { color: #FBBF24; }
.badge-err { color: #FF5C77; }

/* Banner de Bienvenida Vibrante sobre Fondo Oscuro */
.hero-banner {
    background: linear-gradient(115deg, #051638 0%, #0057FF 56%, #00A8E8 100%);
    border: 1.5px solid #00E5FF;
    border-radius: 18px;
    padding: 1.45rem 2rem;
    margin-bottom: 1.25rem;
    box-shadow: 0 12px 30px rgba(0, 87, 255, 0.32);
    position: relative;
    overflow: hidden;
}
.hero-banner::after {
    content: "";
    position: absolute;
    top: 0;
    right: 0;
    width: 120px;
    height: 6px;
    background: #FF3B5C;
}

/* Alerta ¡AGUANTA! en Oscuro con Acento Rojo Coral */
.aguanta-box {
    background: linear-gradient(135deg, #1F0B16 0%, #2D0F20 100%);
    border: 1.5px solid #FF3B5C;
    border-left: 8px solid #FF3B5C;
    border-radius: 14px;
    padding: 1.15rem 1.4rem;
    margin: 0.85rem 0 1rem 0;
    display: flex;
    align-items: center;
    gap: 1rem;
    box-shadow: 0 8px 24px rgba(255, 59, 92, 0.22);
}

/* Botones principales */
.stButton > button {
    background: linear-gradient(90deg, #0057FF 0%, #00A8E8 100%) !important;
    color: #FFFFFF !important;
    border: 1px solid #00E5FF !important;
    border-radius: 10px !important;
    font-weight: 800 !important;
    letter-spacing: 0.02em;
    transition: all 0.2s ease !important;
}
.stButton > button:hover {
    background: linear-gradient(90deg, #0044CC 0%, #00C8FF 100%) !important;
    box-shadow: 0 6px 20px rgba(0, 229, 255, 0.35) !important;
}

/* Botón maestro Descargar Todo &YA! */
div[data-testid="stDownloadButton"] > button {
    background: linear-gradient(90deg, #E60026 0%, #0057FF 100%) !important;
    color: #FFFFFF !important;
    border: 1.5px solid #00E5FF !important;
    border-radius: 12px !important;
    font-weight: 800 !important;
    font-size: 0.94rem !important;
    padding: 0.6rem 1.2rem !important;
    box-shadow: 0 6px 20px rgba(255, 59, 92, 0.32) !important;
}
div[data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 8px 26px rgba(0, 229, 255, 0.45) !important;
}

/* Pestañas en modo oscuro */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    background-color: #0D1E36;
    padding: 6px;
    border-radius: 12px;
    border: 1px solid rgba(0, 229, 255, 0.25);
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px;
    color: #CBD5E1 !important;
    font-weight: 700;
    padding: 8px 16px;
}
.stTabs [aria-selected="true"] {
    background-color: #0057FF !important;
    color: #FFFFFF !important;
}

/* Pie de página oscuro */
.copyright-footer {
    margin-top: 3rem;
    padding: 0.95rem 1rem;
    border-top: 2px solid #00A8E8;
    background: #0D1E36;
    border-radius: 10px;
    text-align: center;
    font-size: 0.78rem;
    color: #CBD5E1;
    font-weight: 600;
    letter-spacing: 0.03em;
}
</style>
"""

# ==============================================================================
# ICONOS VECTORIALES PROFESIONALES (SVG — BLANCO, CIAN, AZUL ELÉCTRICO, ROJO)
# ==============================================================================
HIVIMAR_LOGO_SVG = """
<svg width="78" height="78" viewBox="0 0 140 140" fill="none" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="archGrad" x1="15" y1="20" x2="125" y2="85" gradientUnits="userSpaceOnUse">
      <stop offset="0%" stop-color="#00E5FF"/>
      <stop offset="50%" stop-color="#FFFFFF"/>
      <stop offset="100%" stop-color="#FF3B5C"/>
    </linearGradient>
    <linearGradient id="steelGrad" x1="30" y1="30" x2="110" y2="110" gradientUnits="userSpaceOnUse">
      <stop offset="0%" stop-color="#FFFFFF"/>
      <stop offset="55%" stop-color="#BAE6FD"/>
      <stop offset="100%" stop-color="#0057FF"/>
    </linearGradient>
  </defs>
  <path d="M 20 82 A 52 52 0 0 1 120 82" stroke="url(#archGrad)" stroke-width="7" stroke-linecap="round"/>
  <path d="M 30 82 A 42 42 0 0 1 110 82" stroke="#00E5FF" stroke-width="2.2" stroke-dasharray="5 4" opacity="0.85"/>
  <circle cx="20" cy="82" r="4.5" fill="#00E5FF"/>
  <circle cx="120" cy="82" r="4.5" fill="#FF3B5C"/>
  <circle cx="70" cy="30" r="5" fill="#FF3B5C" stroke="#FFFFFF" stroke-width="1.5"/>

  <g transform="rotate(-45 70 74)">
    <rect x="64" y="24" width="12" height="96" rx="3" fill="url(#steelGrad)"/>
    <rect x="55" y="18" width="30" height="16" rx="4" fill="#FFFFFF" stroke="#00E5FF" stroke-width="2.2"/>
    <line x1="56" y1="23" x2="84" y2="23" stroke="#0057FF" stroke-width="2"/>
    <line x1="56" y1="28" x2="84" y2="28" stroke="#FF3B5C" stroke-width="2"/>
    <circle cx="70" cy="116" r="8" fill="#0057FF" stroke="#FFFFFF" stroke-width="2"/>
  </g>

  <g transform="rotate(45 70 74)">
    <rect x="64" y="24" width="12" height="96" rx="3" fill="url(#steelGrad)"/>
    <rect x="55" y="18" width="30" height="16" rx="4" fill="#FFFFFF" stroke="#FF3B5C" stroke-width="2.2"/>
    <line x1="56" y1="23" x2="84" y2="23" stroke="#0057FF" stroke-width="2"/>
    <line x1="56" y1="28" x2="84" y2="28" stroke="#FF3B5C" stroke-width="2"/>
    <circle cx="70" cy="116" r="8" fill="#FF3B5C" stroke="#FFFFFF" stroke-width="2"/>
  </g>

  <circle cx="70" cy="74" r="26" fill="#06193E" stroke="#FFFFFF" stroke-width="5"/>
  <circle cx="70" cy="74" r="14" fill="#0057FF" stroke="#00E5FF" stroke-width="3.5"/>
  <circle cx="70" cy="74" r="6" fill="#FF3B5C" stroke="#FFFFFF" stroke-width="1.5"/>
</svg>
"""

SVG_ICONS = {
    "upload": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>""",
    "audit": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v18"/><path d="M5 7h14"/><path d="M5 7l-3 7h6l-3-7z"/><path d="M19 7l-3 7h6l-3-7z"/><line x1="8" y1="21" x2="16" y2="21"/></svg>""",
    "ranking": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>""",
    "trends": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>""",
    "ai": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="16" height="16" rx="2"/><rect x="9" y="9" width="6" height="6"/><line x1="9" y1="1" x2="9" y2="4"/><line x1="15" y1="1" x2="15" y2="4"/><line x1="9" y1="20" x2="9" y2="23"/><line x1="15" y1="20" x2="15" y2="23"/><line x1="20" y1="9" x2="23" y2="9"/><line x1="20" y1="14" x2="23" y2="14"/><line x1="1" y1="9" x2="4" y2="9"/><line x1="1" y1="14" x2="4" y2="14"/></svg>""",
    "table": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="9" x2="9" y2="21"/></svg>""",
    "alert": """<svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="#FF3B5C" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>""",
}


def apply_dark_plotly_layout(fig, title: str = "", height: int = 380, xaxis_title: str = "", yaxis_title: str = "", showlegend: bool = True):
    """Aplica estilo oscuro de alto contraste consistente a todos los gráficos Plotly."""
    fig.update_layout(
        title=dict(text=title, font=dict(color="#FFFFFF", size=15)),
        paper_bgcolor="#0D1E36",
        plot_bgcolor="#08162B",
        font=dict(color="#F8FAFC", family="Plus Jakarta Sans, sans-serif"),
        height=height,
        xaxis_title=xaxis_title,
        yaxis_title=yaxis_title,
        showlegend=showlegend,
        margin=dict(l=20, r=30, t=52, b=24),
        xaxis=dict(gridcolor="rgba(148, 163, 184, 0.14)", zerolinecolor="rgba(148, 163, 184, 0.25)"),
        yaxis=dict(gridcolor="rgba(148, 163, 184, 0.14)", zerolinecolor="rgba(148, 163, 184, 0.25)"),
    )
    return fig


def render_section_header(icon_key: str, title: str, subtitle: str = ""):
    icon_svg = SVG_ICONS.get(icon_key, SVG_ICONS["audit"])
    sub_html = f'<div class="sec-sub">{subtitle}</div>' if subtitle else ""
    st.markdown(
        f"""
        <div class="sec-header">
          <div class="sec-icon-wrap">{icon_svg}</div>
          <div>
            <div class="sec-title">{title}</div>
            {sub_html}
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_source_badge_header(number_str: str, title: str, is_second: bool = False):
    badge_cls = "num-badge num-badge-red" if is_second else "num-badge"
    st.markdown(
        f"""
        <div class="source-header">
          <span class="{badge_cls}">{number_str}</span>
          <span class="source-title-text">{title}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def build_powerbi_theme_json() -> bytes:
    """Genera archivo de tema oficial JSON para Power BI con la identidad Mega Man X / HIVIMAR S.A."""
    theme = {
        "name": "HIVIMAR_STOCK_Y_YA_MEGAMAN_X_DARK",
        "dataColors": [
            "#0057FF",
            "#00E5FF",
            "#FF3B5C",
            "#38BDF8",
            "#0096C7",
            "#FF6B81",
            "#0284C7",
            "#1E40AF",
        ],
        "background": "#061121",
        "foreground": "#FFFFFF",
        "tableAccent": "#00E5FF",
        "good": "#00E5FF",
        "neutral": "#0057FF",
        "bad": "#FF3B5C",
        "maximum": "#0057FF",
        "center": "#00E5FF",
        "minimum": "#FF3B5C",
        "visualStyles": {
            "*": {
                "*": {
                    "background": [{"color": {"solid": {"color": "#0D1E36"}}}],
                    "border": [
                        {
                            "show": True,
                            "color": {"solid": {"color": "#00E5FF"}},
                            "radius": 8,
                        }
                    ],
                    "title": [
                        {
                            "show": True,
                            "fontColor": {"solid": {"color": "#FFFFFF"}},
                            "fontFamily": "Segoe UI Semibold",
                        }
                    ],
                }
            }
        },
    }
    return json.dumps(theme, indent=2, ensure_ascii=False).encode("utf-8")


# ==============================================================================
# DETECCIÓN AUTOMÁTICA DE RUBRO SEGÚN NOMBRE DE ARCHIVO Y CONTENIDO
# ==============================================================================
def detect_industry_sector(df_list: list, filenames: list = None) -> tuple:
    file_corpus = " ".join([str(fn) for fn in (filenames or []) if fn]).lower()

    data_texts = []
    for df in df_list:
        if df is not None and not df.empty:
            data_texts.extend([str(c) for c in df.columns])
            for col in df.columns:
                if df[col].dtype == object:
                    data_texts.extend(df[col].dropna().astype(str).head(200).tolist())
    data_corpus = " ".join(data_texts).lower()

    sectors = [
        (
            "agro_alimentos",
            "Agroindustrial y Alimenticio",
            [
                "papa", "papas", "arroz", "maiz", "maíz", "trigo", "fruta", "verdura",
                "cacao", "banano", "harina", "aceite vegetal", "leche", "lacteo", "lácteo",
                "carne", "pollo", "atun", "atún", "semilla", "abono", "fertilizante",
                "agro", "alimento", "alimentos", "cosecha", "grano", "azucar", "azúcar",
                "cafe", "café", "tuberculo", "tubérculo", "hortaliza"
            ],
        ),
        (
            "farmaceutico",
            "Farmacéutico y Hospitalario",
            [
                "pastilla", "pastillas", "tableta", "capsula", "cápsula", "jarabe",
                "ampolla", "jeringa", "suero", "gasa", "guante", "quirurgico", "quirúrgico",
                "hospital", "medico", "médico", "farmacia", "farmaceutico", "farmacéutico",
                "paracetamol", "ibuprofeno", "antibiotico", "antibiótico", "cateter",
                "catéter", "insumo medico", "medicamento", "medicamentos"
            ],
        ),
        (
            "dulceria",
            "Dulcería, Confitería y Snacks",
            [
                "caramelo", "caramelos", "chocolate", "chocolates", "galleta", "galletas",
                "dulce", "dulces", "chicle", "gomita", "gomitas", "bombon", "bombón",
                "chupete", "wafer", "turron", "turrón", "snack", "snacks", "confite",
                "confiteria", "confitería", "pastel", "helado", "refresco", "gaseosa"
            ],
        ),
        (
            "ferreteria",
            "Ferretería y Construcción",
            [
                "cemento", "tubo", "tuberia", "pvc", "cable", "alambre", "perno",
                "tuerca", "tornillo", "taladro", "pintura", "brocha", "valvula", "válvula",
                "llave", "martillo", "hierro", "acero", "ladrillo", "ceramica", "ferreteria",
                "ferretería", "construccion", "construcción"
            ],
        ),
        (
            "automotriz",
            "Automotriz e Industrial (HIVIMAR)",
            [
                "skf", "ntn", "koyo", "timken", "rodamiento", "ruliman", "rulimán",
                "aceite", "filtro", "sello", "reten", "retén", "piston", "pistón",
                "freno", "embrague", "bujia", "bujía", "correa", "amortiguador",
                "prestone", "acdelco", "grasa", "industrial", "automotriz", "hivimar",
                "repuesto", "repuestos"
            ],
        ),
    ]

    best_key, best_label, best_score = "automotriz", "Automotriz e Industrial (HIVIMAR)", 0
    for key, label, keywords in sectors:
        file_hits = sum(5 for kw in keywords if kw in file_corpus)
        content_hits = sum(data_corpus.count(kw) for kw in keywords)
        total_score = file_hits + content_hits
        if total_score > best_score:
            best_score = total_score
            best_key = key
            best_label = label

    return best_key, best_label


def render_dynamic_loader_animation(sector_key: str, sector_label: str, analyzed_files: list = None) -> str:
    item_svgs = {
        "agro_alimentos": """
            <g class="wheel-spin">
              <ellipse cx="0" cy="0" rx="22" ry="16" fill="#0057FF" stroke="#00E5FF" stroke-width="2.5"/>
              <circle cx="-8" cy="-4" r="2.2" fill="#FFFFFF"/>
              <circle cx="7" cy="5" r="2.2" fill="#00E5FF"/>
              <circle cx="4" cy="-6" r="2" fill="#FF3B5C"/>
            </g>
        """,
        "farmaceutico": """
            <g class="wheel-spin">
              <rect x="-22" y="-11" width="44" height="22" rx="11" fill="#FFFFFF" stroke="#0057FF" stroke-width="2.2"/>
              <path d="M 0 -11 L 11 -11 A 11 11 0 0 1 11 11 L 0 11 Z" fill="#FF3B5C"/>
              <line x1="0" y1="-11" x2="0" y2="11" stroke="#06193E" stroke-width="1.5"/>
            </g>
        """,
        "dulceria": """
            <g class="wheel-spin">
              <circle cx="0" cy="0" r="19" fill="#0057FF" stroke="#FFFFFF" stroke-width="3"/>
              <circle cx="0" cy="0" r="7" fill="#FF3B5C"/>
              <circle cx="-11" cy="-6" r="2.2" fill="#00E5FF"/>
              <circle cx="10" cy="-7" r="2.2" fill="#FFFFFF"/>
            </g>
        """,
        "ferreteria": """
            <g class="wheel-spin">
              <polygon points="0,-21 18,-10 18,10 0,21 -18,10 -18,-10" fill="#FFFFFF" stroke="#0057FF" stroke-width="2.8"/>
              <circle cx="0" cy="0" r="9" fill="#06193E" stroke="#00E5FF" stroke-width="2.5"/>
            </g>
        """,
        "automotriz": """
            <g class="wheel-spin">
              <circle cx="0" cy="0" r="20" fill="#06193E" stroke="#FFFFFF" stroke-width="3.5"/>
              <circle cx="0" cy="0" r="10" fill="#0057FF" stroke="#00E5FF" stroke-width="2.5"/>
              <circle cx="0" cy="-15" r="2.8" fill="#00E5FF"/>
              <circle cx="0" cy="15" r="2.8" fill="#00E5FF"/>
              <circle cx="-15" cy="0" r="2.8" fill="#FF3B5C"/>
              <circle cx="15" cy="0" r="2.8" fill="#FF3B5C"/>
              <circle cx="-10.6" cy="-10.6" r="2.6" fill="#FFFFFF"/>
              <circle cx="10.6" cy="10.6" r="2.6" fill="#FFFFFF"/>
            </g>
        """,
    }

    chosen_svg = item_svgs.get(sector_key, item_svgs["automotriz"])
    files_caption = ""
    if analyzed_files:
        clean_names = [str(x) for x in analyzed_files if x]
        if clean_names:
            files_caption = f"Archivos analizados: <strong>{' | '.join(clean_names)}</strong> &bull; "

    return f"""
    <div style="background: linear-gradient(135deg, #06193E 0%, #0057FF 100%);
                border: 2px solid #00E5FF; border-radius: 16px; padding: 1.4rem 1.8rem;
                margin: 1rem 0; text-align: center; box-shadow: 0 10px 28px rgba(0,87,255,0.35);">
      <style>
        @keyframes winkPhase {{
          0%, 100% {{ opacity: 1; transform: scale(1); }}
          45% {{ opacity: 1; transform: scale(1.06); }}
          55% {{ opacity: 0; transform: scale(0.85); }}
          90% {{ opacity: 0; }}
        }}
        @keyframes trackPhase {{
          0%, 48% {{ opacity: 0; transform: translateY(6px); }}
          58%, 100% {{ opacity: 1; transform: translateY(0); }}
        }}
        @keyframes wheelMove {{
          0% {{ transform: translate(35px, 34px); }}
          100% {{ transform: translate(365px, 34px); }}
        }}
        @keyframes wheelMove2 {{
          0% {{ transform: translate(-20px, 34px); }}
          100% {{ transform: translate(310px, 34px); }}
        }}
        @keyframes spinFast {{
          0% {{ transform: rotate(0deg); }}
          100% {{ transform: rotate(720deg); }}
        }}
        @keyframes dashTrack {{
          0% {{ stroke-dashoffset: 40; }}
          100% {{ stroke-dashoffset: 0; }}
        }}
        .wink-group {{ animation: winkPhase 2.6s infinite ease-in-out; transform-origin: center; }}
        .track-group {{ animation: trackPhase 2.6s infinite ease-in-out; }}
        .roller-1 {{ animation: wheelMove 1.3s infinite linear; }}
        .roller-2 {{ animation: wheelMove2 1.3s infinite linear; }}
        .wheel-spin {{ animation: spinFast 0.75s infinite linear; }}
        .track-line {{ stroke-dasharray: 12 8; animation: dashTrack 0.4s infinite linear; }}
      </style>

      <div style="font-size: 1.32rem; font-weight: 800; color: #FFFFFF; letter-spacing: 0.04em; margin-bottom: 0.3rem;">
        DAME CHANCE QUE ESTOY CARGANDO...
      </div>
      <div style="font-size: 0.86rem; color: #E0F2FE; margin-bottom: 0.75rem;">
        {files_caption}Rubro identificado automáticamente: <strong style="color:#00E5FF;">{sector_label}</strong>
      </div>

      <svg width="100%" height="74" viewBox="0 0 420 74" style="max-width: 520px; overflow: visible;">
        <g class="wink-group">
          <circle cx="210" cy="36" r="26" fill="#06193E" stroke="#00E5FF" stroke-width="3"/>
          <circle cx="201" cy="30" r="3.5" fill="#FFFFFF"/>
          <path d="M 215 30 Q 221 25 226 30" stroke="#00E5FF" stroke-width="3.2" stroke-linecap="round" fill="none"/>
          <path d="M 198 43 Q 210 53 222 43" stroke="#FFFFFF" stroke-width="3" stroke-linecap="round" fill="none"/>
        </g>

        <g class="track-group">
          <line x1="15" y1="58" x2="405" y2="58" stroke="#00E5FF" stroke-width="3" stroke-linecap="round"/>
          <line class="track-line" x1="15" y1="64" x2="405" y2="64" stroke="#FF3B5C" stroke-width="2.5"/>
          <g class="roller-1">{chosen_svg}</g>
          <g class="roller-2">{chosen_svg}</g>
        </g>
      </svg>
    </div>
    """


# ==============================================================================
# MOTOR AGNÓSTICO DE LECTURA UNIVERSAL (SOPORTA MÚLTIPLES MESES, LÍNEAS Y ERPs)
# ==============================================================================
def normalize_text(text) -> str:
    if text is None:
        return ""
    s = str(text).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^a-z0-9_ ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def sanitize_label(label: str, fallback: str = "FUENTE") -> str:
    if not label:
        return fallback
    cleaned = re.sub(r"[^A-Za-z0-9ÁÉÍÓÚÜÑáéíóúüñ _-]+", "", str(label)).strip()
    cleaned = re.sub(r"\s+", "_", cleaned).upper()
    return cleaned[:24] if cleaned else fallback


def parse_numeric_series(series: pd.Series) -> pd.Series:
    """Convierte columnas numéricas o textos con formato de miles/moneda a float sin errores."""
    if pd.api.types.is_numeric_dtype(series):
        return series.fillna(0.0).astype(float)
    cleaned = (
        series.astype(str)
        .str.replace(r"[^\d,.\-]", "", regex=True)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def derive_smart_source_label(filename: str, sheet_name: str, fallback: str) -> str:
    generic_sheets = {
        "hoja1", "hoja2", "hoja3", "sheet1", "sheet2", "sheet3",
        "datos", "datos_csv", "inventario", "reporte", "export", "tabla1", "base"
    }
    norm_sheet = normalize_text(sheet_name)
    base_file = filename.rsplit(".", 1)[0] if filename else ""
    if norm_sheet in generic_sheets and base_file:
        return sanitize_label(base_file, fallback)
    if base_file and norm_sheet not in generic_sheets:
        b_clean = sanitize_label(base_file, "")
        s_clean = sanitize_label(sheet_name, "")
        if b_clean and s_clean and b_clean != s_clean and len(b_clean) <= 12:
            return f"{b_clean}_{s_clean}"[:24]
        return s_clean or b_clean or fallback
    return sanitize_label(sheet_name or base_file, fallback)


def promote_real_header_row(df: pd.DataFrame) -> pd.DataFrame:
    """
    Detecta automáticamente si un reporte ERP (SAP, HIVIPRO, Oracle, Excel mensual)
    tiene filas de título o metadatos arriba antes de los encabezados reales.
    """
    if df is None or df.empty:
        return df

    header_keywords = {
        "codigo", "cod", "material", "sku", "producto", "descripcion", "texto",
        "stock", "cantidad", "existencia", "saldo", "libre", "conteo", "fisico",
        "lote", "ubicacion", "bodega", "linea", "categoria", "familia",
        "ingreso", "ingresos", "egreso", "egresos", "venta", "ventas", "precio", "costo"
    }

    current_cols_norm = [normalize_text(c) for c in df.columns]
    unnamed_ratio = sum(1 for c in current_cols_norm if not c or c.startswith("unnamed")) / max(len(current_cols_norm), 1)
    current_hits = sum(
        1 for c in current_cols_norm
        if any(kw in c.split() or (len(kw) >= 4 and kw in c) for kw in header_keywords)
    )

    if current_hits >= 2 and unnamed_ratio < 0.5:
        return df.dropna(how="all").reset_index(drop=True)

    best_row_idx = None
    best_score = 0
    scan_limit = min(15, len(df))
    for r_idx in range(scan_limit):
        row_vals = [normalize_text(v) for v in df.iloc[r_idx].tolist()]
        non_empty = [v for v in row_vals if v and v != "nan" and v != "none"]
        if len(non_empty) < 2:
            continue
        hits = sum(
            1 for v in non_empty
            if any(kw in v.split() or (len(kw) >= 4 and kw in v) for kw in header_keywords)
        )
        if hits > best_score and hits >= 2:
            best_score = hits
            best_row_idx = r_idx

    if best_row_idx is not None:
        raw_headers = df.iloc[best_row_idx].tolist()
        seen = {}
        clean_headers = []
        for idx, val in enumerate(raw_headers):
            col_str = str(val).strip() if pd.notna(val) and str(val).strip() != "" else f"Columna_{idx+1}"
            if col_str in seen:
                seen[col_str] += 1
                col_str = f"{col_str}_{seen[col_str]}"
            else:
                seen[col_str] = 1
            clean_headers.append(col_str)
        df_clean = df.iloc[best_row_idx + 1:].copy()
        df_clean.columns = clean_headers
        return df_clean.dropna(how="all").reset_index(drop=True)

    return df.dropna(how="all").reset_index(drop=True)


@st.cache_data(show_spinner=False, ttl=1800)
def parse_workbook_bytes(filename: str, raw_bytes: bytes) -> dict:
    """Lee y cachea libros Excel o CSV para evitar reprocesamiento innecesario."""
    fname = filename.lower()
    if fname.endswith(".csv"):
        for enc in ["utf-8", "latin-1", "cp1252"]:
            try:
                df = pd.read_csv(io.BytesIO(raw_bytes), encoding=enc, sep=None, engine="python")
                df = promote_real_header_row(df)
                return {"Datos_CSV": df} if df is not None and not df.empty else {}
            except Exception:
                continue
        df = pd.read_csv(io.BytesIO(raw_bytes), encoding="latin-1")
        df = promote_real_header_row(df)
        return {"Datos_CSV": df} if df is not None and not df.empty else {}

    xls = pd.ExcelFile(io.BytesIO(raw_bytes))
    sheets = {}
    for name in xls.sheet_names:
        df_tmp = promote_real_header_row(pd.read_excel(xls, sheet_name=name))
        if df_tmp is not None and not df_tmp.empty:
            sheets[name] = df_tmp
    return sheets


def read_uploaded_workbook(uploaded_file) -> dict:
    return parse_workbook_bytes(uploaded_file.name, uploaded_file.getvalue())


def load_google_sheet_from_url(url: str) -> dict:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not match:
        raise ValueError("Enlace de Google Sheets no válido. Pega la URL completa de la hoja compartida.")
    sheet_id = match.group(1)
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    xls = pd.ExcelFile(export_url)
    sheets = {}
    for name in xls.sheet_names:
        df_tmp = promote_real_header_row(pd.read_excel(xls, sheet_name=name))
        if df_tmp is not None and not df_tmp.empty:
            sheets[name] = df_tmp
    return sheets


def auto_detect_schema(df: pd.DataFrame) -> dict:
    cols = list(df.columns)
    norm_map = {c: normalize_text(c) for c in cols}

    patterns = {
        "codigo": [
            "codigo_sap", "cod_sap", "codigo", "material", "sku", "cod_mat",
            "id_producto", "item", "clave", "referencia", "cod", "numero_parte", "parte"
        ],
        "descripcion": [
            "producto", "texto breve de material", "texto breve", "descripcion",
            "nombre", "articulo", "detalle", "material_desc", "item_name", "glosa"
        ],
        "stock": [
            "stock", "librutiliz", "libre utilizacion", "conteo", "existencia",
            "cantidad", "saldo", "inventario", "fisico", "logico", "unidades", "cant", "total"
        ],
        "ubicacion": [
            "lote", "ubicacion", "bodega", "almacen", "rack", "pasillo",
            "percha", "estante", "localizacion", "zona", "sucursal"
        ],
        "categoria": [
            "linea", "categoria", "familia", "grupo", "rubro", "clase",
            "departamento", "tipo", "segmento", "marca"
        ],
        "ingresos": [
            "ingresos", "ingreso", "entradas", "entrada", "compras", "compra",
            "recepcion", "abastecimiento", "alta", "altas"
        ],
        "egresos": [
            "egresos", "egreso", "ventas", "venta", "salidas", "salida",
            "despacho", "consumo", "facturado", "vendido", "baja", "bajas", "rotacion"
        ],
        "precio": [
            "precio", "costo", "valor", "monto", "importe", "pvp", "costo_unit", "precio_unit"
        ],
    }

    detected = {k: None for k in patterns}

    for field, kw_list in patterns.items():
        best_col = None
        best_score = 0
        for orig_col, norm_col in norm_map.items():
            if not norm_col or norm_col.startswith("unnamed"):
                continue
            for idx, kw in enumerate(kw_list):
                if norm_col == kw:
                    score = 100 - idx
                elif kw in norm_col.split():
                    score = 80 - idx
                elif kw in norm_col and len(kw) >= 4:
                    score = 60 - idx
                else:
                    score = 0
                if score > best_score and orig_col not in detected.values():
                    best_score = score
                    best_col = orig_col
        detected[field] = best_col

    if detected["codigo"] is None:
        for c in cols:
            if c not in detected.values() and df[c].notna().sum() > 0:
                detected["codigo"] = c
                break

    if detected["descripcion"] is None:
        for c in cols:
            if c not in detected.values() and df[c].dtype == object:
                detected["descripcion"] = c
                break

    if detected["stock"] is None:
        numeric_candidates = [
            c for c in cols
            if c not in detected.values() and pd.api.types.is_numeric_dtype(df[c])
        ]
        if numeric_candidates:
            detected["stock"] = numeric_candidates[-1]
        else:
            for c in cols:
                if c not in detected.values():
                    coerced = pd.to_numeric(df[c], errors="coerce")
                    if coerced.notna().sum() > len(df) * 0.4:
                        detected["stock"] = c
                        break

    return detected


def consolidate_source(df: pd.DataFrame, schema: dict, source_label: str, default_line_tag: str = None) -> pd.DataFrame:
    """
    Agrupa materiales repetidos en distintas ubicaciones/líneas, normaliza códigos
    para evitar incompatibilidades entre meses o líneas distintas, y conserva todas las métricas.
    """
    work = df.copy()
    cod_col = schema.get("codigo")
    desc_col = schema.get("descripcion")
    stk_col = schema.get("stock")
    ubi_col = schema.get("ubicacion")
    cat_col = schema.get("categoria")
    ing_col = schema.get("ingresos")
    egr_col = schema.get("egresos")
    prc_col = schema.get("precio")

    if not cod_col or cod_col not in work.columns:
        work["_AUTO_COD"] = [f"ITEM-{i+1:05d}" for i in range(len(work))]
        cod_col = "_AUTO_COD"

    work = work[work[cod_col].notna()].copy()
    work["_COD_KEY"] = (
        work[cod_col]
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.replace(r"\s+", " ", regex=True)
    )
    work = work[~work["_COD_KEY"].str.lower().isin(["nan", "none", "total", "total general", "subtotal", ""])]

    if stk_col and stk_col in work.columns:
        work["_QTY"] = parse_numeric_series(work[stk_col])
    else:
        work["_QTY"] = 0.0

    if desc_col and desc_col in work.columns:
        work["_DESC"] = work[desc_col].fillna("Sin descripción").astype(str).str.strip()
    else:
        work["_DESC"] = work["_COD_KEY"]

    has_ubi = bool(ubi_col and ubi_col in work.columns and work[ubi_col].notna().any())
    has_cat = bool(cat_col and cat_col in work.columns and work[cat_col].notna().any())
    has_ing = bool(ing_col and ing_col in work.columns and work[ing_col].notna().any())
    has_egr = bool(egr_col and egr_col in work.columns and work[egr_col].notna().any())
    has_prc = bool(prc_col and prc_col in work.columns and work[prc_col].notna().any())

    if has_ubi:
        work["_UBI_ITEM"] = (
            work[ubi_col].fillna("S/U").astype(str).str.strip()
            + " ("
            + work["_QTY"].map(lambda x: f"{int(x)}" if x == int(x) else f"{x:.2f}")
            + ")"
        )

    if has_cat:
        work["_CAT"] = work[cat_col].fillna(default_line_tag or "General").astype(str).str.strip()
    elif "_SHEET_ORIGIN" in work.columns:
        work["_CAT"] = work["_SHEET_ORIGIN"].astype(str)
        has_cat = True
    elif default_line_tag:
        work["_CAT"] = str(default_line_tag)
        has_cat = True

    if has_ing:
        work["_ING"] = parse_numeric_series(work[ing_col])
    if has_egr:
        work["_EGR"] = parse_numeric_series(work[egr_col])
    if has_prc:
        work["_PRC"] = parse_numeric_series(work[prc_col])

    agg_dict = {
        "_DESC": "first",
        "_QTY": "sum",
    }
    if has_ubi:
        agg_dict["_UBI_ITEM"] = lambda s: " + ".join(dict.fromkeys(s))
    if has_cat:
        agg_dict["_CAT"] = lambda s: " / ".join(dict.fromkeys([x for x in s if x]))
    if has_ing:
        agg_dict["_ING"] = "sum"
    if has_egr:
        agg_dict["_EGR"] = "sum"
    if has_prc:
        agg_dict["_PRC"] = "mean"

    grouped = work.groupby("_COD_KEY", as_index=False, sort=False).agg(agg_dict)
    grouped.rename(
        columns={
            "_COD_KEY": "CODIGO",
            "_DESC": "DESCRIPCION",
            "_QTY": f"STOCK_{source_label}",
        },
        inplace=True,
    )
    counts = work.groupby("_COD_KEY").size().reset_index(name=f"REGISTROS_{source_label}")
    grouped = grouped.merge(counts, left_on="CODIGO", right_on="_COD_KEY", how="left").drop(columns=["_COD_KEY"])

    if has_ubi:
        grouped.rename(columns={"_UBI_ITEM": f"UBICACIONES_{source_label}"}, inplace=True)
    if has_cat:
        grouped.rename(columns={"_CAT": f"LINEA_{source_label}"}, inplace=True)
    if has_ing:
        grouped.rename(columns={"_ING": f"INGRESOS_{source_label}"}, inplace=True)
    if has_egr:
        grouped.rename(columns={"_EGR": f"EGRESOS_{source_label}"}, inplace=True)
    if has_prc:
        grouped.rename(columns={"_PRC": f"PRECIO_{source_label}"}, inplace=True)

    return grouped


def merge_all_workbook_sheets(sheets_dict: dict) -> pd.DataFrame:
    """Consolida todas las pestañas de un libro de Excel etiquetando cada línea/pestaña."""
    frames = []
    for s_name, s_df in sheets_dict.items():
        if s_df is not None and not s_df.empty:
            tmp = s_df.copy()
            tmp["_SHEET_ORIGIN"] = str(s_name)
            frames.append(tmp)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)


def reconcile_two_sources(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    name_a: str,
    name_b: str,
    operation_mode: str = "consolidacion",
) -> pd.DataFrame:
    """
    Cruza y consolida ambas fuentes (2 meses distintos, 2 líneas de producto o Lógico vs Físico).
    """
    merged = pd.merge(df_a, df_b, on="CODIGO", how="outer", suffixes=("_A", "_B"))

    desc_a = merged.get("DESCRIPCION_A", pd.Series(index=merged.index, dtype=str))
    desc_b = merged.get("DESCRIPCION_B", pd.Series(index=merged.index, dtype=str))
    merged["DESCRIPCION"] = desc_a.fillna(desc_b).fillna("Sin descripción")

    line_a_col = f"LINEA_{name_a}"
    line_b_col = f"LINEA_{name_b}"
    if line_a_col in merged.columns and line_b_col in merged.columns:
        la = merged[line_a_col].fillna("")
        lb = merged[line_b_col].fillna("")
        merged["LINEA_CATEGORIA"] = np.where(
            (la != "") & (lb != "") & (la != lb),
            la + " | " + lb,
            np.where(la != "", la, np.where(lb != "", lb, "General")),
        )
    elif line_a_col in merged.columns:
        merged["LINEA_CATEGORIA"] = merged[line_a_col].fillna("General")
    elif line_b_col in merged.columns:
        merged["LINEA_CATEGORIA"] = merged[line_b_col].fillna("General")
    else:
        merged["LINEA_CATEGORIA"] = "General"

    col_stk_a = f"STOCK_{name_a}"
    col_stk_b = f"STOCK_{name_b}"
    merged["EN_FUENTE_A"] = merged[col_stk_a].notna()
    merged["EN_FUENTE_B"] = merged[col_stk_b].notna()

    merged[col_stk_a] = merged[col_stk_a].fillna(0.0)
    merged[col_stk_b] = merged[col_stk_b].fillna(0.0)

    merged["STOCK_CONSOLIDADO"] = merged[col_stk_a] + merged[col_stk_b]
    merged["DIFERENCIA"] = merged[col_stk_b] - merged[col_stk_a]
    merged["VARIACION_%"] = np.where(
        merged[col_stk_a] > 0,
        ((merged[col_stk_b] - merged[col_stk_a]) / merged[col_stk_a] * 100.0).round(2),
        np.where(merged[col_stk_b] > 0, 100.0, 0.0),
    )

    def classify_status(row):
        if operation_mode == "auditoria":
            if not row["EN_FUENTE_A"]:
                return f"Solo en {name_b}"
            if not row["EN_FUENTE_B"]:
                return f"Solo en {name_a}"
            if abs(row["DIFERENCIA"]) < 1e-9:
                return "Cuadrado"
            return "Sobrante (+)" if row["DIFERENCIA"] > 0 else "Faltante (-)"
        else:
            if not row["EN_FUENTE_A"]:
                return f"Solo en {name_b}"
            if not row["EN_FUENTE_B"]:
                return f"Solo en {name_a}"
            if abs(row["DIFERENCIA"]) < 1e-9:
                return "Cuadrado"
            return "Creciente (+)" if row["DIFERENCIA"] > 0 else "Decreciente (-)"

    merged["ESTADO_CUADRE"] = merged.apply(classify_status, axis=1)

    ordered_cols = ["CODIGO", "DESCRIPCION", "LINEA_CATEGORIA"]
    for c in [f"UBICACIONES_{name_a}", f"UBICACIONES_{name_b}"]:
        if c in merged.columns:
            ordered_cols.append(c)
    for c in [
        f"INGRESOS_{name_a}", f"INGRESOS_{name_b}",
        f"EGRESOS_{name_a}", f"EGRESOS_{name_b}",
        f"PRECIO_{name_a}", f"PRECIO_{name_b}",
    ]:
        if c in merged.columns:
            ordered_cols.append(c)

    ordered_cols.extend([
        col_stk_a,
        col_stk_b,
        "STOCK_CONSOLIDADO",
        "DIFERENCIA",
        "VARIACION_%",
        "ESTADO_CUADRE",
        f"REGISTROS_{name_a}",
        f"REGISTROS_{name_b}",
    ])
    final_cols = [c for c in ordered_cols if c in merged.columns]
    result = merged[final_cols].copy()
    for rcol in [f"REGISTROS_{name_a}", f"REGISTROS_{name_b}"]:
        if rcol in result.columns:
            result[rcol] = result[rcol].fillna(0).astype(int)
    return result


# ==============================================================================
# MOTOR DE INTELIGENCIA ARTIFICIAL PREDICTIVA VECTORIZADO (HASTA 12 MESES)
# ==============================================================================
MONTH_ORDER_MAP = {
    "enero": 1, "ene": 1, "january": 1, "jan": 1,
    "febrero": 2, "feb": 2, "february": 2,
    "marzo": 3, "mar": 3, "march": 3,
    "abril": 4, "abr": 4, "april": 4, "apr": 4,
    "mayo": 5, "may": 5,
    "junio": 6, "jun": 6, "june": 6,
    "julio": 7, "jul": 7, "july": 7,
    "agosto": 8, "ago": 8, "august": 8, "aug": 8,
    "septiembre": 9, "sep": 9, "sept": 9, "september": 9,
    "octubre": 10, "oct": 10, "october": 10,
    "noviembre": 11, "nov": 11, "november": 11,
    "diciembre": 12, "dic": 12, "december": 12, "dec": 12,
}


def extract_period_sort_key(filename: str, fallback_idx: int) -> tuple:
    norm = normalize_text(filename)
    year_match = re.search(r"(20\d{2})", norm)
    year = int(year_match.group(1)) if year_match else 2026
    month_num = fallback_idx + 1
    clean_label = filename.rsplit(".", 1)[0].strip() or f"Mes_{fallback_idx + 1}"
    for word in norm.split():
        if word in MONTH_ORDER_MAP:
            month_num = MONTH_ORDER_MAP[word]
            clean_label = f"{word.capitalize()} {year}"
            break
    return (year, month_num, clean_label)


def run_ai_demand_forecast(monthly_records: list) -> tuple:
    """
    Motor predictivo 100% vectorizado en NumPy:
    - Unifica descripciones por CODIGO para evitar ruptura de series entre meses.
    - Combina Regresión Lineal Ponderada + Suavizado Exponencial Doble de Holt.
    - Filtra picos atípicos (Winsorization) y calcula sugerido de compra + confianza IA.
    """
    if not monthly_records:
        return pd.DataFrame(), pd.DataFrame()

    frames = []
    period_order = []
    for idx, rec in enumerate(monthly_records):
        df = rec["df"].copy()
        if df.empty:
            continue
        p_lbl = rec["label"]
        if p_lbl not in period_order:
            period_order.append(p_lbl)
        df["PERIODO"] = p_lbl
        df["ORDEN_MES"] = idx + 1
        frames.append(df)

    if not frames:
        return pd.DataFrame(), pd.DataFrame()

    panel = pd.concat(frames, ignore_index=True)
    has_explicit_sales = "EGRESOS_M" in panel.columns and panel["EGRESOS_M"].sum() > 0

    # Mapa canónico de descripción por CODIGO (toma la más reciente no vacía)
    desc_series = (
        panel.sort_values("ORDEN_MES")
        .dropna(subset=["DESCRIPCION"])
        .groupby("CODIGO")["DESCRIPCION"]
        .last()
    )

    piv_stock = panel.pivot_table(
        index="CODIGO",
        columns="PERIODO",
        values="STOCK_M",
        aggfunc="sum",
        fill_value=0.0,
    ).reindex(columns=period_order, fill_value=0.0)

    if has_explicit_sales:
        piv_sales = panel.pivot_table(
            index="CODIGO",
            columns="PERIODO",
            values="EGRESOS_M",
            aggfunc="sum",
            fill_value=0.0,
        ).reindex(columns=period_order, fill_value=0.0)
    else:
        diffs = -piv_stock.diff(axis=1)
        piv_sales = diffs.clip(lower=0.0).fillna(0.0)

    codes = piv_stock.index.to_numpy()
    descs = desc_series.reindex(codes).fillna("Sin descripción").to_numpy()

    stock_mat = piv_stock.to_numpy(dtype=float)
    sales_mat = piv_sales.to_numpy(dtype=float)
    n_skus, n_periods = stock_mat.shape

    current_stock = stock_mat[:, -1]
    last_month_sales = sales_mat[:, -1]

    # Si las salidas se infieren por diferencia de stock, excluimos el 0 inicial del mes 1
    if not has_explicit_sales and n_periods >= 2:
        active_sales = sales_mat[:, 1:]
    else:
        active_sales = sales_mat

    t_act = active_sales.shape[1]

    # Suavizado de picos atípicos (Winsorization) para evitar distorsiones por traslados puntuales
    if t_act >= 4:
        p90 = np.percentile(active_sales, 90, axis=1, keepdims=True)
        med = np.median(active_sales, axis=1, keepdims=True)
        upper_cap = np.maximum(p90 * 1.5, med * 3.0)
        clean_sales = np.minimum(active_sales, upper_cap)
    else:
        clean_sales = active_sales.copy()

    avg_monthly = np.mean(clean_sales, axis=1)

    if t_act >= 2:
        # 1. Regresión Lineal Ponderada (vectorizada)
        x = np.arange(1, t_act + 1, dtype=float)
        weights = np.linspace(0.65, 1.55, t_act)
        w_sum = float(np.sum(weights))
        w_mean_x = float(np.sum(weights * x) / w_sum)
        w_mean_y = np.sum(clean_sales * weights[None, :], axis=1) / w_sum

        dx = x - w_mean_x
        denom = float(np.sum(weights * (dx ** 2)))
        if denom > 0:
            slope = np.sum((clean_sales - w_mean_y[:, None]) * (weights * dx)[None, :], axis=1) / denom
        else:
            slope = np.zeros(n_skus, dtype=float)
        intercept = w_mean_y - slope * w_mean_x
        reg_pred = np.maximum(0.0, intercept + slope * (t_act + 1))

        # 2. Suavizado Exponencial Doble de Holt (nivel + tendencia)
        alpha, beta = 0.45, 0.25
        level = clean_sales[:, 0].copy()
        trend = clean_sales[:, 1] - clean_sales[:, 0]
        for t in range(1, t_act):
            prev_level = level.copy()
            level = alpha * clean_sales[:, t] + (1.0 - alpha) * (level + trend)
            trend = beta * (level - prev_level) + (1.0 - beta) * trend
        holt_pred = np.maximum(0.0, level + trend)

        forecast = np.round(0.50 * reg_pred + 0.50 * holt_pred, 1)

        # Si lleva 3 o más períodos seguidos en cero, amortiguar a 0
        if t_act >= 3:
            dormant_mask = np.all(clean_sales[:, -3:] == 0, axis=1)
            forecast = np.where(dormant_mask, 0.0, forecast