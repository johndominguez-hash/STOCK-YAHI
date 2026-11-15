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
            forecast = np.where(dormant_mask, 0.0, forecastimport io
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
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
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
    "<circle cx='32' cy='32' r='14' fill='%23FFFFFF' stroke='%23E60026' stroke-width='4'/>"
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
# PALETA DE IDENTIDAD: MEGAMAN X (BLANCO, AZUL ELÉCTRICO, CELESTE CIAN, ROJO)
# + ESTÉTICA CORPORATIVA HIVIMAR S.A.
# ==============================================================================
PALETTE = {
    "white": "#FFFFFF",
    "ice_bg": "#F4F8FF",
    "electric_blue": "#0057FF",
    "deep_blue": "#06193E",
    "cyan": "#00C8FF",
    "cyan_bright": "#00E5FF",
    "red": "#E60026",
    "red_soft": "#FF1E42",
    "slate_text": "#0B192C",
    "muted_text": "#475569",
    "border_cyan": "rgba(0, 180, 216, 0.35)",
}

CORPORATE_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

:root {
    --x-white: #FFFFFF;
    --x-ice: #F4F8FF;
    --x-blue: #0057FF;
    --x-deep: #06193E;
    --x-cyan: #00C8FF;
    --x-red: #E60026;
    --x-text: #0B192C;
    --x-muted: #475569;
}

html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    color: var(--x-text);
}

/* Fondo general Blanco Cristal con destellos Azul Eléctrico y Celeste Cian */
.stApp {
    background:
        radial-gradient(circle at 92% 4%, rgba(0, 200, 255, 0.14) 0%, transparent 36%),
        radial-gradient(circle at 8% 95%, rgba(0, 87, 255, 0.08) 0%, transparent 40%),
        linear-gradient(180deg, #F6FAFF 0%, #FFFFFF 55%, #EFF6FF 100%) !important;
}

/* Barra lateral Armadura X (Azul Eléctrico Profundo + Cian + Rojo) */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #051638 0%, #09245B 55%, #061A42 100%) !important;
    border-right: 2px solid #00C8FF !important;
}

section[data-testid="stSidebar"] * {
    color: #FFFFFF !important;
}

section[data-testid="stSidebar"] .stRadio > div {
    gap: 0.35rem;
}

section[data-testid="stSidebar"] .stRadio label {
    background-color: rgba(255, 255, 255, 0.06);
    border-radius: 10px;
    padding: 0.65rem 0.9rem !important;
    transition: all 0.2s ease;
    border: 1px solid rgba(0, 200, 255, 0.22);
    color: #F8FAFC !important;
    font-weight: 600 !important;
}

section[data-testid="stSidebar"] .stRadio label:hover {
    background: linear-gradient(90deg, rgba(0, 87, 255, 0.65), rgba(0, 200, 255, 0.35)) !important;
    border-color: #00E5FF !important;
    transform: translateX(2px);
}

section[data-testid="stSidebar"] .stButton > button {
    background: linear-gradient(90deg, #0057FF 0%, #00A8E8 100%) !important;
    color: #FFFFFF !important;
    border: 1px solid #00E5FF !important;
}

/* Encabezados de sección con iconografía vectorial */
.sec-header {
    display: flex;
    align-items: center;
    gap: 0.8rem;
    margin: 0.5rem 0 1rem 0;
    padding: 0.65rem 0.9rem;
    background: #FFFFFF;
    border-left: 5px solid #0057FF;
    border-top: 1px solid rgba(0, 87, 255, 0.14);
    border-right: 1px solid rgba(0, 87, 255, 0.14);
    border-bottom: 1px solid rgba(0, 200, 255, 0.35);
    border-radius: 12px;
    box-shadow: 0 4px 14px rgba(0, 87, 255, 0.06);
}
.sec-icon-wrap {
    width: 40px;
    height: 40px;
    border-radius: 10px;
    background: linear-gradient(135deg, #0057FF 0%, #00C8FF 100%);
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
    box-shadow: 0 4px 12px rgba(0, 87, 255, 0.25);
}
.sec-title {
    font-size: 1.25rem;
    font-weight: 800;
    color: #06193E;
    letter-spacing: -0.01em;
    margin: 0;
}
.sec-sub {
    font-size: 0.82rem;
    color: #475569;
    margin: 0;
}

/* Tarjeta de Fuente con insignia numérica vectorial */
.source-header {
    display: flex;
    align-items: center;
    gap: 0.65rem;
    margin-bottom: 0.6rem;
    padding: 0.5rem 0.75rem;
    background: #FFFFFF;
    border-radius: 10px;
    border: 1px solid rgba(0, 87, 255, 0.2);
}
.num-badge {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 30px;
    height: 30px;
    border-radius: 8px;
    background: linear-gradient(135deg, #0057FF 0%, #00C8FF 100%);
    color: #FFFFFF;
    font-size: 0.84rem;
    font-weight: 800;
    letter-spacing: 0.04em;
    flex-shrink: 0;
}
.num-badge-red {
    background: linear-gradient(135deg, #E60026 0%, #FF1E42 100%);
}
.source-title-text {
    font-size: 1.02rem;
    font-weight: 800;
    color: #06193E;
}

/* Tarjetas KPI Ejecutivas (Blanco Puro + Borde Azul Eléctrico / Cian / Rojo) */
.kpi-card {
    background: #FFFFFF;
    border: 1.5px solid rgba(0, 87, 255, 0.22);
    border-top: 4px solid #0057FF;
    border-radius: 14px;
    padding: 1.1rem 1.25rem;
    box-shadow: 0 8px 22px rgba(6, 25, 62, 0.07);
    transition: transform 0.2s ease, border-color 0.2s ease;
    height: 100%;
}
.kpi-card.cyan-top { border-top-color: #00C8FF; }
.kpi-card.red-top { border-top-color: #E60026; }
.kpi-card:hover {
    transform: translateY(-2px);
    border-color: #00C8FF;
}
.kpi-label {
    font-size: 0.76rem;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    color: #475569;
    font-weight: 800;
    margin-bottom: 0.35rem;
}
.kpi-value {
    font-size: 1.75rem;
    font-weight: 800;
    color: #06193E;
    line-height: 1.15;
}
.kpi-sub {
    font-size: 0.82rem;
    margin-top: 0.4rem;
    font-weight: 700;
}
.badge-ok { color: #0057FF; }
.badge-cyan { color: #0284C7; }
.badge-warn { color: #D97706; }
.badge-err { color: #E60026; }

/* Banner de Bienvenida (Identidad Mega Man X + HIVIMAR S.A.) */
.hero-banner {
    background: linear-gradient(115deg, #06193E 0%, #0057FF 58%, #0096C7 100%);
    border: 2px solid #00C8FF;
    border-radius: 18px;
    padding: 1.45rem 2rem;
    margin-bottom: 1.25rem;
    box-shadow: 0 12px 28px rgba(0, 87, 255, 0.22);
    position: relative;
    overflow: hidden;
}
.hero-banner::after {
    content: "";
    position: absolute;
    top: 0;
    right: 0;
    width: 110px;
    height: 6px;
    background: #E60026;
}

/* Alerta ¡AGUANTA! con acento Rojo Zero */
.aguanta-box {
    background: linear-gradient(135deg, #FFFFFF 0%, #FFF1F2 100%);
    border: 2px solid #E60026;
    border-left: 8px solid #E60026;
    border-radius: 14px;
    padding: 1.15rem 1.4rem;
    margin: 0.75rem 0 1rem 0;
    display: flex;
    align-items: center;
    gap: 1rem;
    box-shadow: 0 8px 20px rgba(230, 0, 38, 0.12);
}

/* Botones principales */
.stButton > button {
    background: linear-gradient(90deg, #0057FF 0%, #0096FF 100%) !important;
    color: #FFFFFF !important;
    border: 1px solid #00C8FF !important;
    border-radius: 10px !important;
    font-weight: 800 !important;
    letter-spacing: 0.02em;
    transition: all 0.2s ease !important;
}
.stButton > button:hover {
    background: linear-gradient(90deg, #0044CC 0%, #00B4D8 100%) !important;
    box-shadow: 0 6px 18px rgba(0, 87, 255, 0.3) !important;
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
    box-shadow: 0 6px 18px rgba(230, 0, 38, 0.28) !important;
}
div[data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 8px 24px rgba(0, 87, 255, 0.42) !important;
}

/* Pestañas */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    background-color: #E6F0FF;
    padding: 6px;
    border-radius: 12px;
    border: 1px solid rgba(0, 87, 255, 0.2);
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px;
    color: #06193E;
    font-weight: 700;
    padding: 8px 16px;
}
.stTabs [aria-selected="true"] {
    background-color: #0057FF !important;
    color: #FFFFFF !important;
}

/* Pie de página con derechos de autor */
.copyright-footer {
    margin-top: 3rem;
    padding: 0.9rem 1rem;
    border-top: 2px solid #0057FF;
    background: #FFFFFF;
    border-radius: 10px;
    text-align: center;
    font-size: 0.76rem;
    color: #06193E;
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
      <stop offset="100%" stop-color="#E60026"/>
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
  <circle cx="120" cy="82" r="4.5" fill="#E60026"/>
  <circle cx="70" cy="30" r="5" fill="#E60026" stroke="#FFFFFF" stroke-width="1.5"/>

  <g transform="rotate(-45 70 74)">
    <rect x="64" y="24" width="12" height="96" rx="3" fill="url(#steelGrad)"/>
    <rect x="55" y="18" width="30" height="16" rx="4" fill="#FFFFFF" stroke="#00E5FF" stroke-width="2.2"/>
    <line x1="56" y1="23" x2="84" y2="23" stroke="#0057FF" stroke-width="2"/>
    <line x1="56" y1="28" x2="84" y2="28" stroke="#E60026" stroke-width="2"/>
    <circle cx="70" cy="116" r="8" fill="#0057FF" stroke="#FFFFFF" stroke-width="2"/>
  </g>

  <g transform="rotate(45 70 74)">
    <rect x="64" y="24" width="12" height="96" rx="3" fill="url(#steelGrad)"/>
    <rect x="55" y="18" width="30" height="16" rx="4" fill="#FFFFFF" stroke="#E60026" stroke-width="2.2"/>
    <line x1="56" y1="23" x2="84" y2="23" stroke="#0057FF" stroke-width="2"/>
    <line x1="56" y1="28" x2="84" y2="28" stroke="#E60026" stroke-width="2"/>
    <circle cx="70" cy="116" r="8" fill="#E60026" stroke="#FFFFFF" stroke-width="2"/>
  </g>

  <circle cx="70" cy="74" r="26" fill="#06193E" stroke="#FFFFFF" stroke-width="5"/>
  <circle cx="70" cy="74" r="14" fill="#0057FF" stroke="#00E5FF" stroke-width="3.5"/>
  <circle cx="70" cy="74" r="6" fill="#E60026" stroke="#FFFFFF" stroke-width="1.5"/>
</svg>
"""

SVG_ICONS = {
    "upload": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>""",
    "audit": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v18"/><path d="M5 7h14"/><path d="M5 7l-3 7h6l-3-7z"/><path d="M19 7l-3 7h6l-3-7z"/><line x1="8" y1="21" x2="16" y2="21"/></svg>""",
    "ranking": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>""",
    "trends": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>""",
    "ai": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="16" height="16" rx="2"/><rect x="9" y="9" width="6" height="6"/><line x1="9" y1="1" x2="9" y2="4"/><line x1="15" y1="1" x2="15" y2="4"/><line x1="9" y1="20" x2="9" y2="23"/><line x1="15" y1="20" x2="15" y2="23"/><line x1="20" y1="9" x2="23" y2="9"/><line x1="20" y1="14" x2="23" y2="14"/><line x1="1" y1="9" x2="4" y2="9"/><line x1="1" y1="14" x2="4" y2="14"/></svg>""",
    "table": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="9" x2="9" y2="21"/></svg>""",
    "alert": """<svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="#E60026" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>""",
}


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
        "name": "HIVIMAR_STOCK_Y_YA_MEGAMAN_X",
        "dataColors": [
            "#0057FF",
            "#00C8FF",
            "#E60026",
            "#06193E",
            "#0096C7",
            "#FF4D6D",
            "#38BDF8",
            "#1E40AF",
        ],
        "background": "#FFFFFF",
        "foreground": "#06193E",
        "tableAccent": "#0057FF",
        "good": "#00C8FF",
        "neutral": "#0057FF",
        "bad": "#E60026",
        "maximum": "#0057FF",
        "center": "#00C8FF",
        "minimum": "#E60026",
        "visualStyles": {
            "*": {
                "*": {
                    "background": [{"color": {"solid": {"color": "#FFFFFF"}}}],
                    "border": [
                        {
                            "show": True,
                            "color": {"solid": {"color": "#00C8FF"}},
                            "radius": 8,
                        }
                    ],
                    "title": [
                        {
                            "show": True,
                            "fontColor": {"solid": {"color": "#06193E"}},
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
              <circle cx="4" cy="-6" r="2" fill="#E60026"/>
            </g>
        """,
        "farmaceutico": """
            <g class="wheel-spin">
              <rect x="-22" y="-11" width="44" height="22" rx="11" fill="#FFFFFF" stroke="#0057FF" stroke-width="2.2"/>
              <path d="M 0 -11 L 11 -11 A 11 11 0 0 1 11 11 L 0 11 Z" fill="#E60026"/>
              <line x1="0" y1="-11" x2="0" y2="11" stroke="#06193E" stroke-width="1.5"/>
            </g>
        """,
        "dulceria": """
            <g class="wheel-spin">
              <circle cx="0" cy="0" r="19" fill="#0057FF" stroke="#FFFFFF" stroke-width="3"/>
              <circle cx="0" cy="0" r="7" fill="#E60026"/>
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
              <circle cx="-15" cy="0" r="2.8" fill="#E60026"/>
              <circle cx="15" cy="0" r="2.8" fill="#E60026"/>
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
                margin: 1rem 0; text-align: center; box-shadow: 0 10px 28px rgba(0,87,255,0.28);">
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
          <line class="track-line" x1="15" y1="64" x2="405" y2="64" stroke="#E60026" stroke-width="2.5"/>
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
        # Combinar pista de archivo (ej. Agosto) + pestaña si aporta contexto
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


def read_uploaded_workbook(uploaded_file) -> dict:
    fname = uploaded_file.name.lower()
    raw_bytes = uploaded_file.getvalue()
    if fname.endswith(".csv"):
        for enc in ["utf-8", "latin-1", "cp1252"]:
            try:
                df = pd.read_csv(io.BytesIO(raw_bytes), encoding=enc, sep=None, engine="python")
                df = promote_real_header_row(df)
                return {"Datos_CSV": df}
            except Exception:
                continue
        df = pd.read_csv(io.BytesIO(raw_bytes), encoding="latin-1")
        return {"Datos_CSV": promote_real_header_row(df)}

    xls = pd.ExcelFile(io.BytesIO(raw_bytes))
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
        clean_num = (
            work[stk_col]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.strip()
        )
        work["_QTY"] = pd.to_numeric(clean_num, errors="coerce").fillna(0.0)
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
        work["_ING"] = pd.to_numeric(work[ing_col], errors="coerce").fillna(0.0)
    if has_egr:
        work["_EGR"] = pd.to_numeric(work[egr_col], errors="coerce").fillna(0.0)
    if has_prc:
        work["_PRC"] = pd.to_numeric(work[prc_col], errors="coerce").fillna(0.0)

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
    Cruza y consolida ambas fuentes (ya sean 2 meses distintos como Agosto vs Octubre,
    2 líneas de producto distintas, o Lógico vs Físico) sin incompatibilidades.
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
# MOTOR DE INTELIGENCIA ARTIFICIAL PREDICTIVA (12 MESES)
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
    clean_label = filename.rsplit(".", 1)[0]
    for word in norm.split():
        if word in MONTH_ORDER_MAP:
            month_num = MONTH_ORDER_MAP[word]
            clean_label = f"{word.capitalize()} {year}"
            break
    return (year, month_num, clean_label)


def run_ai_demand_forecast(monthly_records: list) -> tuple:
    frames = []
    for idx, rec in enumerate(monthly_records):
        df = rec["df"].copy()
        df["PERIODO"] = rec["label"]
        df["ORDEN_MES"] = idx + 1
        frames.append(df)

    panel = pd.concat(frames, ignore_index=True)
    has_explicit_sales = "EGRESOS_M" in panel.columns

    period_order = (
        panel[["ORDEN_MES", "PERIODO"]]
        .drop_duplicates()
        .sort_values("ORDEN_MES")["PERIODO"]
        .tolist()
    )

    piv_stock = panel.pivot_table(
        index=["CODIGO", "DESCRIPCION"],
        columns="PERIODO",
        values="STOCK_M",
        aggfunc="sum",
        fill_value=0.0,
    ).reindex(columns=period_order, fill_value=0.0)

    if has_explicit_sales:
        piv_sales = panel.pivot_table(
            index=["CODIGO", "DESCRIPCION"],
            columns="PERIODO",
            values="EGRESOS_M",
            aggfunc="sum",
            fill_value=0.0,
        ).reindex(columns=period_order, fill_value=0.0)
    else:
        diffs = -piv_stock.diff(axis=1)
        piv_sales = diffs.clip(lower=0.0).fillna(0.0)

    ai_rows = []
    n_periods = len(period_order)
    x = np.arange(1, n_periods + 1, dtype=float)
    weights = np.linspace(0.6, 1.5, n_periods)

    for idx_tuple in piv_stock.index:
        cod, desc = idx_tuple
        stock_series = piv_stock.loc[idx_tuple].values.astype(float)
        sales_series = piv_sales.loc[idx_tuple].values.astype(float)

        current_stock = stock_series[-1]
        avg_monthly = float(np.mean(sales_series))
        last_month_sales = float(sales_series[-1])

        if n_periods >= 2:
            w_mean_x = np.average(x, weights=weights)
            w_mean_y = np.average(sales_series, weights=weights)
            denom = np.sum(weights * (x - w_mean_x) ** 2)
            slope = float(np.sum(weights * (x - w_mean_x) * (sales_series - w_mean_y)) / denom) if denom > 0 else 0.0
            intercept = w_mean_y - slope * w_mean_x
            reg_pred = max(0.0, intercept + slope * (n_periods + 1))

            ema = sales_series[0]
            for val in sales_series[1:]:
                ema = 0.45 * val + 0.55 * ema

            forecast = round(0.55 * reg_pred + 0.45 * ema, 1)
            growth_pct = ((forecast - avg_monthly) / avg_monthly * 100.0) if avg_monthly > 0 else 0.0
        else:
            forecast = round(last_month_sales, 1)
            growth_pct = 0.0

        coverage_months = (current_stock / forecast) if forecast > 0 else (99.0 if current_stock > 0 else 0.0)

        if forecast > 0 and current_stock < forecast:
            signal = "[CRÍTICO] Reabastecer Urgente (Riesgo Quiebre)"
        elif growth_pct >= 10.0 and forecast > 0:
            signal = "[ALZA] Demanda en Aceleración"
        elif forecast == 0 and current_stock > 0:
            signal = "[INMOVILIZADO] Sin Rotación"
        elif growth_pct <= -10.0:
            signal = "[BAJA] Demanda en Desaceleración"
        else:
            signal = "[ÓPTIMO] Demanda Estable"

        ai_rows.append({
            "CODIGO": cod,
            "DESCRIPCION": desc,
            "STOCK_ACTUAL": current_stock,
            "PROMEDIO_MENSUAL": round(avg_monthly, 1),
            "ULTIMO_MES": round(last_month_sales, 1),
            "EXPECTATIVA_VENTA_IA": forecast,
            "VARIACION_ESTIMADA_%": round(growth_pct, 1),
            "COBERTURA_MESES": round(coverage_months, 1) if coverage_months < 99 else 99.0,
            "DIAGNOSTICO_IA": signal,
        })

    df_ai = pd.DataFrame(ai_rows).sort_values("EXPECTATIVA_VENTA_IA", ascending=False)
    return df_ai, piv_sales


# ==============================================================================
# EXPORTADOR MAESTRO EXCEL (&YA!) CON PALETA MEGAMAN X, GRÁFICOS NATIVOS
# Y MODELO DE DATOS COMPATIBLE CON POWER BI
# ==============================================================================
def style_excel_chart_megaman_x(chart, palette_hex_list: list):
    """Aplica identidad visual Mega Man X (Azul Eléctrico, Cian, Rojo) a gráficos nativos Excel."""
    chart.style = 10
    chart.legend.position = "b"
    for idx, series in enumerate(chart.series):
        color_hex = palette_hex_list[idx % len(palette_hex_list)]
        series.graphicalProperties.solidFill = color_hex
        series.graphicalProperties.line.solidFill = color_hex


def build_master_excel_export(
    df_reconciled,
    name_a: str,
    name_b: str,
    df_top_sales,
    df_no_move,
    df_ai,
    sector_label: str,
) -> bytes:
    wb = Workbook()

    # Paleta Mega Man X / HIVIMAR S.A.
    HEX_BLUE = "0057FF"
    HEX_DEEP = "06193E"
    HEX_CYAN = "00C8FF"
    HEX_RED = "E60026"
    HEX_WHITE = "FFFFFF"
    HEX_ICE = "F0F7FF"
    HEX_CYAN_SOFT = "E0F7FF"
    HEX_RED_SOFT = "FFE4E8"

    fill_deep = PatternFill(start_color=HEX_DEEP, end_color=HEX_DEEP, fill_type="solid")
    fill_blue = PatternFill(start_color=HEX_BLUE, end_color=HEX_BLUE, fill_type="solid")
    fill_cyan = PatternFill(start_color=HEX_CYAN, end_color=HEX_CYAN, fill_type="solid")
    fill_red = PatternFill(start_color=HEX_RED, end_color=HEX_RED, fill_type="solid")
    fill_ice = PatternFill(start_color=HEX_ICE, end_color=HEX_ICE, fill_type="solid")
    fill_ok = PatternFill(start_color=HEX_CYAN_SOFT, end_color=HEX_CYAN_SOFT, fill_type="solid")
    fill_err = PatternFill(start_color=HEX_RED_SOFT, end_color=HEX_RED_SOFT, fill_type="solid")

    font_title = Font(name="Calibri", size=14, bold=True, color=HEX_WHITE)
    font_header = Font(name="Calibri", size=11, bold=True, color=HEX_WHITE)
    font_bold = Font(name="Calibri", size=11, bold=True, color=HEX_DEEP)
    font_regular = Font(name="Calibri", size=10, color="0F172A")
    thin_border = Border(
        left=Side(style="thin", color="BFDBFE"),
        right=Side(style="thin", color="BFDBFE"),
        top=Side(style="thin", color="BFDBFE"),
        bottom=Side(style="thin", color="BFDBFE"),
    )

    # ==========================================================================
    # PESTAÑA 1: 1. Resumen y Gráficos (Panel Ejecutivo con Gráficos Nativos)
    # ==========================================================================
    ws_res = wb.active
    ws_res.title = "1. Resumen y Gráficos"
    ws_res.sheet_properties.tabColor = HEX_BLUE
    ws_res.views.sheetView[0].showGridLines = True

    ws_res.merge_cells("A1:K1")
    ws_res["A1"] = "STOCK&YA! — PANEL EJECUTIVO DE CONSOLIDACIÓN Y AUDITORÍA (HIVIMAR S.A.)"
    ws_res["A1"].fill = fill_deep
    ws_res["A1"].font = font_title
    ws_res["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_res.row_dimensions[1].height = 32

    ws_res["A2"] = f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws_res["E2"] = f"Rubro detectado: {sector_label}"
    ws_res["A3"] = "Autoría: Derechos de autor reservados por John Aarón Domínguez Jara 01.593.321"
    ws_res["A2"].font = font_bold
    ws_res["E2"].font = font_bold
    ws_res["A3"].font = Font(name="Calibri", size=9, italic=True, color="475569")

    if df_reconciled is not None and not df_reconciled.empty:
        col_a = f"STOCK_{name_a}"
        col_b = f"STOCK_{name_b}"
        total_skus = len(df_reconciled)
        cuadrados = int((df_reconciled["ESTADO_CUADRE"] == "Cuadrado").sum())
        solo_a = int((df_reconciled["ESTADO_CUADRE"] == f"Solo en {name_a}").sum())
        solo_b = int((df_reconciled["ESTADO_CUADRE"] == f"Solo en {name_b}").sum())
        con_dif = total_skus - cuadrados
        exactitud = (cuadrados / total_skus * 100.0) if total_skus > 0 else 0.0
        tot_a = float(df_reconciled[col_a].sum())
        tot_b = float(df_reconciled[col_b].sum())
        tot_cons = float(df_reconciled["STOCK_CONSOLIDADO"].sum())
        net_diff = tot_b - tot_a

        # Tabla 1: Indicadores Ejecutivos (A5:B14)
        ws_res.cell(row=5, column=1, value="INDICADOR EJECUTIVO HIVIMAR").fill = fill_blue
        ws_res.cell(row=5, column=1).font = font_header
        ws_res.cell(row=5, column=2, value="VALOR").fill = fill_blue
        ws_res.cell(row=5, column=2).font = font_header

        kpis = [
            ("Total Materiales Únicos Consolidados", total_skus),
            ("Materiales Cuadrados Exactos", cuadrados),
            ("Materiales con Variación / Diferencia", con_dif),
            (f"Materiales Exclusivos en {name_a}", solo_a),
            (f"Materiales Exclusivos en {name_b}", solo_b),
            ("Porcentaje de Exactitud / Coincidencia (%)", round(exactitud, 2)),
            (f"Unidades Totales ({name_a})", tot_a),
            (f"Unidades Totales ({name_b})", tot_b),
            ("Inventario Consolidado Total (Suma)", tot_cons),
            (f"Diferencia Neta ({name_b} - {name_a})", net_diff),
        ]
        r_cur = 6
        for k_name, k_val in kpis:
            c1 = ws_res.cell(row=r_cur, column=1, value=k_name)
            c2 = ws_res.cell(row=r_cur, column=2, value=k_val)
            c1.font = font_bold
            c2.font = font_regular
            c1.border = thin_border
            c2.border = thin_border
            if r_cur % 2 == 0:
                c1.fill = fill_ice
                c2.fill = fill_ice
            r_cur += 1

        # Tabla Auxiliar 2: Comparativo de Volúmenes por Fuente (A17:B20) -> Alimenta Gráfico 1
        ws_res.cell(row=17, column=1, value="FUENTE / PERÍODO").fill = fill_blue
        ws_res.cell(row=17, column=1).font = font_header
        ws_res.cell(row=17, column=2, value="UNIDADES TOTALES").fill = fill_blue
        ws_res.cell(row=17, column=2).font = font_header

        vol_rows = [
            (name_a, tot_a),
            (name_b, tot_b),
            ("CONSOLIDADO TOTAL", tot_cons),
        ]
        for idx_v, (lbl_v, val_v) in enumerate(vol_rows, start=18):
            ca = ws_res.cell(row=idx_v, column=1, value=lbl_v)
            cb = ws_res.cell(row=idx_v, column=2, value=val_v)
            ca.font = font_bold
            cb.font = font_regular
            ca.border = thin_border
            cb.border = thin_border

        # Tabla Auxiliar 3: Distribución de Estado del Catálogo (A22:B26) -> Alimenta Gráfico 2
        ws_res.cell(row=22, column=1, value="ESTADO DEL CATÁLOGO").fill = fill_red
        ws_res.cell(row=22, column=1).font = font_header
        ws_res.cell(row=22, column=2, value="CANTIDAD SKUs").fill = fill_red
        ws_res.cell(row=22, column=2).font = font_header

        status_counts = df_reconciled["ESTADO_CUADRE"].value_counts().reset_index()
        status_counts.columns = ["Estado", "Cantidad"]
        for idx_s, row_s in status_counts.iterrows():
            r_pos = 23 + idx_s
            c_st = ws_res.cell(row=r_pos, column=1, value=str(row_s["Estado"]))
            c_qt = ws_res.cell(row=r_pos, column=2, value=int(row_s["Cantidad"]))
            c_st.font = font_bold
            c_qt.font = font_regular
            c_st.border = thin_border
            c_qt.border = thin_border

        # GRÁFICO NATIVO 1: Comparativo de Unidades por Fuente / Período (D5)
        chart_vol = BarChart()
        chart_vol.type = "col"
        chart_vol.title = f"Comparativo de Unidades: {name_a} vs {name_b}"
        chart_vol.y_axis.title = "Unidades"
        chart_vol.x_axis.title = "Origen / Período"
        data_ref_vol = Reference(ws_res, min_col=2, min_row=17, max_row=20)
        cats_ref_vol = Reference(ws_res, min_col=1, min_row=18, max_row=20)
        chart_vol.add_data(data_ref_vol, titles_from_data=True)
        chart_vol.set_categories(cats_ref_vol)
        chart_vol.width = 15
        chart_vol.height = 9.5
        style_excel_chart_megaman_x(chart_vol, [HEX_BLUE, HEX_CYAN, HEX_RED])
        ws_res.add_chart(chart_vol, "D5")

        # GRÁFICO NATIVO 2: Distribución de Estados del Inventario (D21)
        chart_pie = DoughnutChart()
        chart_pie.title = "Composición de Auditoría y Consolidación (SKUs)"
        max_st_row = 22 + len(status_counts)
        data_ref_st = Reference(ws_res, min_col=2, min_row=22, max_row=max_st_row)
        cats_ref_st = Reference(ws_res, min_col=1, min_row=23, max_row=max_st_row)
        chart_pie.add_data(data_ref_st, titles_from_data=True)
        chart_pie.set_categories(cats_ref_st)
        chart_pie.dataLabels = DataLabelList()
        chart_pie.dataLabels.showPercent = True
        chart_pie.width = 15
        chart_pie.height = 9.5
        style_excel_chart_megaman_x(chart_pie, [HEX_CYAN, HEX_BLUE, HEX_RED, HEX_DEEP])
        ws_res.add_chart(chart_pie, "D21")

        # Resumen por Línea / Categoría en I5:L... + Gráfico 3 en D37
        if "LINEA_CATEGORIA" in df_reconciled.columns:
            line_summary = (
                df_reconciled.groupby("LINEA_CATEGORIA", as_index=False)[[col_a, col_b, "STOCK_CONSOLIDADO"]]
                .sum()
                .sort_values("STOCK_CONSOLIDADO", ascending=False)
                .head(12)
            )
            ws_res.cell(row=5, column=11, value="LÍNEA / CATEGORÍA").fill = fill_blue
            ws_res.cell(row=5, column=11).font = font_header
            ws_res.cell(row=5, column=12, value=f"STOCK {name_a}").fill = fill_blue
            ws_res.cell(row=5, column=12).font = font_header
            ws_res.cell(row=5, column=13, value=f"STOCK {name_b}").fill = fill_cyan
            ws_res.cell(row=5, column=13).font = Font(name="Calibri", size=11, bold=True, color=HEX_DEEP)
            ws_res.cell(row=5, column=14, value="CONSOLIDADO").fill = fill_red
            ws_res.cell(row=5, column=14).font = font_header

            for idx_l, row_l in enumerate(line_summary.itertuples(index=False), start=6):
                ws_res.cell(row=idx_l, column=11, value=str(row_l[0])).border = thin_border
                ws_res.cell(row=idx_l, column=12, value=float(row_l[1])).border = thin_border
                ws_res.cell(row=idx_l, column=13, value=float(row_l[2])).border = thin_border
                ws_res.cell(row=idx_l, column=14, value=float(row_l[3])).border = thin_border

            chart_lines = BarChart()
            chart_lines.type = "bar"
            chart_lines.grouping = "clustered"
            chart_lines.title = f"Consolidado por Línea / Categoría ({name_a} vs {name_b})"
            max_l_row = 5 + len(line_summary)
            d_ref_l = Reference(ws_res, min_col=12, max_col=13, min_row=5, max_row=max_l_row)
            c_ref_l = Reference(ws_res, min_col=11, min_row=6, max_row=max_l_row)
            chart_lines.add_data(d_ref_l, titles_from_data=True)
            chart_lines.set_categories(c_ref_l)
            chart_lines.width = 18
            chart_lines.height = 10
            style_excel_chart_megaman_x(chart_lines, [HEX_BLUE, HEX_CYAN])
            ws_res.add_chart(chart_lines, "K19")

    ws_res.column_dimensions["A"].width = 42
    ws_res.column_dimensions["B"].width = 20
    ws_res.column_dimensions["K"].width = 28
    ws_res.column_dimensions["L"].width = 16
    ws_res.column_dimensions["M"].width = 16
    ws_res.column_dimensions["N"].width = 18

    # ==========================================================================
    # FUNCIÓN AUXILIAR PARA HOJAS DE DATOS + GRÁFICO INTEGRADO EN CADA HOJA
    # ==========================================================================
    def write_df_sheet_with_chart(
        sheet_title: str,
        tab_hex: str,
        df_data: pd.DataFrame,
        chart_cfg: dict = None,
    ):
        ws = wb.create_sheet(title=sheet_title[:31])
        ws.sheet_properties.tabColor = tab_hex
        ws.freeze_panes = "A2"

        for c_idx, col_name in enumerate(df_data.columns, start=1):
            cell = ws.cell(row=1, column=c_idx, value=str(col_name))
            cell.fill = fill_deep if c_idx <= 3 else fill_blue
            cell.font = font_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws.row_dimensions[1].height = 24

        for r_idx, row_vals in enumerate(df_data.itertuples(index=False), start=2):
            for c_idx, val in enumerate(row_vals, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.font = font_regular
                cell.border = thin_border
                if r_idx % 2 == 0:
                    cell.fill = fill_ice
                col_title = str(df_data.columns[c_idx - 1])
                if col_title == "ESTADO_CUADRE":
                    if val in ["Cuadrado", "Creciente (+)"]:
                        cell.fill = fill_ok
                        cell.font = Font(name="Calibri", size=10, bold=True, color="0369A1")
                    else:
                        cell.fill = fill_err
                        cell.font = Font(name="Calibri", size=10, bold=True, color=HEX_RED)

        for col_idx in range(1, len(df_data.columns) + 1):
            col_letter = get_column_letter(col_idx)
            max_len = max(
                len(str(df_data.columns[col_idx - 1])),
                int(df_data.iloc[:, col_idx - 1].astype(str).str.len().max()) if len(df_data) > 0 else 10,
            )
            ws.column_dimensions[col_letter].width = min(max_len + 4, 44)

        if chart_cfg and len(df_data) > 0:
            max_chart_rows = min(len(df_data) + 1, 16)
            anchor_col = get_column_letter(len(df_data.columns) + 2)
            ch = BarChart()
            ch.type = chart_cfg.get("type", "bar")
            ch.grouping = "clustered"
            ch.title = chart_cfg.get("title", sheet_title)
            val_cols = chart_cfg.get("val_cols", [])
            cat_col = chart_cfg.get("cat_col", 2)
            if val_cols:
                for idx_vc, vc in enumerate(val_cols):
                    d_ref = Reference(ws, min_col=vc, min_row=1, max_row=max_chart_rows)
                    ch.add_data(d_ref, titles_from_data=True)
                c_ref = Reference(ws, min_col=cat_col, min_row=2, max_row=max_chart_rows)
                ch.set_categories(c_ref)
                ch.width = 18
                ch.height = 11
                style_excel_chart_megaman_x(ch, chart_cfg.get("colors", [HEX_BLUE, HEX_CYAN, HEX_RED]))
                ws.add_chart(ch, f"{anchor_col}2")

    # ==========================================================================
    # PESTAÑAS 2 A 6 CON NOMBRES CLAROS Y GRÁFICOS PROPIOS
    # ==========================================================================
    if df_reconciled is not None and not df_reconciled.empty:
        cols_rec = list(df_reconciled.columns)
        idx_desc = cols_rec.index("DESCRIPCION") + 1
        idx_sa = cols_rec.index(f"STOCK_{name_a}") + 1
        idx_sb = cols_rec.index(f"STOCK_{name_b}") + 1
        idx_diff = cols_rec.index("DIFERENCIA") + 1

        write_df_sheet_with_chart(
            "2. Inventario Consolidado",
            HEX_CYAN,
            df_reconciled,
            chart_cfg={
                "type": "bar",
                "title": f"Top 15 Materiales: {name_a} vs {name_b}",
                "val_cols": [idx_sa, idx_sb],
                "cat_col": idx_desc,
                "colors": [HEX_BLUE, HEX_CYAN],
            },
        )

        df_diff_only = df_reconciled[df_reconciled["ESTADO_CUADRE"] != "Cuadrado"].copy()
        if not df_diff_only.empty:
            df_diff_only = df_diff_only.reindex(
                df_diff_only["DIFERENCIA"].abs().sort_values(ascending=False).index
            )
            write_df_sheet_with_chart(
                "3. Diferencias y Ajustes",
                HEX_RED,
                df_diff_only,
                chart_cfg={
                    "type": "bar",
                    "title": "Top 15 Mayores Diferencias / Variaciones",
                    "val_cols": [idx_diff],
                    "cat_col": idx_desc,
                    "colors": [HEX_RED],
                },
            )

    if df_top_sales is not None and not df_top_sales.empty:
        cols_top = list(df_top_sales.columns)
        idx_d_top = cols_top.index("DESCRIPCION") + 1
        idx_m_top = cols_top.index("UNIDADES_MOVIDAS") + 1
        write_df_sheet_with_chart(
            "4. Top 20 Mayor Rotación",
            HEX_BLUE,
            df_top_sales,
            chart_cfg={
                "type": "bar",
                "title": "Top 15 Productos con Mayor Movimiento",
                "val_cols": [idx_m_top],
                "cat_col": idx_d_top,
                "colors": [HEX_BLUE],
            },
        )

    if df_no_move is not None and not df_no_move.empty:
        cols_nm = list(df_no_move.columns)
        idx_d_nm = cols_nm.index("DESCRIPCION") + 1
        idx_s_nm = cols_nm.index(f"STOCK_{name_b}") + 1
        write_df_sheet_with_chart(
            "5. Top 20 Sin Movimiento",
            HEX_RED,
            df_no_move,
            chart_cfg={
                "type": "bar",
                "title": "Top 15 Productos Inmovilizados (Stock Estancado)",
                "val_cols": [idx_s_nm],
                "cat_col": idx_d_nm,
                "colors": [HEX_RED],
            },
        )

    if df_ai is not None and not df_ai.empty:
        cols_ai = list(df_ai.columns)
        idx_d_ai = cols_ai.index("DESCRIPCION") + 1
        idx_act_ai = cols_ai.index("STOCK_ACTUAL") + 1
        idx_exp_ai = cols_ai.index("EXPECTATIVA_VENTA_IA") + 1
        write_df_sheet_with_chart(
            "6. Proyección Demanda IA",
            HEX_CYAN,
            df_ai,
            chart_cfg={
                "type": "bar",
                "title": "Stock Actual vs Expectativa de Venta Próximo Mes (IA)",
                "val_cols": [idx_act_ai, idx_exp_ai],
                "cat_col": idx_d_ai,
                "colors": [HEX_BLUE, HEX_RED],
            },
        )

    # ==========================================================================
    # PESTAÑA FINAL: 7. Base Power BI (Datos) — TABLA ESTRUCTURADA STAR-SCHEMA
    # ==========================================================================
    if df_reconciled is not None and not df_reconciled.empty:
        ws_bi = wb.create_sheet(title="7. Base Power BI (Datos)")
        ws_bi.sheet_properties.tabColor = HEX_DEEP

        bi_df = pd.DataFrame({
            "Fecha_Corte": datetime.now().strftime("%Y-%m-%d"),
            "Rubro_Sector": sector_label,
            "Fuente_Origen_1": name_a,
            "Fuente_Origen_2": name_b,
            "Codigo_Material": df_reconciled["CODIGO"].astype(str),
            "Descripcion_Material": df_reconciled["DESCRIPCION"].astype(str),
            "Linea_Categoria": df_reconciled["LINEA_CATEGORIA"].astype(str),
            "Stock_Fuente_1": df_reconciled[f"STOCK_{name_a}"].astype(float),
            "Stock_Fuente_2": df_reconciled[f"STOCK_{name_b}"].astype(float),
            "Stock_Consolidado_Total": df_reconciled["STOCK_CONSOLIDADO"].astype(float),
            "Diferencia_Neta": df_reconciled["DIFERENCIA"].astype(float),
            "Variacion_Porcentual": df_reconciled["VARIACION_%"].astype(float),
            "Estado_Auditoria": df_reconciled["ESTADO_CUADRE"].astype(str),
            "Identidad_Color_HEX": np.where(
                df_reconciled["ESTADO_CUADRE"].isin(["Cuadrado", "Creciente (+)"]),
                "#0057FF",
                "#E60026",
            ),
        })

        for c_idx, col_name in enumerate(bi_df.columns, start=1):
            cell = ws_bi.cell(row=1, column=c_idx, value=col_name)
            cell.fill = fill_deep
            cell.font = font_header
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for r_idx, row_vals in enumerate(bi_df.itertuples(index=False), start=2):
            for c_idx, val in enumerate(row_vals, start=1):
                ws_bi.cell(row=r_idx, column=c_idx, value=val)

        max_bi_row = len(bi_df) + 1
        max_bi_col = get_column_letter(len(bi_df.columns))
        bi_table = Table(displayName="TB_STOCK_HIVIMAR_BI", ref=f"A1:{max_bi_col}{max_bi_row}")
        bi_style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        bi_table.tableStyleInfo = bi_style
        ws_bi.add_table(bi_table)

        for col_idx in range(1, len(bi_df.columns) + 1):
            col_letter = get_column_letter(col_idx)
            ws_bi.column_dimensions[col_letter].width = 22

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ==============================================================================
# APLICACIÓN PRINCIPAL STREAMLIT
# ==============================================================================
def main():
    st.markdown(CORPORATE_CSS, unsafe_allow_html=True)

    if "welcome_seen" not in st.session_state:
        st.session_state.welcome_seen = False
    if "confirmed_load" not in st.session_state:
        st.session_state.confirmed_load = False
    if "confirmed_12m" not in st.session_state:
        st.session_state.confirmed_12m = False
    if "file_signature" not in st.session_state:
        st.session_state.file_signature = ""
    if "files12_signature" not in st.session_state:
        st.session_state.files12_signature = ""

    @st.dialog("Bienvenido a STOCK&YA! — Guía Ejecutiva HIVIMAR S.A.", width="large")
    def show_interactive_guide():
        st.markdown(
            """
            ### Plataforma de Consolidación e Inteligencia de Inventarios — **HIVIMAR S.A.**
            Diseñada con estética industrial de alta precisión para unificar reportes de **SAP, Oracle, Odoo, Microsoft Dynamics, CSV o Google Sheets** sin errores de compatibilidad.

            #### Capacidades principales:
            1. **Consolidación Universal (2 Meses, 2 Líneas o Lógico vs Físico):** Carga archivos de distintos meses (ej. **Agosto y Octubre**) o **líneas distintas**. El motor detecta automáticamente filas de encabezado, consolida ubicaciones múltiples y suma/compara sin conflictos.
            2. **Top 20 Rotación y Capital Inmovilizado:** Visualiza los 20 productos de mayor salida y los 20 sin movimiento con porcentajes reales.
            3. **Termómetro Creciente / Decreciente (%):** Analiza variaciones porcentuales por línea de producto y por SKU.
            4. **Consolidado 12 Meses + IA Predictiva:** Carga hasta **12 archivos mensuales** para proyectar la demanda del próximo mes.
            5. **Exportación Excel + Power BI (`&YA!`):** Descarga un libro de Excel con **gráficos nativos integrados**, pestañas claras y tabla estructurada lista para **Power BI**.
            """
        )
        if st.button("Entendido, iniciar plataforma", type="primary", use_container_width=True):
            st.session_state.welcome_seen = True
            st.rerun()

    if not st.session_state.welcome_seen:
        show_interactive_guide()

    with st.sidebar:
        col_logo, col_brand = st.columns([1, 2.2])
        with col_logo:
            st.markdown(HIVIMAR_LOGO_SVG, unsafe_allow_html=True)
        with col_brand:
            st.markdown(
                """
                <div style="padding-top: 6px;">
                  <div style="font-size: 1.52rem; font-weight: 800; color: #FFFFFF; letter-spacing: -0.02em;">
                    STOCK<span style="color:#00E5FF;">&amp;</span><span style="color:#FF1E42;">YA!</span>
                  </div>
                  <div style="font-size: 0.76rem; font-weight: 700; color: #BAE6FD; letter-spacing: 0.08em;">
                    HIVIMAR S.A.
                  </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown("---")
        nav_option = st.radio(
            "MENÚ PRINCIPAL",
            [
                "Inicio y Carga Inteligente",
                "Consolidación y Cuadre",
                "Top 20 Ventas y Sin Movimiento",
                "Creciente y Decreciente (% Real)",
                "Consolidado 12 Meses + IA",
            ],
            label_visibility="collapsed",
        )

        st.markdown("---")
        if st.button("Guía Ejecutiva de Uso", use_container_width=True):
            show_interactive_guide()

        st.markdown(
            """
            <div style="margin-top: 1.5rem; font-size: 0.74rem; color: #E0F2FE; line-height: 1.45;">
              Compatible con Excel (.xlsx/.xls), CSV, Google Sheets y Power BI.<br>
              Motor multi-línea y multi-período activado.
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown(
        """
        <div class="hero-banner">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem;">
            <div>
              <span style="background:rgba(255,255,255,0.16); border:1px solid #00E5FF; color:#FFFFFF; padding:4px 12px; border-radius:999px; font-size:0.75rem; font-weight:800; letter-spacing:0.06em;">
                PLATAFORMA EJECUTIVA DE CONSOLIDACIÓN &bull; HIVIMAR S.A.
              </span>
              <h1 style="margin:0.45rem 0 0.25rem 0; font-size:2rem; font-weight:800; color:#FFFFFF;">
                Bienvenido a STOCK<span style="color:#00E5FF;">&amp;</span><span style="color:#FF1E42;">YA!</span>
              </h1>
              <p style="margin:0; color:#F0F9FF; font-size:0.95rem;">
                Consolida meses distintos, líneas de producto múltiples y sistemas ERP (SAP, Oracle, Odoo, Excel, Google Sheets) con gráficos ejecutivos y exportación lista para Power BI.
              </p>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    df_reconciled = st.session_state.get("df_reconciled", None)
    name_a = st.session_state.get("name_a", "FUENTE_1")
    name_b = st.session_state.get("name_b", "FUENTE_2")
    sector_key = st.session_state.get("sector_key", "automotriz")
    sector_label = st.session_state.get("sector_label", "Automotriz e Industrial (HIVIMAR)")

    if nav_option == "Inicio y Carga Inteligente" or df_reconciled is None:
        render_section_header(
            "upload",
            "Carga Universal de Fuentes (Meses, Líneas o Sistemas ERP)",
            "Soporta archivos con distinta estructura, múltiples líneas de producto y cruce entre períodos (ej. Agosto vs Octubre).",
        )

        op_mode_label = st.radio(
            "Objetivo del análisis:",
            [
                "Consolidación Multi-Línea / Comparativo de Meses (Suma + Variación)",
                "Auditoría y Cuadre Exacto (Sistema Lógico vs Conteo Físico)",
            ],
            horizontal=True,
            key="op_mode_radio",
        )
        operation_mode = "auditoria" if "Auditoría" in op_mode_label else "consolidacion"

        col_src1, col_src2 = st.columns(2)

        sheets_1 = {}
        sheets_2 = {}
        fname_1 = ""
        fname_2 = ""
        sel_sheet_1 = ""
        sel_sheet_2 = ""

        with col_src1:
            render_source_badge_header("01", "Primera Fuente (Ej. Mes Anterior / Línea 1 / SAP)", is_second=False)
            mode_1 = st.radio(
                "Origen Fuente 1:",
                ["Subir Archivo (Excel / CSV)", "Enlace de Google Sheets"],
                horizontal=True,
                key="mode_1",
            )
            if mode_1 == "Subir Archivo (Excel / CSV)":
                up1 = st.file_uploader(
                    "Selecciona el primer archivo",
                    type=["xlsx", "xls", "csv"],
                    key="up1",
                )
                if up1:
                    fname_1 = up1.name
                    sheets_1 = read_uploaded_workbook(up1)
            else:
                gs_url1 = st.text_input("Pega la URL de Google Sheets (Fuente 1):", key="gs1")
                if gs_url1:
                    try:
                        fname_1 = "Google_Sheets_1"
                        sheets_1 = load_google_sheet_from_url(gs_url1)
                    except Exception as exc:
                        st.error(f"Error al leer Google Sheets 1: {exc}")

            if sheets_1:
                options_1 = list(sheets_1.keys())
                if len(options_1) > 1:
                    options_1 = ["CONSOLIDAR TODAS LAS PESTAÑAS"] + options_1
                sel_sheet_1 = st.selectbox(
                    "Pestaña a analizar (Fuente 1):",
                    options_1,
                    index=0,
                    key="sel_s1",
                )
                smart_lbl_1 = derive_smart_source_label(fname_1, sel_sheet_1, "FUENTE_1")
                name_a = sanitize_label(
                    st.text_input("Etiqueta Fuente 1:", value=smart_lbl_1, key=f"lbl_a_{fname_1}_{sel_sheet_1}"),
                    "FUENTE_1",
                )

        with col_src2:
            render_source_badge_header("02", "Segunda Fuente (Ej. Mes Actual / Línea 2 / Conteo Físico)", is_second=True)
            mode_2 = st.radio(
                "Origen Fuente 2:",
                ["Subir Archivo (Excel / CSV)", "Enlace de Google Sheets"],
                horizontal=True,
                key="mode_2",
            )
            if mode_2 == "Subir Archivo (Excel / CSV)":
                up2 = st.file_uploader(
                    "Selecciona el segundo archivo (o el mismo libro para cruzar pestañas)",
                    type=["xlsx", "xls", "csv"],
                    key="up2",
                )
                if up2:
                    fname_2 = up2.name
                    sheets_2 = read_uploaded_workbook(up2)
                elif sheets_1 and len(sheets_1) > 1:
                    fname_2 = fname_1
                    sheets_2 = sheets_1
            else:
                gs_url2 = st.text_input("Pega la URL de Google Sheets (Fuente 2):", key="gs2")
                if gs_url2:
                    try:
                        fname_2 = "Google_Sheets_2"
                        sheets_2 = load_google_sheet_from_url(gs_url2)
                    except Exception as exc:
                        st.error(f"Error al leer Google Sheets 2: {exc}")

            if sheets_2:
                options_2 = list(sheets_2.keys())
                if len(options_2) > 1:
                    options_2 = ["CONSOLIDAR TODAS LAS PESTAÑAS"] + options_2
                default_idx_2 = 2 if (fname_1 == fname_2 and len(options_2) > 2) else (1 if (fname_1 == fname_2 and len(options_2) > 1) else 0)
                sel_sheet_2 = st.selectbox(
                    "Pestaña a analizar (Fuente 2):",
                    options_2,
                    index=default_idx_2,
                    key="sel_s2",
                )
                smart_lbl_2 = derive_smart_source_label(fname_2, sel_sheet_2, "FUENTE_2")
                if smart_lbl_2 == name_a:
                    smart_lbl_2 = f"{smart_lbl_2}_2"[:24]
                name_b = sanitize_label(
                    st.text_input("Etiqueta Fuente 2:", value=smart_lbl_2, key=f"lbl_b_{fname_2}_{sel_sheet_2}"),
                    "FUENTE_2",
                )

        if sheets_1 and sheets_2:
            if name_a == name_b:
                name_a = f"{name_a}_1"[:24]
                name_b = f"{name_b}_2"[:24]

            df_raw_1 = (
                merge_all_workbook_sheets(sheets_1)
                if sel_sheet_1 == "CONSOLIDAR TODAS LAS PESTAÑAS"
                else sheets_1[sel_sheet_1]
            )
            df_raw_2 = (
                merge_all_workbook_sheets(sheets_2)
                if sel_sheet_2 == "CONSOLIDAR TODAS LAS PESTAÑAS"
                else sheets_2[sel_sheet_2]
            )

            sig = f"{fname_1}_{sel_sheet_1}_{len(df_raw_1)}_{fname_2}_{sel_sheet_2}_{len(df_raw_2)}_{name_a}_{name_b}_{operation_mode}"

            if sig != st.session_state.file_signature:
                st.session_state.confirmed_load = False

            if not st.session_state.confirmed_load:
                st.markdown(
                    f"""
                    <div class="aguanta-box">
                      <div style="flex-shrink:0;">{SVG_ICONS["alert"]}</div>
                      <div>
                        <div style="font-size:1.2rem; font-weight:800; color:#E60026;">
                          ¡AGUANTA! ¿Seguro que los archivos a consolidar son los correctos?
                        </div>
                        <div style="color:#0B192C; font-size:0.93rem; margin-top:0.25rem;">
                          Vas a procesar <strong>{name_a}</strong> ({len(df_raw_1):,} registros) junto con <strong>{name_b}</strong> ({len(df_raw_2):,} registros).
                        </div>
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                c_yes, _ = st.columns([1, 2])
                with c_yes:
                    if st.button("¡SIMÓN! Confirmar y Procesar", type="primary", use_container_width=True):
                        st.session_state.confirmed_load = True
                        st.session_state.file_signature = sig
                        st.rerun()
            else:
                detected_filenames = list(dict.fromkeys([fname_1, sel_sheet_1, fname_2, sel_sheet_2]))
                sector_key, sector_label = detect_industry_sector(
                    [df_raw_1, df_raw_2],
                    filenames=detected_filenames,
                )

                if st.session_state.get("last_animated_sig") != sig:
                    anim_placeholder = st.empty()
                    anim_placeholder.markdown(
                        render_dynamic_loader_animation(
                            sector_key,
                            sector_label,
                            analyzed_files=[fname_1, fname_2],
                        ),
                        unsafe_allow_html=True,
                    )
                    time.sleep(1.6)
                    anim_placeholder.empty()
                    st.session_state.last_animated_sig = sig

                schema_1 = auto_detect_schema(df_raw_1)
                schema_2 = auto_detect_schema(df_raw_2)

                with st.expander("Ajustar o verificar columnas detectadas automáticamente (opcional)"):
                    c_det1, c_det2 = st.columns(2)
                    cols_1_opts = ["(Ninguna)"] + [str(c) for c in df_raw_1.columns]
                    cols_2_opts = ["(Ninguna)"] + [str(c) for c in df_raw_2.columns]
                    with c_det1:
                        st.markdown(f"**Columnas {name_a}:**")
                        c1_cod = st.selectbox("Código / Material (F1):", cols_1_opts, index=cols_1_opts.index(schema_1["codigo"]) if schema_1["codigo"] in cols_1_opts else 0, key="m1_cod")
                        c1_desc = st.selectbox("Descripción (F1):", cols_1_opts, index=cols_1_opts.index(schema_1["descripcion"]) if schema_1["descripcion"] in cols_1_opts else 0, key="m1_desc")
                        c1_stk = st.selectbox("Stock / Cantidad (F1):", cols_1_opts, index=cols_1_opts.index(schema_1["stock"]) if schema_1["stock"] in cols_1_opts else 0, key="m1_stk")
                        c1_cat = st.selectbox("Línea / Categoría (F1):", cols_1_opts, index=cols_1_opts.index(schema_1["categoria"]) if schema_1["categoria"] in cols_1_opts else 0, key="m1_cat")
                        schema_1["codigo"] = None if c1_cod == "(Ninguna)" else c1_cod
                        schema_1["descripcion"] = None if c1_desc == "(Ninguna)" else c1_desc
                        schema_1["stock"] = None if c1_stk == "(Ninguna)" else c1_stk
                        schema_1["categoria"] = None if c1_cat == "(Ninguna)" else c1_cat
                    with c_det2:
                        st.markdown(f"**Columnas {name_b}:**")
                        c2_cod = st.selectbox("Código / Material (F2):", cols_2_opts, index=cols_2_opts.index(schema_2["codigo"]) if schema_2["codigo"] in cols_2_opts else 0, key="m2_cod")
                        c2_desc = st.selectbox("Descripción (F2):", cols_2_opts, index=cols_2_opts.index(schema_2["descripcion"]) if schema_2["descripcion"] in cols_2_opts else 0, key="m2_desc")
                        c2_stk = st.selectbox("Stock / Cantidad (F2):", cols_2_opts, index=cols_2_opts.index(schema_2["stock"]) if schema_2["stock"] in cols_2_opts else 0, key="m2_stk")
                        c2_cat = st.selectbox("Línea / Categoría (F2):", cols_2_opts, index=cols_2_opts.index(schema_2["categoria"]) if schema_2["categoria"] in cols_2_opts else 0, key="m2_cat")
                        schema_2["codigo"] = None if c2_cod == "(Ninguna)" else c2_cod
                        schema_2["descripcion"] = None if c2_desc == "(Ninguna)" else c2_desc
                        schema_2["stock"] = None if c2_stk == "(Ninguna)" else c2_stk
                        schema_2["categoria"] = None if c2_cat == "(Ninguna)" else c2_cat

                df_cons_1 = consolidate_source(df_raw_1, schema_1, name_a, default_line_tag=name_a)
                df_cons_2 = consolidate_source(df_raw_2, schema_2, name_b, default_line_tag=name_b)
                df_reconciled = reconcile_two_sources(
                    df_cons_1, df_cons_2, name_a, name_b, operation_mode=operation_mode
                )

                st.session_state.df_reconciled = df_reconciled
                st.session_state.name_a = name_a
                st.session_state.name_b = name_b
                st.session_state.sector_key = sector_key
                st.session_state.sector_label = sector_label

                st.success(
                    f"Consolidación completada con éxito. Rubro detectado: **{sector_label}**. "
                    f"Se unificaron **{len(df_reconciled):,} materiales únicos** entre **{name_a}** y **{name_b}**."
                )

    df_top_sales = None
    df_no_move = None

    if df_reconciled is not None and not df_reconciled.empty:
        col_a = f"STOCK_{name_a}"
        col_b = f"STOCK_{name_b}"

        egr_cols = [c for c in df_reconciled.columns if c.startswith("EGRESOS_")]
        df_mov = df_reconciled.copy()

        if egr_cols:
            df_mov["UNIDADES_MOVIDAS"] = df_mov[egr_cols[0]]
            mov_label = "Unidades Vendidas / Egresadas"
        else:
            df_mov["UNIDADES_MOVIDAS"] = (df_mov[col_a] - df_mov[col_b]).clip(lower=0.0)
            mov_label = f"Salida Neta ({name_a} a {name_b})"

        total_moved = float(df_mov["UNIDADES_MOVIDAS"].sum())
        df_top_sales = (
            df_mov[df_mov["UNIDADES_MOVIDAS"] > 0]
            .sort_values("UNIDADES_MOVIDAS", ascending=False)
            .head(20)
            .copy()
        )
        if not df_top_sales.empty and total_moved > 0:
            df_top_sales["PARTICIPACION_%"] = (df_top_sales["UNIDADES_MOVIDAS"] / total_moved * 100.0).round(2)

        total_stock_b = float(df_mov[col_b].sum())
        df_no_move = (
            df_mov[(df_mov["UNIDADES_MOVIDAS"] == 0) & (df_mov[col_b] > 0)]
            .sort_values(col_b, ascending=False)
            .head(20)
            .copy()
        )
        if not df_no_move.empty and total_stock_b > 0:
            df_no_move["PESO_INVENTARIO_%"] = (df_no_move[col_b] / total_stock_b * 100.0).round(2)

        total_skus = len(df_reconciled)
        cuadrados_cnt = int((df_reconciled["ESTADO_CUADRE"] == "Cuadrado").sum())
        descuadrados_cnt = total_skus - cuadrados_cnt
        exactitud_pct = (cuadrados_cnt / total_skus * 100.0) if total_skus > 0 else 0.0

        sum_a = float(df_reconciled[col_a].sum())
        sum_b = float(df_reconciled[col_b].sum())
        sum_cons = float(df_reconciled["STOCK_CONSOLIDADO"].sum())
        var_inv_pct = ((sum_b - sum_a) / sum_a * 100.0) if sum_a > 0 else 0.0

        c_info, c_exp1, c_exp2 = st.columns([2.4, 1.3, 1.1])
        with c_info:
            st.caption(
                f"**Consolidado Activo:** `{name_a}` + `{name_b}`  |  **Rubro:** {sector_label}  |  "
                f"**Materiales Únicos:** {total_skus:,}"
            )
        with c_exp1:
            excel_bytes = build_master_excel_export(
                df_reconciled=df_reconciled,
                name_a=name_a,
                name_b=name_b,
                df_top_sales=df_top_sales,
                df_no_move=df_no_move,
                df_ai=st.session_state.get("df_ai", None),
                sector_label=sector_label,
            )
            st.download_button(
                label="Descargar Excel + Gráficos &YA!",
                data=excel_bytes,
                file_name=f"STOCK_Y_YA_HIVIMAR_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with c_exp2:
            st.download_button(
                label="Tema Power BI (.JSON)",
                data=build_powerbi_theme_json(),
                file_name="Tema_PowerBI_HIVIMAR_MegaManX.json",
                mime="application/json",
                use_container_width=True,
            )

        if nav_option in ["Inicio y Carga Inteligente", "Consolidación y Cuadre"]:
            render_section_header(
                "audit",
                "Consolidación Multi-Línea / Período y Auditoría Exacta",
                "Suma consolidada, comparativo entre fuentes y detección de variaciones.",
            )

            k1, k2, k3, k4 = st.columns(4)
            with k1:
                st.markdown(
                    f"""
                    <div class="kpi-card">
                      <div class="kpi-label">Stock Consolidado Total</div>
                      <div class="kpi-value">{sum_cons:,.0f} uds</div>
                      <div class="kpi-sub badge-ok">{name_a}: {sum_a:,.0f} + {name_b}: {sum_b:,.0f}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with k2:
                st.markdown(
                    f"""
                    <div class="kpi-card cyan-top">
                      <div class="kpi-label">Coincidencia / Exactitud</div>
                      <div class="kpi-value">{exactitud_pct:.2f}%</div>
                      <div class="kpi-sub badge-cyan">{cuadrados_cnt:,} de {total_skus:,} SKUs idénticos</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with k3:
                st.markdown(
                    f"""
                    <div class="kpi-card red-top">
                      <div class="kpi-label">Variación Neta ({name_b} - {name_a})</div>
                      <div class="kpi-value">{sum_b - sum_a:+,.0f} uds</div>
                      <div class="kpi-sub {'badge-ok' if (sum_b - sum_a) >= 0 else 'badge-err'}">
                        {var_inv_pct:+.2f}% entre fuentes
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with k4:
                st.markdown(
                    f"""
                    <div class="kpi-card">
                      <div class="kpi-label">SKUs con Novedad / Exclusivos</div>
                      <div class="kpi-value">{descuadrados_cnt:,}</div>
                      <div class="kpi-sub badge-err">Diferencias o líneas exclusivas</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            st.markdown("<br>", unsafe_allow_html=True)

            # Gráfico Ejecutivo Comparativo en la vista principal
            fig_overview = go.Figure()
            fig_overview.add_trace(
                go.Bar(
                    x=[name_a, name_b, "CONSOLIDADO TOTAL"],
                    y=[sum_a, sum_b, sum_cons],
                    text=[f"{sum_a:,.0f}", f"{sum_b:,.0f}", f"{sum_cons:,.0f}"],
                    textposition="outside",
                    marker_color=["#0057FF", "#00C8FF", "#E60026"],
                )
            )
            fig_overview.update_layout(
                title=f"Balance General de Unidades: {name_a} vs {name_b} vs Consolidado",
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#F4F8FF",
                font=dict(color="#06193E"),
                height=340,
                yaxis_title="Unidades",
                showlegend=False,
            )
            st.plotly_chart(fig_overview, use_container_width=True)

            f_col1, f_col2 = st.columns([1.5, 2])
            with f_col1:
                filtro_estado = st.selectbox(
                    "Filtrar registros:",
                    [
                        "Todos los materiales consolidados",
                        "Solo con Variación / Diferencia",
                        "Solo Cuadrados / Idénticos",
                        f"Solo en {name_a}",
                        f"Solo en {name_b}",
                        "Materiales Multi-Ubicación (Agrupados)",
                    ],
                )
            with f_col2:
                busqueda_sku = st.text_input("Buscar por Código, Descripción o Línea:", placeholder="Ej. 1001747, SKF, Automotriz...")

            df_view = df_reconciled.copy()
            if filtro_estado == "Solo con Variación / Diferencia":
                df_view = df_view[df_view["ESTADO_CUADRE"] != "Cuadrado"]
            elif filtro_estado == "Solo Cuadrados / Idénticos":
                df_view = df_view[df_view["ESTADO_CUADRE"] == "Cuadrado"]
            elif filtro_estado == f"Solo en {name_a}":
                df_view = df_view[df_view["ESTADO_CUADRE"] == f"Solo en {name_a}"]
            elif filtro_estado == f"Solo en {name_b}":
                df_view = df_view[df_view["ESTADO_CUADRE"] == f"Solo en {name_b}"]
            elif filtro_estado == "Materiales Multi-Ubicación (Agrupados)":
                reg_cols = [c for c in df_view.columns if c.startswith("REGISTROS_")]
                cond = pd.Series(False, index=df_view.index)
                for rc in reg_cols:
                    cond = cond | (df_view[rc] > 1)
                df_view = df_view[cond]

            if busqueda_sku.strip():
                q = busqueda_sku.strip().lower()
                df_view = df_view[
                    df_view["CODIGO"].str.lower().str.contains(q, na=False)
                    | df_view["DESCRIPCION"].str.lower().str.contains(q, na=False)
                    | df_view["LINEA_CATEGORIA"].str.lower().str.contains(q, na=False)
                ]

            st.dataframe(df_view, use_container_width=True, height=430)

        if nav_option == "Top 20 Ventas y Sin Movimiento":
            render_section_header(
                "ranking",
                "Ranking de Rotación y Capital Inmovilizado",
                "Participación porcentual real sobre unidades movidas e inventario estancado.",
            )
            tab_top, tab_dead = st.tabs(["Top 20 Más Vendidos / Mayor Salida", "Top 20 Sin Movimiento (Inmovilizado)"])

            with tab_top:
                if df_top_sales is not None and not df_top_sales.empty:
                    fig_top = px.bar(
                        df_top_sales.sort_values("UNIDADES_MOVIDAS", ascending=True),
                        x="UNIDADES_MOVIDAS",
                        y="DESCRIPCION",
                        orientation="h",
                        text="PARTICIPACION_%",
                        color="UNIDADES_MOVIDAS",
                        color_continuous_scale=["#00C8FF", "#0057FF", "#06193E"],
                        title=f"Top 20 Productos — {mov_label}",
                    )
                    fig_top.update_traces(texttemplate="%{x:,.0f} uds (%{text}%)", textposition="outside")
                    fig_top.update_layout(
                        paper_bgcolor="#FFFFFF",
                        plot_bgcolor="#F4F8FF",
                        font=dict(color="#06193E"),
                        height=620,
                        showlegend=False,
                        coloraxis_showscale=False,
                        xaxis_title="Unidades",
                        yaxis_title="",
                    )
                    st.plotly_chart(fig_top, use_container_width=True)
                    st.dataframe(
                        df_top_sales[["CODIGO", "DESCRIPCION", "LINEA_CATEGORIA", col_a, col_b, "UNIDADES_MOVIDAS", "PARTICIPACION_%"]],
                        use_container_width=True,
                    )
                else:
                    st.info("No se registraron egresos/salidas mayores a cero entre los archivos seleccionados.")

            with tab_dead:
                if df_no_move is not None and not df_no_move.empty:
                    total_dead_units = float(df_mov[(df_mov["UNIDADES_MOVIDAS"] == 0) & (df_mov[col_b] > 0)][col_b].sum())
                    dead_pct = (total_dead_units / sum_b * 100.0) if sum_b > 0 else 0.0
                    st.warning(
                        f"El **{dead_pct:.2f}%** del inventario actual ({total_dead_units:,.0f} unidades) "
                        f"no registró rotación en el período analizado."
                    )
                    fig_dead = px.bar(
                        df_no_move.sort_values(col_b, ascending=True),
                        x=col_b,
                        y="DESCRIPCION",
                        orientation="h",
                        text="PESO_INVENTARIO_%",
                        color=col_b,
                        color_continuous_scale=["#0057FF", "#E60026"],
                        title="Top 20 Productos Sin Movimiento (Mayor Stock Estancado)",
                    )
                    fig_dead.update_traces(texttemplate="%{x:,.0f} uds (%{text}%)", textposition="outside")
                    fig_dead.update_layout(
                        paper_bgcolor="#FFFFFF",
                        plot_bgcolor="#F4F8FF",
                        font=dict(color="#06193E"),
                        height=620,
                        coloraxis_showscale=False,
                        xaxis_title="Stock Estancado (Unidades)",
                        yaxis_title="",
                    )
                    st.plotly_chart(fig_dead, use_container_width=True)
                    st.dataframe(
                        df_no_move[["CODIGO", "DESCRIPCION", "LINEA_CATEGORIA", col_b, "PESO_INVENTARIO_%", "ESTADO_CUADRE"]],
                        use_container_width=True,
                    )
                else:
                    st.success("Todos los productos con stock registraron movimiento en el período.")

        if nav_option == "Creciente y Decreciente (% Real)":
            render_section_header(
                "trends",
                "Análisis Porcentual Creciente y Decreciente",
                "Variaciones porcentuales reales por línea de producto y por SKU.",
            )

            c_var1, c_var2 = st.columns(2)
            with c_var1:
                st.markdown(
                    f"""
                    <div class="kpi-card">
                      <div class="kpi-label">Variación Neta de Inventario ({name_a} vs {name_b})</div>
                      <div class="kpi-value">{var_inv_pct:+.2f}%</div>
                      <div class="kpi-sub {'badge-ok' if var_inv_pct >= 0 else 'badge-err'}">
                        De {sum_a:,.0f} uds a {sum_b:,.0f} uds ({sum_b - sum_a:+,.0f} uds)
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            with c_var2:
                egr_a_col = f"EGRESOS_{name_a}"
                egr_b_col = f"EGRESOS_{name_b}"
                if egr_a_col in df_reconciled.columns and egr_b_col in df_reconciled.columns:
                    v_a = float(df_reconciled[egr_a_col].sum())
                    v_b = float(df_reconciled[egr_b_col].sum())
                    v_pct = ((v_b - v_a) / v_a * 100.0) if v_a > 0 else 0.0
                    st.markdown(
                        f"""
                        <div class="kpi-card cyan-top">
                          <div class="kpi-label">Variación Porcentual de Ventas / Egresos</div>
                          <div class="kpi-value">{v_pct:+.2f}%</div>
                          <div class="kpi-sub {'badge-ok' if v_pct >= 0 else 'badge-err'}">
                            De {v_a:,.0f} uds a {v_b:,.0f} uds ({v_b - v_a:+,.0f} uds)
                          </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                else:
                    skus_subieron = int((df_reconciled["DIFERENCIA"] > 0).sum())
                    skus_bajaron = int((df_reconciled["DIFERENCIA"] < 0).sum())
                    pct_sub = skus_subieron / total_skus * 100.0
                    pct_baj = skus_bajaron / total_skus * 100.0
                    st.markdown(
                        f"""
                        <div class="kpi-card cyan-top">
                          <div class="kpi-label">Dinámica del Catálogo (SKUs en Alza vs Baja)</div>
                          <div class="kpi-value">+{pct_sub:.1f}% / -{pct_baj:.1f}%</div>
                          <div class="kpi-sub badge-cyan">
                            {skus_subieron:,} SKUs crecieron &bull; {skus_bajaron:,} SKUs decrecieron
                          </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

            st.markdown("<br>", unsafe_allow_html=True)

            if "LINEA_CATEGORIA" in df_reconciled.columns:
                cat_grp = (
                    df_reconciled.groupby("LINEA_CATEGORIA", as_index=False)[[col_a, col_b]]
                    .sum()
                )
                cat_grp["VARIACION_%"] = np.where(
                    cat_grp[col_a] > 0,
                    (cat_grp[col_b] - cat_grp[col_a]) / cat_grp[col_a] * 100.0,
                    0.0,
                ).round(2)
                cat_grp["TENDENCIA"] = np.where(cat_grp["VARIACION_%"] >= 0, "Creciente (+)", "Decreciente (-)")

                fig_cat = px.bar(
                    cat_grp.sort_values("VARIACION_%"),
                    x="VARIACION_%",
                    y="LINEA_CATEGORIA",
                    orientation="h",
                    color="TENDENCIA",
                    color_discrete_map={"Creciente (+)": "#0057FF", "Decreciente (-)": "#E60026"},
                    text="VARIACION_%",
                    title="Variación Porcentual de Inventario por Línea / Categoría (%)",
                )
                fig_cat.update_traces(texttemplate="%{text:+.2f}%", textposition="outside")
                fig_cat.update_layout(
                    paper_bgcolor="#FFFFFF",
                    plot_bgcolor="#F4F8FF",
                    font=dict(color="#06193E"),
                    height=420,
                    xaxis_title="Variación Porcentual (%)",
                    yaxis_title="",
                )
                st.plotly_chart(fig_cat, use_container_width=True)

            df_pct_sku = df_reconciled[df_reconciled[col_a] > 0].copy()
            df_pct_sku["VAR_%"] = ((df_pct_sku[col_b] - df_pct_sku[col_a]) / df_pct_sku[col_a] * 100.0).round(2)
            top_up = df_pct_sku.sort_values("VAR_%", ascending=False).head(10)
            top_down = df_pct_sku.sort_values("VAR_%", ascending=True).head(10)
            df_movers = pd.concat([top_up, top_down]).drop_duplicates("CODIGO")
            df_movers["DIRECCION"] = np.where(df_movers["VAR_%"] >= 0, "Creciente (+)", "Decreciente (-)")

            fig_movers = px.bar(
                df_movers.sort_values("VAR_%"),
                x="VAR_%",
                y="DESCRIPCION",
                orientation="h",
                color="DIRECCION",
                color_discrete_map={"Creciente (+)": "#0057FF", "Decreciente (-)": "#E60026"},
                text="VAR_%",
                title="Top Productos con Mayor Variación Porcentual (Creciente vs Decreciente)",
            )
            fig_movers.update_traces(texttemplate="%{text:+.1f}%", textposition="outside")
            fig_movers.update_layout(
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#F4F8FF",
                font=dict(color="#06193E"),
                height=560,
                xaxis_title="Variación %",
                yaxis_title="",
            )
            st.plotly_chart(fig_movers, use_container_width=True)

    if nav_option == "Consolidado 12 Meses + IA":
        render_section_header(
            "ai",
            "Consolidado Histórico (Hasta 12 Meses) y Proyección de Demanda con IA",
            "Ordenamiento cronológico automático, consolidación multi-ubicación y modelado predictivo de demanda.",
        )
        st.markdown(
            "Sube hasta **12 archivos mensuales** (Excel o CSV). El motor los ordenará cronológicamente, "
            "consolidará ubicaciones repetidas y calculará la **expectativa de venta del próximo mes**."
        )

        uploaded_12 = st.file_uploader(
            "Selecciona los archivos mensuales (hasta 12 archivos):",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key="up_12m",
        )

        if uploaded_12:
            sig12 = "_".join(sorted(f.name for f in uploaded_12))
            if sig12 != st.session_state.files12_signature:
                st.session_state.confirmed_12m = False

            if not st.session_state.confirmed_12m:
                st.markdown(
                    f"""
                    <div class="aguanta-box">
                      <div style="flex-shrink:0;">{SVG_ICONS["alert"]}</div>
                      <div>
                        <div style="font-size:1.2rem; font-weight:800; color:#E60026;">
                          ¡AGUANTA! ¿Seguro que los archivos mensuales son los correctos?
                        </div>
                        <div style="color:#0B192C; font-size:0.9rem; margin-top:0.25rem;">
                          Has seleccionado <strong>{len(uploaded_12)} archivos mensuales</strong> para consolidación anual y proyección IA.
                        </div>
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                if st.button("¡SIMÓN! Procesar Serie de 12 Meses", type="primary"):
                    st.session_state.confirmed_12m = True
                    st.session_state.files12_signature = sig12
                    st.rerun()
            else:
                sorted_files = sorted(
                    enumerate(uploaded_12[:12]),
                    key=lambda pair: extract_period_sort_key(pair[1].name, pair[0]),
                )

                monthly_records = []
                sample_dfs = []
                file_names_12 = []
                for idx, f_obj in sorted_files:
                    _, _, period_lbl = extract_period_sort_key(f_obj.name, idx)
                    file_names_12.append(f_obj.name)
                    s_dict = read_uploaded_workbook(f_obj)
                    raw_m = merge_all_workbook_sheets(s_dict) if len(s_dict) > 1 else list(s_dict.values())[0]
                    sample_dfs.append(raw_m)
                    sch_m = auto_detect_schema(raw_m)
                    cons_m = consolidate_source(raw_m, sch_m, "M")
                    monthly_records.append({"label": period_lbl, "df": cons_m})

                sec_k12, sec_l12 = detect_industry_sector(sample_dfs, filenames=file_names_12)
                if st.session_state.get("last_12m_anim") != sig12:
                    ph12 = st.empty()
                    ph12.markdown(
                        render_dynamic_loader_animation(
                            sec_k12,
                            sec_l12,
                            analyzed_files=file_names_12[:3],
                        ),
                        unsafe_allow_html=True,
                    )
                    time.sleep(1.5)
                    ph12.empty()
                    st.session_state.last_12m_anim = sig12

                df_ai, piv_sales = run_ai_demand_forecast(monthly_records)
                st.session_state.df_ai = df_ai

                tot_forecast = float(df_ai["EXPECTATIVA_VENTA_IA"].sum())
                urgent_cnt = int(df_ai["DIAGNOSTICO_IA"].str.contains("Reabastecer").sum())
                accel_cnt = int(df_ai["DIAGNOSTICO_IA"].str.contains("Aceleración").sum())

                a1, a2, a3 = st.columns(3)
                with a1:
                    st.markdown(
                        f"""
                        <div class="kpi-card">
                          <div class="kpi-label">Expectativa Total Próximo Mes (IA)</div>
                          <div class="kpi-value">{tot_forecast:,.0f} uds</div>
                          <div class="kpi-sub badge-ok">Proyección sobre {len(monthly_records)} meses</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with a2:
                    st.markdown(
                        f"""
                        <div class="kpi-card red-top">
                          <div class="kpi-label">Alertas de Quiebre de Stock</div>
                          <div class="kpi-value">{urgent_cnt:,} SKUs</div>
                          <div class="kpi-sub badge-err">Stock actual menor a expectativa IA</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with a3:
                    st.markdown(
                        f"""
                        <div class="kpi-card cyan-top">
                          <div class="kpi-label">Productos en Aceleración</div>
                          <div class="kpi-value">{accel_cnt:,} SKUs</div>
                          <div class="kpi-sub badge-cyan">Crecimiento proyectado &ge; +10%</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                st.markdown("<br>", unsafe_allow_html=True)

                monthly_totals = piv_sales.sum(axis=0)
                hist_x = list(monthly_totals.index)
                hist_y = list(monthly_totals.values)

                fig_ai = go.Figure()
                fig_ai.add_trace(
                    go.Scatter(
                        x=hist_x,
                        y=hist_y,
                        mode="lines+markers+text",
                        name="Histórico Real",
                        text=[f"{v:,.0f}" for v in hist_y],
                        textposition="top center",
                        line=dict(color="#0057FF", width=3.5),
                        marker=dict(size=9, color="#00C8FF", line=dict(width=2, color="#0057FF")),
                    )
                )
                fig_ai.add_trace(
                    go.Scatter(
                        x=[hist_x[-1], "Próximo Mes (IA)"],
                        y=[hist_y[-1], tot_forecast],
                        mode="lines+markers+text",
                        name="Proyección IA",
                        text=["", f"{tot_forecast:,.0f} uds"],
                        textposition="top center",
                        line=dict(color="#E60026", width=3.5, dash="dash"),
                        marker=dict(size=12, color="#E60026", symbol="diamond"),
                    )
                )
                fig_ai.update_layout(
                    title="Curva Histórica Consolidada y Expectativa de Venta Próximo Mes (IA)",
                    paper_bgcolor="#FFFFFF",
                    plot_bgcolor="#F4F8FF",
                    font=dict(color="#06193E"),
                    height=390,
                    yaxis_title="Unidades",
                )
                st.plotly_chart(fig_ai, use_container_width=True)

                render_section_header(
                    "table",
                    "Tabla de Expectativa de Venta y Recomendación de Compra por Producto",
                )
                st.dataframe(df_ai, use_container_width=True, height=400)

    st.markdown(
        """
        <div class="copyright-footer">
          STOCK&amp;YA! &bull; Plataforma de Inteligencia de Inventarios para HIVIMAR S.A. &nbsp;|&nbsp;
          Derechos de autor reservados por John Aarón Domínguez Jara 01.593.321
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
