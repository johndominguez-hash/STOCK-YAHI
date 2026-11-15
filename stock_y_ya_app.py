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
    page_title="STOCK&YA! | Inteligencia Gerencial de Inventarios - HIVIMAR S.A.",
    page_icon=PAGE_FAVICON_SVG,
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==============================================================================
# PALETA DE IDENTIDAD OSCURA DE ALTO CONTRASTE + ACENTOS VIBRANTES
# ==============================================================================
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

/* Fondo general Oscuro Profundo para máximo contraste */
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

/* Espacio superior seguro para que el encabezado fijo nunca tape contenido */
.block-container {
    padding-top: 2.6rem !important;
    padding-bottom: 2.5rem !important;
}

/* Diálogos / Modales en oscuro */
[data-testid="stDialog"] [role="dialog"],
[role="dialog"] {
    background-color: #0D1E36 !important;
    color: #FFFFFF !important;
    border: 1.5px solid #00E5FF !important;
    border-radius: 16px !important;
}
[role="dialog"] * {
    color: #F8FAFC !important;
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

/* Legibilidad garantizada para todos los widgets en modo oscuro */
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
    padding: 14px !important;
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
    font-size: 1.72rem;
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
    padding: 1.35rem 1.85rem;
    margin-bottom: 1.15rem;
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

/* Tarjeta de Acción Gerencial */
.exec-advice-card {
    background: linear-gradient(135deg, #0B192E 0%, #102542 100%);
    border: 1px solid rgba(0, 229, 255, 0.32);
    border-left: 5px solid #00E5FF;
    border-radius: 12px;
    padding: 1rem 1.2rem;
    margin-bottom: 0.85rem;
}
.exec-advice-card.red-accent {
    border-left-color: #FF3B5C;
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
    "executive": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>""",
    "ai": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="16" height="16" rx="2"/><rect x="9" y="9" width="6" height="6"/><line x1="9" y1="1" x2="9" y2="4"/><line x1="15" y1="1" x2="15" y2="4"/><line x1="9" y1="20" x2="9" y2="23"/><line x1="15" y1="20" x2="15" y2="23"/><line x1="20" y1="9" x2="23" y2="9"/><line x1="20" y1="14" x2="23" y2="14"/><line x1="1" y1="9" x2="4" y2="9"/><line x1="1" y1="14" x2="4" y2="14"/></svg>""",
    "table": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="9" x2="9" y2="21"/></svg>""",
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
        margin=dict(l=20, r=35, t=52, b=28),
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
                border: 2px solid #00E5FF; border-radius: 16px; padding: 1.2rem 1.6rem;
                margin: 0.75rem 0; text-align: center; box-shadow: 0 10px 28px rgba(0,87,255,0.35);">
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

      <div style="font-size: 1.25rem; font-weight: 800; color: #FFFFFF; letter-spacing: 0.04em; margin-bottom: 0.25rem;">
        CONSOLIDANDO Y AUDITANDO INVENTARIOS...
      </div>
      <div style="font-size: 0.85rem; color: #E0F2FE; margin-bottom: 0.65rem;">
        {files_caption}Rubro identificado automáticamente: <strong style="color:#00E5FF;">{sector_label}</strong>
      </div>

      <svg width="100%" height="68" viewBox="0 0 420 68" style="max-width: 500px; overflow: visible;">
        <g class="wink-group">
          <circle cx="210" cy="34" r="24" fill="#06193E" stroke="#00E5FF" stroke-width="3"/>
          <circle cx="201" cy="28" r="3.5" fill="#FFFFFF"/>
          <path d="M 215 28 Q 221 23 226 28" stroke="#00E5FF" stroke-width="3.2" stroke-linecap="round" fill="none"/>
          <path d="M 198 41 Q 210 51 222 41" stroke="#FFFFFF" stroke-width="3" stroke-linecap="round" fill="none"/>
        </g>

        <g class="track-group">
          <line x1="15" y1="54" x2="405" y2="54" stroke="#00E5FF" stroke-width="3" stroke-linecap="round"/>
          <line class="track-line" x1="15" y1="60" x2="405" y2="60" stroke="#FF3B5C" stroke-width="2.5"/>
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
        "datos", "datos_csv", "inventario", "reporte", "export", "tabla1", "base",
        "consolidar todas las pestanas"
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
# MOTOR GERENCIAL ABC + VALORIZACIÓN DE CAPITAL Y SALUD DE INVENTARIO
# ==============================================================================
def compute_managerial_abc_analysis(
    df_reconciled: pd.DataFrame,
    name_a: str,
    name_b: str,
    fallback_unit_cost: float = 15.0,
) -> tuple:
    """
    Calcula la Matriz Gerencial ABC (Pareto 80/15/5), Valorización Económica ($ USD),
    Capital Inmovilizado y el Índice Ejecutivo de Salud de Inventario (0-100 pts).
    """
    df = df_reconciled.copy()
    col_a = f"STOCK_{name_a}"
    col_b = f"STOCK_{name_b}"

    price_cols = [c for c in df.columns if c.startswith("PRECIO_")]
    if price_cols:
        unit_cost = df[price_cols].mean(axis=1).replace(0, np.nan).fillna(fallback_unit_cost)
        has_real_price = True
    else:
        unit_cost = pd.Series(fallback_unit_cost, index=df.index)
        has_real_price = False

    df["COSTO_UNIT_USD"] = unit_cost.round(2)
    df["VALOR_CONSOLIDADO_USD"] = (df["STOCK_CONSOLIDADO"] * df["COSTO_UNIT_USD"]).round(2)
    df["VALOR_FUENTE_B_USD"] = (df[col_b] * df["COSTO_UNIT_USD"]).round(2)
    df["IMPACTO_DIFERENCIA_USD"] = (df["DIFERENCIA"].abs() * df["COSTO_UNIT_USD"]).round(2)

    df = df.sort_values("VALOR_CONSOLIDADO_USD", ascending=False).reset_index(drop=True)
    total_val = float(df["VALOR_CONSOLIDADO_USD"].sum())

    if total_val > 0:
        df["PARTICIPACION_VALOR_%"] = (df["VALOR_CONSOLIDADO_USD"] / total_val * 100.0).round(2)
        df["ACUMULADO_PARETO_%"] = (df["VALOR_CONSOLIDADO_USD"].cumsum() / total_val * 100.0).round(2)
    else:
        df["PARTICIPACION_VALOR_%"] = 0.0
        df["ACUMULADO_PARETO_%"] = 0.0

    df["CLASE_ABC"] = np.where(
        df["ACUMULADO_PARETO_%"] <= 80.0,
        "A (Crítico 80%)",
        np.where(df["ACUMULADO_PARETO_%"] <= 95.0, "B (Intermedio 15%)", "C (Cola 5%)"),
    )

    egr_cols = [c for c in df.columns if c.startswith("EGRESOS_")]
    if egr_cols:
        mov_units = df[egr_cols[0]]
    else:
        mov_units = (df[col_a] - df[col_b]).clip(lower=0.0)

    def get_managerial_action(row, m_u):
        is_abc_a = row["CLASE_ABC"].startswith("A")
        has_diff = row["ESTADO_CUADRE"] != "Cuadrado"
        if is_abc_a and has_diff:
            return "[PRIORIDAD 1] Auditar Diferencia en SKU Clase A"
        if is_abc_a and m_u == 0 and row[col_b] > 0:
            return "[CAPITAL DETENIDO] Impulsar Venta / Liquidar Clase A"
        if has_diff:
            return "[CONTROL] Regularizar Ajuste de Inventario"
        if m_u > 0:
            return "[SALUDABLE] Rotación Activa y Cuadrado"
        return "[MONITOREAR] Stock Estable Sin Egresos"

    df["ACCION_GERENCIAL"] = [
        get_managerial_action(row, m_u)
        for (_, row), m_u in zip(df.iterrows(), mov_units)
    ]

    total_skus = len(df)
    exact_rate = float((df["ESTADO_CUADRE"] == "Cuadrado").sum()) / max(total_skus, 1)
    active_rate = float((mov_units > 0).sum()) / max(total_skus, 1)
    abc_a_exact = float(
        (df[df["CLASE_ABC"].str.startswith("A")]["ESTADO_CUADRE"] == "Cuadrado").sum()
    ) / max(int(df["CLASE_ABC"].str.startswith("A").sum()), 1)

    health_score = round((exact_rate * 45.0 + abc_a_exact * 35.0 + max(active_rate, 0.35) * 20.0), 1)
    health_score = min(100.0, max(0.0, health_score))

    summary_metrics = {
        "has_real_price": has_real_price,
        "total_val_usd": total_val,
        "val_b_usd": float(df["VALOR_FUENTE_B_USD"].sum()),
        "diff_val_usd": float(df["IMPACTO_DIFERENCIA_USD"].sum()),
        "class_a_skus": int(df["CLASE_ABC"].str.startswith("A").sum()),
        "class_a_usd": float(df[df["CLASE_ABC"].str.startswith("A")]["VALOR_CONSOLIDADO_USD"].sum()),
        "immobilized_a_usd": float(
            df[df["ACCION_GERENCIAL"].str.contains("CAPITAL DETENIDO")]["VALOR_FUENTE_B_USD"].sum()
        ),
        "health_score": health_score,
    }
    return df, summary_metrics


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

    if not has_explicit_sales and n_periods >= 2:
        active_sales = sales_mat[:, 1:]
    else:
        active_sales = sales_mat

    t_act = active_sales.shape[1]

    if t_act >= 4:
        p90 = np.percentile(active_sales, 90, axis=1, keepdims=True)
        med = np.median(active_sales, axis=1, keepdims=True)
        upper_cap = np.maximum(p90 * 1.5, med * 3.0)
        clean_sales = np.minimum(active_sales, upper_cap)
    else:
        clean_sales = active_sales.copy()

    avg_monthly = np.mean(clean_sales, axis=1)

    if t_act >= 2:
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

        alpha, beta = 0.45, 0.25
        level = clean_sales[:, 0].copy()
        trend = clean_sales[:, 1] - clean_sales[:, 0]
        for t in range(1, t_act):
            prev_level = level.copy()
            level = alpha * clean_sales[:, t] + (1.0 - alpha) * (level + trend)
            trend = beta * (level - prev_level) + (1.0 - beta) * trend
        holt_pred = np.maximum(0.0, level + trend)

        forecast = np.round(0.50 * reg_pred + 0.50 * holt_pred, 1)

        if t_act >= 3:
            dormant_mask = np.all(clean_sales[:, -3:] == 0, axis=1)
            forecast = np.where(dormant_mask, 0.0, forecast)

        growth_pct = np.where(avg_monthly > 0, (forecast - avg_monthly) / avg_monthly * 100.0, 0.0)
    else:
        forecast = np.round(last_month_sales, 1)
        growth_pct = np.zeros(n_skus, dtype=float)

    coverage_months = np.where(
        forecast > 0,
        np.minimum(current_stock / np.maximum(forecast, 1e-6), 99.0),
        np.where(current_stock > 0, 99.0, 0.0),
    )

    target_stock = np.ceil(forecast * 1.15)
    suggested_buy = np.where(
        (forecast > 0) & (current_stock < target_stock),
        target_stock - current_stock,
        0.0,
    )

    std_sales = np.std(clean_sales, axis=1)
    cv = np.where(avg_monthly > 0, std_sales / (avg_monthly + 1e-6), 1.0)
    stability = np.clip(1.0 - 0.45 * cv, 0.35, 0.98)
    history_factor = min(0.65 + 0.06 * t_act, 1.0)
    confidence_pct = np.round(stability * history_factor * 100.0, 1)

    conds = [
        (forecast > 0) & (current_stock < forecast),
        (growth_pct >= 10.0) & (forecast > 0),
        (forecast == 0) & (current_stock > 0),
        (growth_pct <= -10.0) & (forecast > 0),
    ]
    choices = [
        "[CRÍTICO] Reabastecer Urgente (Riesgo Quiebre)",
        "[ALZA] Demanda en Aceleración",
        "[INMOVILIZADO] Sin Rotación",
        "[BAJA] Demanda en Desaceleración",
    ]
    diagnostics = np.select(conds, choices, default="[ÓPTIMO] Demanda Estable")

    df_ai = pd.DataFrame({
        "CODIGO": codes,
        "DESCRIPCION": descs,
        "STOCK_ACTUAL": current_stock,
        "PROMEDIO_MENSUAL": np.round(avg_monthly, 1),
        "ULTIMO_MES": np.round(last_month_sales, 1),
        "EXPECTATIVA_VENTA_IA": forecast,
        "SUGERIDO_COMPRA_IA": suggested_buy,
        "VARIACION_ESTIMADA_%": np.round(growth_pct, 1),
        "COBERTURA_MESES": np.round(coverage_months, 1),
        "CONFIANZA_IA_%": confidence_pct,
        "DIAGNOSTICO_IA": diagnostics,
    }).sort_values(["EXPECTATIVA_VENTA_IA", "SUGERIDO_COMPRA_IA"], ascending=[False, False]).reset_index(drop=True)

    return df_ai, piv_sales


# ==============================================================================
# EXPORTADOR MAESTRO EXCEL (&YA!) — CERO SOLAPAMIENTO, ETIQUETAS DE PRODUCTO
# Y VALORES VISIBLES EN CADA GRÁFICO + BASE POWER BI
# ==============================================================================
def configure_legible_bar_chart(
    chart: BarChart,
    title: str,
    x_title: str,
    y_title: str,
    palette_hex_list: list,
    width: float = 23.0,
    height: float = 12.5,
    show_values: bool = True,
):
    """
    Garantiza que los ejes, nombres de productos (Código | Descripción),
    leyenda y valores numéricos sean 100% legibles en Excel sin ocultarse ni cortarse.
    """
    chart.style = 10
    chart.title = title
    chart.width = width
    chart.height = height
    chart.legend.position = "t"

    # Forzar visibilidad de ambos ejes y posicionar etiquetas fuera de las barras
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.title = x_title
    chart.y_axis.title = y_title
    chart.y_axis.tickLblPos = "low"
    chart.x_axis.tickLblPos = "low"
    chart.y_axis.majorTickMark = "out"
    chart.x_axis.majorTickMark = "out"

    if show_values:
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        chart.dataLabels.showSerName = False
        chart.dataLabels.showCatName = False

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
    df_abc,
    sector_label: str,
) -> bytes:
    wb = Workbook()

    HEX_BLUE = "0057FF"
    HEX_DEEP = "06193E"
    HEX_CYAN = "00B4D8"
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
    # HOJA 1: RESUMEN Y GRÁFICOS (DISEÑO DE REJILLA ESPACIADA — CERO SOLAPAMIENTO)
    # ==========================================================================
    ws_res = wb.active
    ws_res.title = "1. Resumen y Gráficos"
    ws_res.sheet_properties.tabColor = HEX_BLUE
    ws_res.views.sheetView[0].showGridLines = True

    ws_res.merge_cells("A1:R1")
    ws_res["A1"] = "STOCK&YA! — PANEL EJECUTIVO DE CONSOLIDACIÓN Y AUDITORÍA (HIVIMAR S.A.)"
    ws_res["A1"].fill = fill_deep
    ws_res["A1"].font = font_title
    ws_res["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_res.row_dimensions[1].height = 32

    ws_res["A2"] = f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws_res["D2"] = f"Rubro detectado: {sector_label}"
    ws_res["A3"] = "Autoría: Derechos de autor reservados por John Aarón Domínguez Jara 01.593.321"
    ws_res["A2"].font = font_bold
    ws_res["D2"].font = font_bold
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

        # BLOQUE IZQUIERDO SUPERIOR (A5:B15): KPIs
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
        for idx_k, (k_name, k_val) in enumerate(kpis, start=6):
            c1 = ws_res.cell(row=idx_k, column=1, value=k_name)
            c2 = ws_res.cell(row=idx_k, column=2, value=k_val)
            c1.font = font_bold
            c2.font = font_regular
            c1.border = thin_border
            c2.border = thin_border
            if idx_k % 2 == 0:
                c1.fill = fill_ice
                c2.fill = fill_ice

        # BLOQUE IZQUIERDO MEDIO (A17:B20): Tabla Comparativo Volumen
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

        # BLOQUE IZQUIERDO INFERIOR (A22:B27): Estado del Catálogo
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

        # GRÁFICO 1 (CELDA D5 -> Ocupa D5:J23, sin tocar nada más)
        chart_vol = BarChart()
        chart_vol.type = "col"
        data_ref_vol = Reference(ws_res, min_col=2, min_row=17, max_row=20)
        cats_ref_vol = Reference(ws_res, min_col=1, min_row=18, max_row=20)
        chart_vol.add_data(data_ref_vol, titles_from_data=True)
        chart_vol.set_categories(cats_ref_vol)
        configure_legible_bar_chart(
            chart_vol,
            title=f"Comparativo de Unidades: {name_a} vs {name_b} vs Consolidado",
            x_title="Origen / Período",
            y_title="Unidades",
            palette_hex_list=[HEX_BLUE, HEX_CYAN, HEX_RED],
            width=16.5,
            height=10.5,
            show_values=True,
        )
        ws_res.add_chart(chart_vol, "D5")

        # GRÁFICO 2 (CELDA L5 -> Ocupa L5:R23, completamente separado a la derecha)
        chart_pie = DoughnutChart()
        chart_pie.title = "Composición de Auditoría y Consolidación (SKUs)"
        max_st_row = 22 + len(status_counts)
        data_ref_st = Reference(ws_res, min_col=2, min_row=22, max_row=max_st_row)
        cats_ref_st = Reference(ws_res, min_col=1, min_row=23, max_row=max_st_row)
        chart_pie.add_data(data_ref_st, titles_from_data=True)
        chart_pie.set_categories(cats_ref_st)
        chart_pie.legend.position = "b"
        chart_pie.dataLabels = DataLabelList()
        chart_pie.dataLabels.showPercent = True
        chart_pie.dataLabels.showVal = True
        chart_pie.dataLabels.showCatName = True
        chart_pie.width = 16.5
        chart_pie.height = 10.5
        for idx, series in enumerate(chart_pie.series):
            c_hex = [HEX_CYAN, HEX_BLUE, HEX_RED, HEX_DEEP][idx % 4]
            series.graphicalProperties.solidFill = c_hex
        ws_res.add_chart(chart_pie, "L5")

        # FILA 28 EN ADELANTE: TABLA DE LÍNEAS (A28:D40) + GRÁFICO HORIZONTAL (F28:R48)
        if "LINEA_CATEGORIA" in df_reconciled.columns:
            line_summary = (
                df_reconciled.groupby("LINEA_CATEGORIA", as_index=False)[[col_a, col_b, "STOCK_CONSOLIDADO"]]
                .sum()
                .sort_values("STOCK_CONSOLIDADO", ascending=False)
                .head(12)
            )
            start_l_row = 28
            ws_res.cell(row=start_l_row, column=1, value="LÍNEA / CATEGORÍA").fill = fill_deep
            ws_res.cell(row=start_l_row, column=1).font = font_header
            ws_res.cell(row=start_l_row, column=2, value=f"STOCK {name_a}").fill = fill_blue
            ws_res.cell(row=start_l_row, column=2).font = font_header
            ws_res.cell(row=start_l_row, column=3, value=f"STOCK {name_b}").fill = fill_cyan
            ws_res.cell(row=start_l_row, column=3).font = font_header
            ws_res.cell(row=start_l_row, column=4, value="CONSOLIDADO").fill = fill_red
            ws_res.cell(row=start_l_row, column=4).font = font_header

            for idx_l, row_l in enumerate(line_summary.itertuples(index=False), start=start_l_row + 1):
                c_cat = ws_res.cell(row=idx_l, column=1, value=str(row_l[0]))
                c_va = ws_res.cell(row=idx_l, column=2, value=float(row_l[1]))
                c_vb = ws_res.cell(row=idx_l, column=3, value=float(row_l[2]))
                c_vc = ws_res.cell(row=idx_l, column=4, value=float(row_l[3]))
                c_cat.font = font_bold
                for cell_i in (c_cat, c_va, c_vb, c_vc):
                    cell_i.border = thin_border
                    if idx_l % 2 == 0:
                        cell_i.fill = fill_ice

            chart_lines = BarChart()
            chart_lines.type = "bar"
            chart_lines.grouping = "clustered"
            max_l_row = start_l_row + len(line_summary)
            d_ref_l = Reference(ws_res, min_col=2, max_col=3, min_row=start_l_row, max_row=max_l_row)
            c_ref_l = Reference(ws_res, min_col=1, min_row=start_l_row + 1, max_row=max_l_row)
            chart_lines.add_data(d_ref_l, titles_from_data=True)
            chart_lines.set_categories(c_ref_l)
            configure_legible_bar_chart(
                chart_lines,
                title=f"Detalle por Línea / Categoría ({name_a} vs {name_b})",
                x_title="Unidades",
                y_title="Línea / Categoría",
                palette_hex_list=[HEX_BLUE, HEX_CYAN],
                width=24.0,
                height=12.0,
                show_values=True,
            )
            ws_res.add_chart(chart_lines, f"F{start_l_row}")

    ws_res.column_dimensions["A"].width = 42
    ws_res.column_dimensions["B"].width = 18
    ws_res.column_dimensions["C"].width = 18
    ws_res.column_dimensions["D"].width = 18

    # ==========================================================================
    # FUNCIÓN GENÉRICA PARA HOJAS DE DETALLE CON TABLA + RESUMEN DE GRÁFICO LEGIBLE
    # ==========================================================================
    def write_df_sheet_with_legible_chart(
        sheet_title: str,
        tab_hex: str,
        df_data: pd.DataFrame,
        chart_spec: dict = None,
    ):
        ws = wb.create_sheet(title=sheet_title[:31])
        ws.sheet_properties.tabColor = tab_hex
        ws.freeze_panes = "A2"

        # 1. Tabla principal de datos en columnas A..N
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
            ws.column_dimensions[col_letter].width = min(max_len + 4, 42)

        # 2. Cuadro Ejecutivo Top 15 + Gráfico Horizontal debajo del cuadro (cero solapamiento,
        #    con Código + Descripción completa y valores numéricos visibles en cada barra)
        if chart_spec and len(df_data) > 0:
            sort_col = chart_spec["sort_col"]
            ascending = chart_spec.get("ascending", False)
            val_col_names = chart_spec["val_cols"]

            df_top_chart = df_data.copy()
            if chart_spec.get("abs_sort"):
                df_top_chart = df_top_chart.reindex(
                    df_top_chart[sort_col].abs().sort_values(ascending=ascending).index
                ).head(15)
            else:
                df_top_chart = df_top_chart.sort_values(sort_col, ascending=ascending).head(15)

            start_c = len(df_data.columns) + 2  # Columna separada con 1 columna vacía de margen
            ws.cell(row=1, column=start_c, value="PRODUCTO (CÓDIGO | DESCRIPCIÓN)").fill = fill_deep
            ws.cell(row=1, column=start_c).font = font_header
            ws.cell(row=1, column=start_c).alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(start_c)].width = 44

            for v_i, v_name in enumerate(val_col_names, start=1):
                c_hdr = ws.cell(row=1, column=start_c + v_i, value=str(v_name))
                c_hdr.fill = fill_blue if v_i == 1 else (fill_cyan if v_i == 2 else fill_red)
                c_hdr.font = font_header
                c_hdr.alignment = Alignment(horizontal="center", vertical="center")
                ws.column_dimensions[get_column_letter(start_c + v_i)].width = 18

            for r_i, (_, row_item) in enumerate(df_top_chart.iterrows(), start=2):
                cod_str = str(row_item.get("CODIGO", "")).strip()
                desc_str = str(row_item.get("DESCRIPCION", "")).strip()
                full_lbl = f"{cod_str} | {desc_str[:34]}" if cod_str else desc_str[:40]
                c_lbl = ws.cell(row=r_i, column=start_c, value=full_lbl)
                c_lbl.font = font_bold
                c_lbl.border = thin_border
                if r_i % 2 == 0:
                    c_lbl.fill = fill_ice

                for v_i, v_name in enumerate(val_col_names, start=1):
                    val_num = float(row_item.get(v_name, 0.0))
                    c_val = ws.cell(row=r_i, column=start_c + v_i, value=val_num)
                    c_val.font = font_regular
                    c_val.border = thin_border
                    if r_i % 2 == 0:
                        c_val.fill = fill_ice

            max_r_chart = len(df_top_chart) + 1
            ch = BarChart()
            ch.type = "bar"
            ch.grouping = "clustered"
            d_ref = Reference(
                ws,
                min_col=start_c + 1,
                max_col=start_c + len(val_col_names),
                min_row=1,
                max_row=max_r_chart,
            )
            c_ref = Reference(ws, min_col=start_c, min_row=2, max_row=max_r_chart)
            ch.add_data(d_ref, titles_from_data=True)
            ch.set_categories(c_ref)
            configure_legible_bar_chart(
                ch,
                title=chart_spec["title"],
                x_title=chart_spec.get("x_title", "Unidades"),
                y_title="Código | Producto",
                palette_hex_list=chart_spec.get("colors", [HEX_BLUE, HEX_CYAN, HEX_RED]),
                width=24.0,
                height=13.5,
                show_values=True,
            )
            # Anclar el gráfico justo DEBAJO del cuadro resumen (fila 19) para que ambos se lean completos
            ws.add_chart(ch, f"{get_column_letter(start_c)}19")

    if df_reconciled is not None and not df_reconciled.empty:
        write_df_sheet_with_legible_chart(
            "2. Inventario Consolidado",
            HEX_CYAN,
            df_reconciled,
            chart_spec={
                "title": f"Top 15 Productos por Stock Consolidado: {name_a} vs {name_b}",
                "sort_col": "STOCK_CONSOLIDADO",
                "val_cols": [f"STOCK_{name_a}", f"STOCK_{name_b}", "STOCK_CONSOLIDADO"],
                "x_title": "Unidades",
                "colors": [HEX_BLUE, HEX_CYAN, HEX_RED],
            },
        )

        df_diff_only = df_reconciled[df_reconciled["ESTADO_CUADRE"] != "Cuadrado"].copy()
        if not df_diff_only.empty:
            df_diff_only = df_diff_only.reindex(
                df_diff_only["DIFERENCIA"].abs().sort_values(ascending=False).index
            )
            write_df_sheet_with_legible_chart(
                "3. Diferencias y Ajustes",
                HEX_RED,
                df_diff_only,
                chart_spec={
                    "title": f"Top 15 Productos con Mayor Diferencia ({name_b} vs {name_a})",
                    "sort_col": "DIFERENCIA",
                    "abs_sort": True,
                    "val_cols": [f"STOCK_{name_a}", f"STOCK_{name_b}", "DIFERENCIA"],
                    "x_title": "Unidades",
                    "colors": [HEX_BLUE, HEX_CYAN, HEX_RED],
                },
            )

    if df_top_sales is not None and not df_top_sales.empty:
        write_df_sheet_with_legible_chart(
            "4. Top 20 Mayor Rotación",
            HEX_BLUE,
            df_top_sales,
            chart_spec={
                "title": "Top 15 Productos con Mayor Rotación / Salida",
                "sort_col": "UNIDADES_MOVIDAS",
                "val_cols": ["UNIDADES_MOVIDAS"],
                "x_title": "Unidades Movidas",
                "colors": [HEX_BLUE],
            },
        )

    if df_no_move is not None and not df_no_move.empty:
        write_df_sheet_with_legible_chart(
            "5. Top 20 Sin Movimiento",
            HEX_RED,
            df_no_move,
            chart_spec={
                "title": f"Top 15 Productos Inmovilizados (Stock Estancado en {name_b})",
                "sort_col": f"STOCK_{name_b}",
                "val_cols": [f"STOCK_{name_b}"],
                "x_title": "Unidades Estancadas",
                "colors": [HEX_RED],
            },
        )

    if df_abc is not None and not df_abc.empty:
        write_df_sheet_with_legible_chart(
            "6. Matriz Gerencial ABC",
            HEX_DEEP,
            df_abc,
            chart_spec={
                "title": "Top 15 Productos Clase A por Valor Consolidado ($ USD)",
                "sort_col": "VALOR_CONSOLIDADO_USD",
                "val_cols": ["VALOR_CONSOLIDADO_USD"],
                "x_title": "Valor en USD ($)",
                "colors": [HEX_BLUE],
            },
        )

    if df_ai is not None and not df_ai.empty:
        write_df_sheet_with_legible_chart(
            "7. Proyección Demanda IA",
            HEX_CYAN,
            df_ai,
            chart_spec={
                "title": "Top 15 Productos: Stock Actual vs Expectativa Venta Próximo Mes (IA)",
                "sort_col": "EXPECTATIVA_VENTA_IA",
                "val_cols": ["STOCK_ACTUAL", "EXPECTATIVA_VENTA_IA", "SUGERIDO_COMPRA_IA"],
                "x_title": "Unidades",
                "colors": [HEX_BLUE, HEX_RED, HEX_CYAN],
            },
        )

    if df_reconciled is not None and not df_reconciled.empty:
        ws_bi = wb.create_sheet(title="8. Base Power BI (Datos)")
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
                "#FF3B5C",
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

    @st.dialog("Bienvenido a STOCK&YA! — Guía Ejecutiva HIVIMAR S.A.", width="large")
    def show_interactive_guide():
        st.markdown(
            """
            ### Plataforma de Consolidación e Inteligencia Gerencial — **HIVIMAR S.A.**
            Diseñada para unificar reportes de **SAP, HIVIPRO, Oracle, Odoo, Excel, CSV o Google Sheets** sin errores de compatibilidad y con visualización gerencial inmediata.

            #### Módulos integrados:
            1. **Consolidación Inteligente:** Carga tus dos fuentes (meses distintos, líneas distintas o Lógico vs Físico) y obtén el cruce, suma consolidada y detección de variaciones al instante.
            2. **Matriz Gerencial ABC & Salud ($ USD):** Clasificación Pareto 80/15/5, valorización de capital inmovilizado e **Índice de Salud de Inventario (0-100 pts)**.
            3. **Top 20 Rotación y Sin Movimiento:** Identifica los productos de mayor salida y los artículos estancados.
            4. **Creciente y Decreciente (% Real):** Termómetro de variaciones porcentuales por línea y por SKU.
            5. **Consolidado 12 Meses + IA:** Proyección de demanda del próximo mes y sugerido de compra automático.
            6. **Exportación Excel + Power BI (`&YA!`):** Gráficos legibles con código, descripción y cantidades visibles, sin solapamientos y con tabla lista para **Power BI**.
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
                    STOCK<span style="color:#00E5FF;">&amp;</span><span style="color:#FF3B5C;">YA!</span>
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
                "Consolidación Inteligente",
                "Matriz Gerencial ABC & Salud",
                "Top 20 Ventas y Sin Movimiento",
                "Creciente y Decreciente (% Real)",
                "Consolidado 12 Meses + IA",
            ],
            label_visibility="collapsed",
        )

        st.markdown("---")
        if st.button("Guía Ejecutiva de Uso", use_container_width=True):
            show_interactive_guide()

        if st.session_state.get("df_reconciled") is not None:
            if st.button("Limpiar / Cargar Nuevos Archivos", use_container_width=True):
                for k in ["df_reconciled", "name_a", "name_b", "last_animated_sig", "df_abc"]:
                    st.session_state.pop(k, None)
                st.rerun()

        st.markdown(
            """
            <div style="margin-top: 1.3rem; font-size: 0.75rem; color: #CBD5E1; line-height: 1.45;">
              Compatible con Excel (.xlsx/.xls), CSV, Google Sheets y Power BI.<br>
              Motor gerencial multi-línea activado.
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown(
        """
        <div class="hero-banner">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem;">
            <div>
              <span style="background:rgba(6,17,33,0.45); border:1px solid #00E5FF; color:#FFFFFF; padding:4px 12px; border-radius:999px; font-size:0.75rem; font-weight:800; letter-spacing:0.06em;">
                PLATAFORMA EJECUTIVA DE CONSOLIDACIÓN &bull; HIVIMAR S.A.
              </span>
              <h1 style="margin:0.45rem 0 0.25rem 0; font-size:1.95rem; font-weight:800; color:#FFFFFF;">
                Bienvenido a STOCK<span style="color:#00E5FF;">&amp;</span><span style="color:#FF3B5C;">YA!</span>
              </h1>
              <p style="margin:0; color:#F0F9FF; font-size:0.93rem;">
                Consolida sistemas (SAP, HIVIPRO, Oracle, Odoo, Excel, Google Sheets), audita diferencias y analiza valorización ABC para toma de decisiones gerenciales.
              </p>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ==========================================================================
    # MÓDULO UNIFICADO: CONSOLIDACIÓN INTELIGENTE (CARGA + RESULTADOS INMEDIATOS)
    # ==========================================================================
    if nav_option == "Consolidación Inteligente":
        render_section_header(
            "upload",
            "Consolidación Inteligente — Carga y Cruce de Fuentes",
            "Sube tus dos archivos (o selecciona dos pestañas de un mismo libro Excel) para consolidar y auditar al instante.",
        )

        c_mode1, c_mode2 = st.columns([2, 1.3])
        with c_mode1:
            op_mode_label = st.radio(
                "Objetivo del análisis:",
                [
                    "Auditoría y Cuadre Exacto (Ej. SAP vs HIVIPRO / Lógico vs Físico)",
                    "Consolidación Multi-Línea / Comparativo de Meses (Suma + Variación)",
                ],
                horizontal=True,
                key="op_mode_radio",
            )
            operation_mode = "auditoria" if "Auditoría" in op_mode_label else "consolidacion"
        with c_mode2:
            same_book_mode = st.checkbox(
                "Cruzar dos pestañas dentro de un mismo archivo Excel",
                value=False,
                key="same_book_chk",
            )

        col_src1, col_src2 = st.columns(2)
        sheets_1 = {}
        sheets_2 = {}
        fname_1 = ""
        fname_2 = ""
        sel_sheet_1 = ""
        sel_sheet_2 = ""

        with col_src1:
            render_source_badge_header("01", "Primera Fuente (Ej. SAP / Mes Anterior / Línea 1)", is_second=False)
            mode_1 = st.radio(
                "Origen Fuente 1:",
                ["Subir Archivo (Excel / CSV)", "Enlace de Google Sheets"],
                horizontal=True,
                key="mode_1",
            )
            if mode_1 == "Subir Archivo (Excel / CSV)":
                up1 = st.file_uploader(
                    "Selecciona el archivo de la Fuente 1",
                    type=["xlsx", "xls", "csv"],
                    key="up1",
                )
                if up1:
                    try:
                        fname_1 = up1.name
                        sheets_1 = read_uploaded_workbook(up1)
                    except Exception as exc:
                        st.error(f"Error al leer el archivo 1: {exc}")
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
                if len(options_1) > 1 and not same_book_mode:
                    options_1 = ["CONSOLIDAR TODAS LAS PESTAÑAS"] + options_1
                sel_sheet_1 = st.selectbox(
                    "Pestaña a analizar (Fuente 1):",
                    options_1,
                    index=0,
                    key="sel_s1",
                )
                smart_lbl_1 = derive_smart_source_label(fname_1, sel_sheet_1, "SAP")
                name_a = sanitize_label(
                    st.text_input("Etiqueta Fuente 1:", value=smart_lbl_1, key=f"lbl_a_{fname_1}_{sel_sheet_1}"),
                    "FUENTE_1",
                )

        with col_src2:
            render_source_badge_header("02", "Segunda Fuente (Ej. HIVIPRO / Mes Actual / Conteo Físico)", is_second=True)
            if same_book_mode:
                if sheets_1:
                    fname_2 = fname_1
                    sheets_2 = sheets_1
                    st.info(f"Usando pestañas del libro cargado: **{fname_1}**")
                else:
                    st.warning("Sube primero el libro Excel en la columna izquierda (01).")
            else:
                mode_2 = st.radio(
                    "Origen Fuente 2:",
                    ["Subir Archivo (Excel / CSV)", "Enlace de Google Sheets"],
                    horizontal=True,
                    key="mode_2",
                )
                if mode_2 == "Subir Archivo (Excel / CSV)":
                    up2 = st.file_uploader(
                        "Selecciona el archivo de la Fuente 2",
                        type=["xlsx", "xls", "csv"],
                        key="up2",
                    )
                    if up2:
                        try:
                            fname_2 = up2.name
                            sheets_2 = read_uploaded_workbook(up2)
                        except Exception as exc:
                            st.error(f"Error al leer el archivo 2: {exc}")
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
                if len(options_2) > 1 and not same_book_mode:
                    options_2 = ["CONSOLIDAR TODAS LAS PESTAÑAS"] + options_2
                default_idx_2 = 1 if (same_book_mode and len(options_2) > 1) else 0
                sel_sheet_2 = st.selectbox(
                    "Pestaña a analizar (Fuente 2):",
                    options_2,
                    index=default_idx_2,
                    key="sel_s2",
                )
                smart_lbl_2 = derive_smart_source_label(fname_2, sel_sheet_2, "HIVIPRO")
                if smart_lbl_2 == name_a:
                    smart_lbl_2 = f"{smart_lbl_2}_2"[:24]
                name_b = sanitize_label(
                    st.text_input("Etiqueta Fuente 2:", value=smart_lbl_2, key=f"lbl_b_{fname_2}_{sel_sheet_2}"),
                    "FUENTE_2",
                )

        if sheets_1 and not sheets_2 and not same_book_mode:
            st.warning(
                "Fuente 1 cargada correctamente. Sube ahora el **segundo archivo** en el bloque **02** "
                "(o marca la casilla superior si ambas pestañas están en el mismo libro Excel)."
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
                time.sleep(1.1)
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

    # ==========================================================================
    # RECUPERACIÓN GLOBAL DE ESTADO PARA TODAS LAS VISTAS
    # ==========================================================================
    df_reconciled = st.session_state.get("df_reconciled", None)
    name_a = st.session_state.get("name_a", "FUENTE_1")
    name_b = st.session_state.get("name_b", "FUENTE_2")
    sector_label = st.session_state.get("sector_label", "Automotriz e Industrial (HIVIMAR)")

    if df_reconciled is None and nav_option not in ["Consolidación Inteligente", "Consolidado 12 Meses + IA"]:
        st.warning(
            "Aún no has cargado los archivos de inventario. Ve a **Consolidación Inteligente** "
            "en el menú izquierdo para subir tus fuentes."
        )
        return

    df_top_sales = None
    df_no_move = None
    df_abc = None
    abc_metrics = None

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

        ref_cost = st.session_state.get("ref_unit_cost", 15.0)
        df_abc, abc_metrics = compute_managerial_abc_analysis(
            df_reconciled, name_a, name_b, fallback_unit_cost=ref_cost
        )
        st.session_state.df_abc = df_abc

        total_skus = len(df_reconciled)
        cuadrados_cnt = int((df_reconciled["ESTADO_CUADRE"] == "Cuadrado").sum())
        descuadrados_cnt = total_skus - cuadrados_cnt
        exactitud_pct = (cuadrados_cnt / total_skus * 100.0) if total_skus > 0 else 0.0

        sum_a = float(df_reconciled[col_a].sum())
        sum_b = float(df_reconciled[col_b].sum())
        sum_cons = float(df_reconciled["STOCK_CONSOLIDADO"].sum())
        var_inv_pct = ((sum_b - sum_a) / sum_a * 100.0) if sum_a > 0 else 0.0

        # Barra superior de exportación global
        st.markdown("---")
        c_info, c_exp1, c_exp2 = st.columns([2.3, 1.4, 1.1])
        with c_info:
            st.markdown(
                f"**Consolidado Activo:** `{name_a}` vs `{name_b}` &nbsp;|&nbsp; "
                f"**Rubro:** {sector_label} &nbsp;|&nbsp; **Materiales Únicos:** {total_skus:,}"
            )
        with c_exp1:
            excel_bytes = build_master_excel_export(
                df_reconciled=df_reconciled,
                name_a=name_a,
                name_b=name_b,
                df_top_sales=df_top_sales,
                df_no_move=df_no_move,
                df_ai=st.session_state.get("df_ai", None),
                df_abc=df_abc,
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

        # ----------------------------------------------------------------------
        # VISTA 1: RESULTADOS DE CONSOLIDACIÓN INTELIGENTE
        # ----------------------------------------------------------------------
        if nav_option == "Consolidación Inteligente":
            render_section_header(
                "audit",
                "Resultados de Consolidación y Auditoría Exacta",
                "Balance consolidado, porcentaje de exactitud y detalle por material.",
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

            col_g1, col_g2 = st.columns(2)
            with col_g1:
                fig_overview = go.Figure()
                fig_overview.add_trace(
                    go.Bar(
                        x=[name_a, name_b, "CONSOLIDADO TOTAL"],
                        y=[sum_a, sum_b, sum_cons],
                        text=[f"{sum_a:,.0f} uds", f"{sum_b:,.0f} uds", f"{sum_cons:,.0f} uds"],
                        textposition="outside",
                        textfont=dict(color="#FFFFFF"),
                        marker_color=["#0057FF", "#00E5FF", "#FF3B5C"],
                    )
                )
                apply_dark_plotly_layout(
                    fig_overview,
                    title=f"Balance General: {name_a} vs {name_b} vs Consolidado",
                    height=370,
                    yaxis_title="Unidades",
                    showlegend=False,
                )
                st.plotly_chart(fig_overview, use_container_width=True)

            with col_g2:
                top10_prod = df_reconciled.sort_values("STOCK_CONSOLIDADO", ascending=False).head(10).copy()
                top10_prod["LABEL"] = top10_prod["CODIGO"].astype(str) + " | " + top10_prod["DESCRIPCION"].str.slice(0, 24)
                fig_top10 = go.Figure()
                fig_top10.add_trace(
                    go.Bar(
                        y=top10_prod["LABEL"][::-1],
                        x=top10_prod[col_a][::-1],
                        name=name_a,
                        orientation="h",
                        marker_color="#0057FF",
                    )
                )
                fig_top10.add_trace(
                    go.Bar(
                        y=top10_prod["LABEL"][::-1],
                        x=top10_prod[col_b][::-1],
                        name=name_b,
                        orientation="h",
                        marker_color="#00E5FF",
                    )
                )
                fig_top10.update_layout(barmode="group")
                apply_dark_plotly_layout(
                    fig_top10,
                    title=f"Top 10 Productos por Volumen ({name_a} vs {name_b})",
                    height=370,
                    xaxis_title="Unidades",
                )
                st.plotly_chart(fig_top10, use_container_width=True)

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
                busqueda_sku = st.text_input("Buscar por Código, Descripción o Línea:", placeholder="Ej. 1001747, SKF, Sellado...")

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

        # ----------------------------------------------------------------------
        # VISTA 2: MATRIZ GERENCIAL ABC & SALUD DE INVENTARIO (NUEVO MÓDULO EJECUTIVO)
        # ----------------------------------------------------------------------
        if nav_option == "Matriz Gerencial ABC & Salud":
            render_section_header(
                "executive",
                "Matriz Gerencial ABC (Pareto 80/15/5) e Índice de Salud de Inventario",
                "Priorización financiera y operativa para toma de decisiones de Gerencia General y Supply Chain.",
            )

            if not abc_metrics["has_real_price"]:
                c_p1, c_p2 = st.columns([1.4, 2.6])
                with c_p1:
                    new_cost = st.number_input(
                        "Costo Unitario Promedio Referencial ($ USD):",
                        min_value=0.1,
                        max_value=50000.0,
                        value=float(st.session_state.get("ref_unit_cost", 15.0)),
                        step=1.0,
                        key="ref_cost_input",
                    )
                    if new_cost != st.session_state.get("ref_unit_cost", 15.0):
                        st.session_state.ref_unit_cost = new_cost
                        st.rerun()
                with c_p2:
                    st.info(
                        "Tus archivos actuales no incluyen columna de precio/costo unitario. "
                        "Ajusta el **Costo Promedio Referencial ($ USD)** a la izquierda para simular la valorización financiera del inventario."
                    )

            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.markdown(
                    f"""
                    <div class="kpi-card cyan-top">
                      <div class="kpi-label">Índice de Salud Gerencial</div>
                      <div class="kpi-value">{abc_metrics['health_score']:.1f} / 100</div>
                      <div class="kpi-sub badge-cyan">Exactitud + Control Clase A + Rotación</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with m2:
                st.markdown(
                    f"""
                    <div class="kpi-card">
                      <div class="kpi-label">Valorización Consolidada</div>
                      <div class="kpi-value">${abc_metrics['total_val_usd']:,.2f}</div>
                      <div class="kpi-sub badge-ok">En {name_b}: ${abc_metrics['val_b_usd']:,.2f} USD</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with m3:
                st.markdown(
                    f"""
                    <div class="kpi-card">
                      <div class="kpi-label">SKUs Clase A (80% del Valor)</div>
                      <div class="kpi-value">{abc_metrics['class_a_skus']:,} SKUs</div>
                      <div class="kpi-sub badge-ok">Representan ${abc_metrics['class_a_usd']:,.2f} USD</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with m4:
                st.markdown(
                    f"""
                    <div class="kpi-card red-top">
                      <div class="kpi-label">Capital Clase A Sin Rotación</div>
                      <div class="kpi-value">${abc_metrics['immobilized_a_usd']:,.2f}</div>
                      <div class="kpi-sub badge-err">Impacto diferencias: ${abc_metrics['diff_val_usd']:,.2f}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            st.markdown("<br>", unsafe_allow_html=True)

            # Recomendaciones Ejecutivas Automáticas para Gerencia
            st.markdown(
                f"""
                <div class="exec-advice-card">
                  <strong style="color:#00E5FF;">DIAGNÓSTICO EJECUTIVO DE CONTROL:</strong>
                  De los <strong>{total_skus:,}</strong> materiales analizados, solo <strong>{abc_metrics['class_a_skus']:,} SKUs (Clase A)</strong>
                  concentran el <strong>80% del capital invertido (${abc_metrics['class_a_usd']:,.2f} USD)</strong>.
                  Enfocar los conteos cíclicos semanales en estos {abc_metrics['class_a_skus']:,} códigos protege el 80% del patrimonio de inventario de HIVIMAR S.A.
                </div>
                """,
                unsafe_allow_html=True,
            )

            col_abc1, col_abc2 = st.columns(2)
            with col_abc1:
                abc_summary = (
                    df_abc.groupby("CLASE_ABC", as_index=False)
                    .agg(SKUs=("CODIGO", "count"), Valor_USD=("VALOR_CONSOLIDADO_USD", "sum"))
                )
                fig_abc_pie = px.pie(
                    abc_summary,
                    names="CLASE_ABC",
                    values="Valor_USD",
                    hole=0.48,
                    color="CLASE_ABC",
                    color_discrete_map={
                        "A (Crítico 80%)": "#0057FF",
                        "B (Intermedio 15%)": "#00E5FF",
                        "C (Cola 5%)": "#FF3B5C",
                    },
                )
                fig_abc_pie.update_traces(textinfo="percent+label", textfont=dict(color="#FFFFFF"))
                apply_dark_plotly_layout(
                    fig_abc_pie,
                    title="Distribución de Capital por Clasificación ABC ($ USD)",
                    height=380,
                )
                st.plotly_chart(fig_abc_pie, use_container_width=True)

            with col_abc2:
                action_summary = (
                    df_abc.groupby("ACCION_GERENCIAL", as_index=False)
                    .agg(Cantidad_SKUs=("CODIGO", "count"), Valor_USD=("VALOR_CONSOLIDADO_USD", "sum"))
                    .sort_values("Valor_USD", ascending=True)
                )
                fig_act = px.bar(
                    action_summary,
                    x="Valor_USD",
                    y="ACCION_GERENCIAL",
                    orientation="h",
                    text="Cantidad_SKUs",
                    color="ACCION_GERENCIAL",
                    color_discrete_sequence=["#00E5FF", "#0057FF", "#FF3B5C", "#38BDF8"],
                )
                fig_act.update_traces(
                    texttemplate="$%{x:,.0f} (%{text} SKUs)",
                    textposition="outside",
                    textfont=dict(color="#FFFFFF"),
                )
                apply_dark_plotly_layout(
                    fig_act,
                    title="Plan de Acción Gerencial por Monto Involucrado ($ USD)",
                    height=380,
                    xaxis_title="Monto ($ USD)",
                    yaxis_title="",
                    showlegend=False,
                )
                st.plotly_chart(fig_act, use_container_width=True)

            st.dataframe(
                df_abc[
                    [
                        "CODIGO",
                        "DESCRIPCION",
                        "LINEA_CATEGORIA",
                        "CLASE_ABC",
                        "STOCK_CONSOLIDADO",
                        "COSTO_UNIT_USD",
                        "VALOR_CONSOLIDADO_USD",
                        "PARTICIPACION_VALOR_%",
                        "ACUMULADO_PARETO_%",
                        "ACCION_GERENCIAL",
                    ]
                ],
                use_container_width=True,
                height=400,
            )

        # ----------------------------------------------------------------------
        # VISTA 3: TOP 20 VENTAS Y SIN MOVIMIENTO
        # ----------------------------------------------------------------------
        if nav_option == "Top 20 Ventas y Sin Movimiento":
            render_section_header(
                "ranking",
                "Ranking de Rotación y Capital Inmovilizado",
                "Participación porcentual real sobre unidades movidas e inventario estancado.",
            )
            tab_top, tab_dead = st.tabs(["Top 20 Más Vendidos / Mayor Salida", "Top 20 Sin Movimiento (Inmovilizado)"])

            with tab_top:
                if df_top_sales is not None and not df_top_sales.empty:
                    df_top_sales["ETIQUETA"] = df_top_sales["CODIGO"].astype(str) + " | " + df_top_sales["DESCRIPCION"].str.slice(0, 30)
                    fig_top = px.bar(
                        df_top_sales.sort_values("UNIDADES_MOVIDAS", ascending=True),
                        x="UNIDADES_MOVIDAS",
                        y="ETIQUETA",
                        orientation="h",
                        text="PARTICIPACION_%",
                        color="UNIDADES_MOVIDAS",
                        color_continuous_scale=["#0057FF", "#00C8FF", "#00E5FF"],
                    )
                    fig_top.update_traces(
                        texttemplate="%{x:,.0f} uds (%{text}%)",
                        textposition="outside",
                        textfont=dict(color="#FFFFFF"),
                    )
                    apply_dark_plotly_layout(
                        fig_top,
                        title=f"Top 20 Productos — {mov_label}",
                        height=620,
                        xaxis_title="Unidades",
                        yaxis_title="",
                        showlegend=False,
                    )
                    fig_top.update_layout(coloraxis_showscale=False)
                    st.plotly_chart(fig_top, use_container_width=True)
                    st.dataframe(
                        df_top_sales[["CODIGO", "DESCRIPCION", "LINEA_CATEGORIA", col_a, col_b, "UNIDADES_MOVIDAS", "PARTICIPACION_%"]],
                        use_container_width=True,
                    )
                else:
                    st.info(
                        "No se detectaron salidas/egresos entre las fuentes seleccionadas (ambas fuentes tienen stock idéntico o sin columna de ventas). "
                        "A continuación se muestra el **Top 20 de Mayor Volumen Consolidado**:"
                    )
                    top20_vol = df_reconciled.sort_values("STOCK_CONSOLIDADO", ascending=False).head(20).copy()
                    top20_vol["ETIQUETA"] = top20_vol["CODIGO"].astype(str) + " | " + top20_vol["DESCRIPCION"].str.slice(0, 30)
                    fig_v20 = px.bar(
                        top20_vol.sort_values("STOCK_CONSOLIDADO", ascending=True),
                        x="STOCK_CONSOLIDADO",
                        y="ETIQUETA",
                        orientation="h",
                        text="STOCK_CONSOLIDADO",
                        color_discrete_sequence=["#00E5FF"],
                    )
                    fig_v20.update_traces(texttemplate="%{x:,.0f} uds", textposition="outside", textfont=dict(color="#FFFFFF"))
                    apply_dark_plotly_layout(
                        fig_v20,
                        title="Top 20 Productos por Volumen Consolidado",
                        height=600,
                        xaxis_title="Unidades Consolidadas",
                        yaxis_title="",
                        showlegend=False,
                    )
                    st.plotly_chart(fig_v20, use_container_width=True)

            with tab_dead:
                if df_no_move is not None and not df_no_move.empty:
                    total_dead_units = float(df_mov[(df_mov["UNIDADES_MOVIDAS"] == 0) & (df_mov[col_b] > 0)][col_b].sum())
                    dead_pct = (total_dead_units / sum_b * 100.0) if sum_b > 0 else 0.0
                    st.warning(
                        f"El **{dead_pct:.2f}%** del inventario actual ({total_dead_units:,.0f} unidades) "
                        f"no registró egresos/variación en el período analizado."
                    )
                    df_no_move["ETIQUETA"] = df_no_move["CODIGO"].astype(str) + " | " + df_no_move["DESCRIPCION"].str.slice(0, 30)
                    fig_dead = px.bar(
                        df_no_move.sort_values(col_b, ascending=True),
                        x=col_b,
                        y="ETIQUETA",
                        orientation="h",
                        text="PESO_INVENTARIO_%",
                        color=col_b,
                        color_continuous_scale=["#0057FF", "#FF3B5C"],
                    )
                    fig_dead.update_traces(
                        texttemplate="%{x:,.0f} uds (%{text}%)",
                        textposition="outside",
                        textfont=dict(color="#FFFFFF"),
                    )
                    apply_dark_plotly_layout(
                        fig_dead,
                        title="Top 20 Productos Sin Movimiento (Mayor Stock Estancado)",
                        height=620,
                        xaxis_title="Stock Estancado (Unidades)",
                        yaxis_title="",
                        showlegend=False,
                    )
                    fig_dead.update_layout(coloraxis_showscale=False)
                    st.plotly_chart(fig_dead, use_container_width=True)
                    st.dataframe(
                        df_no_move[["CODIGO", "DESCRIPCION", "LINEA_CATEGORIA", col_b, "PESO_INVENTARIO_%", "ESTADO_CUADRE"]],
                        use_container_width=True,
                    )
                else:
                    st.success("Todos los productos con stock registraron movimiento en el período.")

        # ----------------------------------------------------------------------
        # VISTA 4: CRECIENTE Y DECRECIENTE (% REAL)
        # ----------------------------------------------------------------------
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
                    color_discrete_map={"Creciente (+)": "#00E5FF", "Decreciente (-)": "#FF3B5C"},
                    text="VARIACION_%",
                )
                fig_cat.update_traces(texttemplate="%{text:+.2f}%", textposition="outside", textfont=dict(color="#FFFFFF"))
                apply_dark_plotly_layout(
                    fig_cat,
                    title="Variación Porcentual de Inventario por Línea / Categoría (%)",
                    height=420,
                    xaxis_title="Variación Porcentual (%)",
                    yaxis_title="",
                )
                st.plotly_chart(fig_cat, use_container_width=True)

            df_pct_sku = df_reconciled[df_reconciled[col_a] > 0].copy()
            if not df_pct_sku.empty:
                df_pct_sku["VAR_%"] = ((df_pct_sku[col_b] - df_pct_sku[col_a]) / df_pct_sku[col_a] * 100.0).round(2)
                df_pct_sku["ETIQUETA"] = df_pct_sku["CODIGO"].astype(str) + " | " + df_pct_sku["DESCRIPCION"].str.slice(0, 28)
                top_up = df_pct_sku.sort_values("VAR_%", ascending=False).head(10)
                top_down = df_pct_sku.sort_values("VAR_%", ascending=True).head(10)
                df_movers = pd.concat([top_up, top_down]).drop_duplicates("CODIGO")
                df_movers["DIRECCION"] = np.where(df_movers["VAR_%"] >= 0, "Creciente (+)", "Decreciente (-)")

                fig_movers = px.bar(
                    df_movers.sort_values("VAR_%"),
                    x="VAR_%",
                    y="ETIQUETA",
                    orientation="h",
                    color="DIRECCION",
                    color_discrete_map={"Creciente (+)": "#00E5FF", "Decreciente (-)": "#FF3B5C"},
                    text="VAR_%",
                )
                fig_movers.update_traces(texttemplate="%{text:+.1f}%", textposition="outside", textfont=dict(color="#FFFFFF"))
                apply_dark_plotly_layout(
                    fig_movers,
                    title="Top Productos con Mayor Variación Porcentual (Creciente vs Decreciente)",
                    height=560,
                    xaxis_title="Variación %",
                    yaxis_title="",
                )
                st.plotly_chart(fig_movers, use_container_width=True)

    # ==========================================================================
    # VISTA 5: CONSOLIDADO 12 MESES + IA PREDICTIVA
    # ==========================================================================
    if nav_option == "Consolidado 12 Meses + IA":
        render_section_header(
            "ai",
            "Consolidado Histórico (Hasta 12 Meses) y Proyección de Demanda con IA",
            "Ordenamiento cronológico automático, modelado predictivo vectorizado (Holt + Regresión) y sugerido de compra.",
        )
        st.markdown(
            "Sube hasta **12 archivos mensuales** (Excel o CSV). El motor los ordenará cronológicamente, "
            "unificará códigos entre meses y calculará la **expectativa de venta y sugerido de compra del próximo mes**."
        )

        uploaded_12 = st.file_uploader(
            "Selecciona los archivos mensuales (hasta 12 archivos):",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key="up_12m",
        )

        if uploaded_12:
            sig12 = "_".join(sorted(f.name for f in uploaded_12))
            sorted_files = sorted(
                enumerate(uploaded_12[:12]),
                key=lambda pair: extract_period_sort_key(pair[1].name, pair[0]),
            )

            monthly_records = []
            sample_dfs = []
            file_names_12 = []
            skipped_files = []

            for idx, f_obj in sorted_files:
                _, _, period_lbl = extract_period_sort_key(f_obj.name, idx)
                try:
                    s_dict = read_uploaded_workbook(f_obj)
                    if not s_dict:
                        skipped_files.append(f_obj.name)
                        continue
                    raw_m = merge_all_workbook_sheets(s_dict) if len(s_dict) > 1 else next(iter(s_dict.values()))
                    if raw_m is None or raw_m.empty:
                        skipped_files.append(f_obj.name)
                        continue
                    file_names_12.append(f_obj.name)
                    sample_dfs.append(raw_m)
                    sch_m = auto_detect_schema(raw_m)
                    cons_m = consolidate_source(raw_m, sch_m, "M")
                    monthly_records.append({"label": period_lbl, "df": cons_m})
                except Exception as exc:
                    skipped_files.append(f"{f_obj.name} ({exc})")

            if skipped_files:
                st.warning(f"Archivos omitidos por estar vacíos o tener formato no legible: {', '.join(skipped_files)}")

            if not monthly_records:
                st.error("No se pudieron extraer datos válidos de los archivos mensuales cargados.")
            else:
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
                    time.sleep(1.1)
                    ph12.empty()
                    st.session_state.last_12m_anim = sig12

                df_ai, piv_sales = run_ai_demand_forecast(monthly_records)
                st.session_state.df_ai = df_ai

                if df_ai.empty:
                    st.warning("No se encontraron registros consolidables en los archivos cargados.")
                else:
                    tot_forecast = float(df_ai["EXPECTATIVA_VENTA_IA"].sum())
                    tot_suggested_buy = float(df_ai["SUGERIDO_COMPRA_IA"].sum())
                    urgent_cnt = int(df_ai["DIAGNOSTICO_IA"].str.contains("Reabastecer").sum())
                    accel_cnt = int(df_ai["DIAGNOSTICO_IA"].str.contains("Aceleración").sum())
                    avg_conf = float(df_ai["CONFIANZA_IA_%"].mean())

                    a1, a2, a3, a4 = st.columns(4)
                    with a1:
                        st.markdown(
                            f"""
                            <div class="kpi-card">
                              <div class="kpi-label">Expectativa Venta Próximo Mes</div>
                              <div class="kpi-value">{tot_forecast:,.0f} uds</div>
                              <div class="kpi-sub badge-ok">Modelo IA ({len(monthly_records)} períodos)</div>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
                    with a2:
                        st.markdown(
                            f"""
                            <div class="kpi-card red-top">
                              <div class="kpi-label">Alertas Quiebre de Stock</div>
                              <div class="kpi-value">{urgent_cnt:,} SKUs</div>
                              <div class="kpi-sub badge-err">Stock menor a demanda IA</div>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
                    with a3:
                        st.markdown(
                            f"""
                            <div class="kpi-card cyan-top">
                              <div class="kpi-label">Sugerido de Compra IA</div>
                              <div class="kpi-value">{tot_suggested_buy:,.0f} uds</div>
                              <div class="kpi-sub badge-cyan">{accel_cnt:,} SKUs en aceleración (+10%)</div>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
                    with a4:
                        st.markdown(
                            f"""
                            <div class="kpi-card">
                              <div class="kpi-label">Precisión / Confianza Promedio IA</div>
                              <div class="kpi-value">{avg_conf:.1f}%</div>
                              <div class="kpi-sub badge-ok">Holt + Regresión Ponderada</div>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )

                    st.markdown("<br>", unsafe_allow_html=True)

                    monthly_totals = piv_sales.sum(axis=0)
                    hist_x = list(monthly_totals.index)
                    hist_y = list(monthly_totals.values)

                    if len(hist_x) > 0:
                        fig_ai = go.Figure()
                        fig_ai.add_trace(
                            go.Scatter(
                                x=hist_x,
                                y=hist_y,
                                mode="lines+markers+text",
                                name="Histórico Real",
                                text=[f"{v:,.0f}" for v in hist_y],
                                textposition="top center",
                                textfont=dict(color="#FFFFFF"),
                                line=dict(color="#00E5FF", width=3.5),
                                marker=dict(size=9, color="#0057FF", line=dict(width=2, color="#00E5FF")),
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
                                textfont=dict(color="#FF5C77"),
                                line=dict(color="#FF3B5C", width=3.5, dash="dash"),
                                marker=dict(size=12, color="#FF3B5C", symbol="diamond"),
                            )
                        )
                        apply_dark_plotly_layout(
                            fig_ai,
                            title="Curva Histórica Consolidada y Expectativa de Venta Próximo Mes (IA)",
                            height=400,
                            yaxis_title="Unidades",
                        )
                        st.plotly_chart(fig_ai, use_container_width=True)

                    render_section_header(
                        "table",
                        "Tabla de Expectativa de Venta y Recomendación de Compra por Producto",
                        "Filtra por diagnóstico IA para priorizar órdenes de reposición.",
                    )

                    col_f_ai1, col_f_ai2 = st.columns([1.4, 2])
                    with col_f_ai1:
                        filtro_ai = st.selectbox(
                            "Filtrar diagnóstico IA:",
                            [
                                "Todos los productos",
                                "Solo [CRÍTICO] Reabastecer Urgente",
                                "Solo [ALZA] Demanda en Aceleración",
                                "Solo [INMOVILIZADO] Sin Rotación",
                                "Solo [BAJA] Demanda en Desaceleración",
                            ],
                        )
                    with col_f_ai2:
                        busqueda_ai = st.text_input("Buscar SKU o descripción en proyección IA:", placeholder="Ej. 1001747, Rodamiento...")

                    df_ai_view = df_ai.copy()
                    if filtro_ai == "Solo [CRÍTICO] Reabastecer Urgente":
                        df_ai_view = df_ai_view[df_ai_view["DIAGNOSTICO_IA"].str.contains("CRÍTICO")]
                    elif filtro_ai == "Solo [ALZA] Demanda en Aceleración":
                        df_ai_view = df_ai_view[df_ai_view["DIAGNOSTICO_IA"].str.contains("ALZA")]
                    elif filtro_ai == "Solo [INMOVILIZADO] Sin Rotación":
                        df_ai_view = df_ai_view[df_ai_view["DIAGNOSTICO_IA"].str.contains("INMOVILIZADO")]
                    elif filtro_ai == "Solo [BAJA] Demanda en Desaceleración":
                        df_ai_view = df_ai_view[df_ai_view["DIAGNOSTICO_IA"].str.contains("BAJA")]

                    if busqueda_ai.strip():
                        q_ai = busqueda_ai.strip().lower()
                        df_ai_view = df_ai_view[
                            df_ai_view["CODIGO"].str.lower().str.contains(q_ai, na=False)
                            | df_ai_view["DESCRIPCION"].str.lower().str.contains(q_ai, na=False)
                        ]

                    st.dataframe(df_ai_view, use_container_width=True, height=420)

    st.markdown(
        """
        <div class="copyright-footer">
          STOCK&amp;YA! &bull; Plataforma de Inteligencia Gerencial de Inventarios para HIVIMAR S.A. &nbsp;|&nbsp;
          Derechos de autor reservados por John Aarón Domínguez Jara 01.593.321
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
