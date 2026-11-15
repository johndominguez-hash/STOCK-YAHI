import io
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
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# ICONO VECTORIAL DE PESTAÑA DEL NAVEGADOR (DATA URI SVG)
# ==============================================================================
PAGE_FAVICON_SVG = (
    "data:image/svg+xml;utf8,"
    "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'>"
    "<circle cx='32' cy='32' r='28' fill='%230A192F' stroke='%2338BDF8' stroke-width='5'/>"
    "<circle cx='32' cy='32' r='13' fill='%23112240' stroke='%2310B981' stroke-width='4'/>"
    "<circle cx='32' cy='32' r='5' fill='%23F8FAFC'/>"
    "</svg>"
)

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA — STOCK&YA! (HIVIMAR S.A.)
# ==============================================================================
st.set_page_config(
    page_title="STOCK&YA! | Inteligencia y Cuadre de Inventarios - HIVIMAR S.A.",
    page_icon=PAGE_FAVICON_SVG,
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==============================================================================
# PALETA CORPORATIVA HIVIMAR S.A. + ESTÉTICA EJECUTIVA
# ==============================================================================
CORPORATE_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

:root {
    --bg-main: #0A192F;
    --bg-sidebar: #071222;
    --bg-card: #112240;
    --bg-card-hover: #172E54;
    --border-subtle: rgba(100, 181, 246, 0.18);
    --text-primary: #F8FAFC;
    --text-secondary: #CBD5E1;
    --text-muted: #94A3B8;
    --accent-cyan: #38BDF8;
    --accent-blue: #1D4ED8;
    --accent-emerald: #10B981;
    --accent-amber: #F59E0B;
    --accent-coral: #F43F5E;
}

html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    color: var(--text-primary);
}

/* Fondo general azul corporativo */
.stApp {
    background: radial-gradient(circle at 85% 10%, #132C52 0%, #0A192F 48%, #071222 100%) !important;
}

/* Barra lateral ejecutiva */
section[data-testid="stSidebar"] {
    background-color: var(--bg-sidebar) !important;
    border-right: 1px solid var(--border-subtle) !important;
}

section[data-testid="stSidebar"] .stRadio > div {
    gap: 0.35rem;
}

section[data-testid="stSidebar"] .stRadio label {
    background-color: transparent;
    border-radius: 10px;
    padding: 0.65rem 0.9rem !important;
    transition: all 0.2s ease;
    border: 1px solid transparent;
    color: var(--text-secondary) !important;
    font-weight: 600 !important;
}

section[data-testid="stSidebar"] .stRadio label:hover {
    background-color: #112240 !important;
    color: #FFFFFF !important;
    border-color: rgba(56, 189, 248, 0.25);
}

/* Encabezados de sección con iconografía vectorial */
.sec-header {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    margin: 0.4rem 0 1rem 0;
    padding-bottom: 0.55rem;
    border-bottom: 1px solid rgba(56, 189, 248, 0.18);
}
.sec-icon-wrap {
    width: 38px;
    height: 38px;
    border-radius: 10px;
    background: linear-gradient(135deg, rgba(56, 189, 248, 0.18), rgba(16, 185, 129, 0.14));
    border: 1px solid rgba(56, 189, 248, 0.35);
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
}
.sec-title {
    font-size: 1.32rem;
    font-weight: 800;
    color: #F8FAFC;
    letter-spacing: -0.01em;
    margin: 0;
}
.sec-sub {
    font-size: 0.82rem;
    color: #94A3B8;
    margin: 0;
}

/* Tarjeta de Fuente con insignia numérica vectorial */
.source-header {
    display: flex;
    align-items: center;
    gap: 0.65rem;
    margin-bottom: 0.6rem;
}
.num-badge {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 28px;
    height: 28px;
    border-radius: 8px;
    background: linear-gradient(135deg, #0284C7 0%, #1D4ED8 100%);
    border: 1px solid rgba(56, 189, 248, 0.5);
    color: #FFFFFF;
    font-size: 0.82rem;
    font-weight: 800;
    letter-spacing: 0.04em;
    flex-shrink: 0;
}
.source-title-text {
    font-size: 1.05rem;
    font-weight: 700;
    color: #F8FAFC;
}

/* Tarjetas KPI Ejecutivas */
.kpi-card {
    background: linear-gradient(145deg, #112240 0%, #0E1D36 100%);
    border: 1px solid var(--border-subtle);
    border-radius: 14px;
    padding: 1.15rem 1.25rem;
    box-shadow: 0 8px 24px rgba(2, 12, 27, 0.35);
    transition: transform 0.2s ease, border-color 0.2s ease;
    height: 100%;
}
.kpi-card:hover {
    transform: translateY(-2px);
    border-color: rgba(56, 189, 248, 0.45);
}
.kpi-label {
    font-size: 0.78rem;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    color: var(--text-muted);
    font-weight: 700;
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
    margin-top: 0.4rem;
    font-weight: 600;
}
.badge-ok { color: #34D399; }
.badge-warn { color: #FBBF24; }
.badge-err { color: #FB7185; }
.badge-info { color: #38BDF8; }

/* Banner de Bienvenida */
.hero-banner {
    background: linear-gradient(120deg, #0F294A 0%, #163A66 55%, #0D213B 100%);
    border: 1px solid rgba(56, 189, 248, 0.3);
    border-radius: 18px;
    padding: 1.5rem 2rem;
    margin-bottom: 1.25rem;
    box-shadow: 0 12px 32px rgba(2, 12, 27, 0.45);
}

/* Alerta ¡AGUANTA! */
.aguanta-box {
    background: linear-gradient(135deg, rgba(245, 158, 11, 0.14) 0%, rgba(17, 34, 64, 0.96) 100%);
    border: 1.5px solid #F59E0B;
    border-radius: 14px;
    padding: 1.2rem 1.45rem;
    margin: 0.75rem 0 1rem 0;
    display: flex;
    align-items: center;
    gap: 1rem;
}

/* Botones principales */
.stButton > button {
    border-radius: 10px !important;
    font-weight: 700 !important;
    letter-spacing: 0.02em;
    transition: all 0.2s ease !important;
}

/* Botón maestro &YA! */
div[data-testid="stDownloadButton"] > button {
    background: linear-gradient(90deg, #0284C7 0%, #10B981 100%) !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 12px !important;
    font-weight: 800 !important;
    font-size: 0.98rem !important;
    padding: 0.65rem 1.3rem !important;
    box-shadow: 0 6px 20px rgba(16, 185, 129, 0.35) !important;
}
div[data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 8px 25px rgba(56, 189, 248, 0.5) !important;
}

/* Tablas y Pestañas */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    background-color: #0E1F3A;
    padding: 6px;
    border-radius: 12px;
    border: 1px solid var(--border-subtle);
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px;
    color: #CBD5E1;
    font-weight: 600;
    padding: 8px 16px;
}
.stTabs [aria-selected="true"] {
    background-color: #1D4ED8 !important;
    color: #FFFFFF !important;
}

/* Pie de página con derechos de autor */
.copyright-footer {
    margin-top: 3rem;
    padding: 0.9rem 1rem;
    border-top: 1px solid rgba(148, 163, 184, 0.18);
    text-align: center;
    font-size: 0.74rem;
    color: #94A3B8;
    letter-spacing: 0.03em;
}
</style>
"""

# ==============================================================================
# ICONOS VECTORIALES PROFESIONALES (SVG)
# ==============================================================================
HIVIMAR_LOGO_SVG = """
<svg width="78" height="78" viewBox="0 0 140 140" fill="none" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="archGrad" x1="15" y1="20" x2="125" y2="85" gradientUnits="userSpaceOnUse">
      <stop offset="0%" stop-color="#38BDF8"/>
      <stop offset="50%" stop-color="#0EA5E9"/>
      <stop offset="100%" stop-color="#10B981"/>
    </linearGradient>
    <linearGradient id="steelGrad" x1="30" y1="30" x2="110" y2="110" gradientUnits="userSpaceOnUse">
      <stop offset="0%" stop-color="#F8FAFC"/>
      <stop offset="50%" stop-color="#94A3B8"/>
      <stop offset="100%" stop-color="#475569"/>
    </linearGradient>
  </defs>
  <path d="M 20 82 A 52 52 0 0 1 120 82" stroke="url(#archGrad)" stroke-width="7" stroke-linecap="round"/>
  <path d="M 30 82 A 42 42 0 0 1 110 82" stroke="#38BDF8" stroke-width="2" stroke-dasharray="5 4" opacity="0.7"/>
  <circle cx="20" cy="82" r="4" fill="#38BDF8"/>
  <circle cx="120" cy="82" r="4" fill="#10B981"/>
  <circle cx="70" cy="30" r="4.5" fill="#F8FAFC"/>

  <g transform="rotate(-45 70 74)">
    <rect x="64" y="24" width="12" height="96" rx="3" fill="url(#steelGrad)" opacity="0.9"/>
    <rect x="55" y="18" width="30" height="16" rx="4" fill="#E2E8F0" stroke="#38BDF8" stroke-width="2"/>
    <line x1="56" y1="23" x2="84" y2="23" stroke="#0F172A" stroke-width="1.8"/>
    <line x1="56" y1="28" x2="84" y2="28" stroke="#0F172A" stroke-width="1.8"/>
    <circle cx="70" cy="116" r="8" fill="#64748B" stroke="#E2E8F0" stroke-width="2"/>
  </g>

  <g transform="rotate(45 70 74)">
    <rect x="64" y="24" width="12" height="96" rx="3" fill="url(#steelGrad)" opacity="0.9"/>
    <rect x="55" y="18" width="30" height="16" rx="4" fill="#E2E8F0" stroke="#10B981" stroke-width="2"/>
    <line x1="56" y1="23" x2="84" y2="23" stroke="#0F172A" stroke-width="1.8"/>
    <line x1="56" y1="28" x2="84" y2="28" stroke="#0F172A" stroke-width="1.8"/>
    <circle cx="70" cy="116" r="8" fill="#64748B" stroke="#E2E8F0" stroke-width="2"/>
  </g>

  <circle cx="70" cy="74" r="26" fill="#0A192F" stroke="url(#steelGrad)" stroke-width="5"/>
  <circle cx="70" cy="74" r="14" fill="#112240" stroke="#38BDF8" stroke-width="3.5"/>
  <circle cx="70" cy="74" r="6" fill="#38BDF8"/>
  <circle cx="70" cy="54" r="3" fill="#F8FAFC"/>
  <circle cx="70" cy="94" r="3" fill="#F8FAFC"/>
  <circle cx="50" cy="74" r="3" fill="#F8FAFC"/>
  <circle cx="90" cy="74" r="3" fill="#F8FAFC"/>
  <circle cx="56" cy="60" r="3" fill="#38BDF8"/>
  <circle cx="84" cy="88" r="3" fill="#38BDF8"/>
  <circle cx="84" cy="60" r="3" fill="#10B981"/>
  <circle cx="56" cy="88" r="3" fill="#10B981"/>
</svg>
"""

SVG_ICONS = {
    "upload": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#38BDF8" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>""",
    "audit": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#38BDF8" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v18"/><path d="M5 7h14"/><path d="M5 7l-3 7h6l-3-7z"/><path d="M19 7l-3 7h6l-3-7z"/><line x1="8" y1="21" x2="16" y2="21"/></svg>""",
    "ranking": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#10B981" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>""",
    "trends": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#38BDF8" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>""",
    "ai": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#10B981" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="16" height="16" rx="2"/><rect x="9" y="9" width="6" height="6"/><line x1="9" y1="1" x2="9" y2="4"/><line x1="15" y1="1" x2="15" y2="4"/><line x1="9" y1="20" x2="9" y2="23"/><line x1="15" y1="20" x2="15" y2="23"/><line x1="20" y1="9" x2="23" y2="9"/><line x1="20" y1="14" x2="23" y2="14"/><line x1="1" y1="9" x2="4" y2="9"/><line x1="1" y1="14" x2="4" y2="14"/></svg>""",
    "table": """<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#38BDF8" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="9" x2="9" y2="21"/></svg>""",
    "alert": """<svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#FBBF24" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>""",
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


def render_source_badge_header(number_str: str, title: str):
    st.markdown(
        f"""
        <div class="source-header">
          <span class="num-badge">{number_str}</span>
          <span class="source-title-text">{title}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ==============================================================================
# DETECCIÓN AUTOMÁTICA DE RUBRO SEGÚN NOMBRE DE ARCHIVO Y CONTENIDO
# ==============================================================================
def detect_industry_sector(df_list: list, filenames: list = None) -> tuple:
    """
    Analiza tanto los nombres de los archivos/hojas cargados como los elementos
    (productos, categorías y encabezados) para identificar el sector automáticamente.
    """
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
        # El nombre del archivo pesa x5 para priorizar la intención directa del documento
        file_hits = sum(5 for kw in keywords if kw in file_corpus)
        content_hits = sum(data_corpus.count(kw) for kw in keywords)
        total_score = file_hits + content_hits
        if total_score > best_score:
            best_score = total_score
            best_key = key
            best_label = label

    return best_key, best_label


def render_dynamic_loader_animation(sector_key: str, sector_label: str, analyzed_files: list = None) -> str:
    """
    Notificación vectorial dinámica durante la carga:
    Muestra 'DAME CHANCE QUE ESTOY CARGANDO...' con guiño vectorial inicial
    que transiciona a ruedas/elementos vectoriales del rubro girando en pista.
    """
    item_svgs = {
        "agro_alimentos": """
            <g class="wheel-spin">
              <ellipse cx="0" cy="0" rx="22" ry="16" fill="#D97706" stroke="#FDE68A" stroke-width="2.5"/>
              <circle cx="-8" cy="-4" r="2" fill="#92400E"/>
              <circle cx="7" cy="5" r="2.2" fill="#92400E"/>
              <circle cx="4" cy="-6" r="1.5" fill="#78350F"/>
              <circle cx="-5" cy="6" r="1.5" fill="#78350F"/>
            </g>
        """,
        "farmaceutico": """
            <g class="wheel-spin">
              <rect x="-22" y="-11" width="44" height="22" rx="11" fill="#F8FAFC" stroke="#38BDF8" stroke-width="2"/>
              <path d="M 0 -11 L 11 -11 A 11 11 0 0 1 11 11 L 0 11 Z" fill="#F43F5E"/>
              <line x1="0" y1="-11" x2="0" y2="11" stroke="#0F172A" stroke-width="1.5"/>
            </g>
        """,
        "dulceria": """
            <g class="wheel-spin">
              <circle cx="0" cy="0" r="19" fill="#EC4899" stroke="#FBCFE8" stroke-width="3"/>
              <circle cx="0" cy="0" r="7" fill="#0A192F"/>
              <circle cx="-11" cy="-6" r="2" fill="#FDE047"/>
              <circle cx="10" cy="-7" r="2" fill="#38BDF8"/>
              <circle cx="8" cy="10" r="2" fill="#34D399"/>
              <circle cx="-9" cy="9" r="2" fill="#FFFFFF"/>
            </g>
        """,
        "ferreteria": """
            <g class="wheel-spin">
              <polygon points="0,-21 18,-10 18,10 0,21 -18,10 -18,-10" fill="#94A3B8" stroke="#F8FAFC" stroke-width="2.5"/>
              <circle cx="0" cy="0" r="9" fill="#0A192F" stroke="#38BDF8" stroke-width="2"/>
            </g>
        """,
        "automotriz": """
            <g class="wheel-spin">
              <circle cx="0" cy="0" r="20" fill="#1E293B" stroke="#E2E8F0" stroke-width="3.5"/>
              <circle cx="0" cy="0" r="10" fill="#0F172A" stroke="#38BDF8" stroke-width="2.5"/>
              <circle cx="0" cy="-15" r="2.6" fill="#38BDF8"/>
              <circle cx="0" cy="15" r="2.6" fill="#38BDF8"/>
              <circle cx="-15" cy="0" r="2.6" fill="#10B981"/>
              <circle cx="15" cy="0" r="2.6" fill="#10B981"/>
              <circle cx="-10.6" cy="-10.6" r="2.6" fill="#F8FAFC"/>
              <circle cx="10.6" cy="10.6" r="2.6" fill="#F8FAFC"/>
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
    <div style="background: linear-gradient(135deg, #0F2747 0%, #163863 100%);
                border: 1.5px solid #38BDF8; border-radius: 16px; padding: 1.4rem 1.8rem;
                margin: 1rem 0; text-align: center; box-shadow: 0 10px 30px rgba(0,0,0,0.4);">
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

      <div style="font-size: 1.32rem; font-weight: 800; color: #38BDF8; letter-spacing: 0.04em; margin-bottom: 0.3rem;">
        DAME CHANCE QUE ESTOY CARGANDO...
      </div>
      <div style="font-size: 0.85rem; color: #CBD5E1; margin-bottom: 0.75rem;">
        {files_caption}Rubro identificado automáticamente: <strong style="color:#34D399;">{sector_label}</strong>
      </div>

      <svg width="100%" height="74" viewBox="0 0 420 74" style="max-width: 520px; overflow: visible;">
        <g class="wink-group">
          <circle cx="210" cy="36" r="26" fill="#1E3A8A" stroke="#38BDF8" stroke-width="3"/>
          <circle cx="201" cy="30" r="3.5" fill="#F8FAFC"/>
          <path d="M 215 30 Q 221 25 226 30" stroke="#34D399" stroke-width="3.2" stroke-linecap="round" fill="none"/>
          <path d="M 198 43 Q 210 53 222 43" stroke="#F8FAFC" stroke-width="3" stroke-linecap="round" fill="none"/>
        </g>

        <g class="track-group">
          <line x1="15" y1="58" x2="405" y2="58" stroke="#38BDF8" stroke-width="3" stroke-linecap="round"/>
          <line class="track-line" x1="15" y1="64" x2="405" y2="64" stroke="#10B981" stroke-width="2"/>
          <g class="roller-1">{chosen_svg}</g>
          <g class="roller-2">{chosen_svg}</g>
        </g>
      </svg>
    </div>
    """


# ==============================================================================
# MOTOR AGNÓSTICO DE LECTURA Y DETECCIÓN AUTOMÁTICA DE COLUMNAS
# ==============================================================================
def normalize_text(text) -> str:
    if text is None:
        return ""
    s = str(text).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^a-z0-9_ ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def load_google_sheet_from_url(url: str) -> dict:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not match:
        raise ValueError("Enlace de Google Sheets no válido. Pega la URL completa de la hoja compartida.")
    sheet_id = match.group(1)
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    xls = pd.ExcelFile(export_url)
    return {name: pd.read_excel(xls, sheet_name=name) for name in xls.sheet_names}


def read_uploaded_workbook(uploaded_file) -> dict:
    fname = uploaded_file.name.lower()
    raw_bytes = uploaded_file.getvalue()
    if fname.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(raw_bytes))
        return {"Datos_CSV": df}
    xls = pd.ExcelFile(io.BytesIO(raw_bytes))
    sheets = {}
    for name in xls.sheet_names:
        df_tmp = pd.read_excel(xls, sheet_name=name)
        if not df_tmp.empty:
            sheets[name] = df_tmp
    return sheets


def auto_detect_schema(df: pd.DataFrame) -> dict:
    """
    Identifica automáticamente las columnas presentes.
    Si una columna opcional (ingresos, egresos/ventas, ubicacion, categoria, precio)
    NO existe en el archivo, queda en None y NO se agrega ni se muestra.
    """
    cols = list(df.columns)
    norm_map = {c: normalize_text(c) for c in cols}

    patterns = {
        "codigo": [
            "codigo_sap", "cod_sap", "codigo", "material", "sku", "cod_mat",
            "id_producto", "item", "clave", "referencia", "cod"
        ],
        "descripcion": [
            "producto", "texto breve de material", "texto breve", "descripcion",
            "nombre", "articulo", "detalle", "material_desc", "item_name"
        ],
        "stock": [
            "stock", "librutiliz", "libre utilizacion", "conteo", "existencia",
            "cantidad", "saldo", "inventario", "fisico", "logico", "unidades", "cant"
        ],
        "ubicacion": [
            "lote", "ubicacion", "bodega", "almacen", "rack", "pasillo",
            "percha", "estante", "localizacion", "zona"
        ],
        "categoria": [
            "linea", "categoria", "familia", "grupo", "rubro", "clase",
            "departamento", "tipo", "segmento"
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
            if norm_col.startswith("unnamed"):
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

    return detected


def consolidate_source(df: pd.DataFrame, schema: dict, source_label: str) -> pd.DataFrame:
    """
    Agrupa materiales repetidos en distintas ubicaciones (ej. 1A: 4u + 3A: 2u + 4C: 1u = 7u),
    conserva el detalle de ubicaciones y solo incluye columnas que realmente existan.
    """
    work = df.copy()
    cod_col = schema["codigo"]
    desc_col = schema["descripcion"]
    stk_col = schema["stock"]
    ubi_col = schema["ubicacion"]
    cat_col = schema["categoria"]
    ing_col = schema["ingresos"]
    egr_col = schema["egresos"]
    prc_col = schema["precio"]

    work = work[work[cod_col].notna()].copy()
    work["_COD_KEY"] = work[cod_col].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    work = work[~work["_COD_KEY"].str.lower().isin(["nan", "none", "total", "total general", ""])]

    work["_QTY"] = pd.to_numeric(work[stk_col], errors="coerce").fillna(0.0)

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
        work["_CAT"] = work[cat_col].fillna("General").astype(str).str.strip()
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
        agg_dict["_CAT"] = "first"
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
        grouped.rename(columns={"_CAT": f"CATEGORIA_{source_label}"}, inplace=True)
    if has_ing:
        grouped.rename(columns={"_ING": f"INGRESOS_{source_label}"}, inplace=True)
    if has_egr:
        grouped.rename(columns={"_EGR": f"EGRESOS_{source_label}"}, inplace=True)
    if has_prc:
        grouped.rename(columns={"_PRC": f"PRECIO_{source_label}"}, inplace=True)

    return grouped


def reconcile_two_sources(df_a: pd.DataFrame, df_b: pd.DataFrame, name_a: str, name_b: str) -> pd.DataFrame:
    """Cruza y cuadra ambas fuentes consolidadas con precisión exacta."""
    merged = pd.merge(df_a, df_b, on="CODIGO", how="outer", suffixes=("_A", "_B"))

    desc_a = merged.get("DESCRIPCION_A", pd.Series(dtype=str))
    desc_b = merged.get("DESCRIPCION_B", pd.Series(dtype=str))
    merged["DESCRIPCION"] = desc_a.fillna(desc_b).fillna("Sin descripción")

    cat_cols = [c for c in merged.columns if c.startswith("CATEGORIA_")]
    if cat_cols:
        cat_series = merged[cat_cols[0]]
        for cc in cat_cols[1:]:
            cat_series = cat_series.fillna(merged[cc])
        merged["CATEGORIA"] = cat_series.fillna("General")

    col_stk_a = f"STOCK_{name_a}"
    col_stk_b = f"STOCK_{name_b}"
    merged["EN_FUENTE_A"] = merged[col_stk_a].notna()
    merged["EN_FUENTE_B"] = merged[col_stk_b].notna()

    merged[col_stk_a] = merged[col_stk_a].fillna(0.0)
    merged[col_stk_b] = merged[col_stk_b].fillna(0.0)
    merged["DIFERENCIA"] = merged[col_stk_b] - merged[col_stk_a]

    def classify_status(row):
        if not row["EN_FUENTE_A"]:
            return f"Solo en {name_b}"
        if not row["EN_FUENTE_B"]:
            return f"Solo en {name_a}"
        if abs(row["DIFERENCIA"]) < 1e-9:
            return "Cuadrado"
        return "Sobrante (+)" if row["DIFERENCIA"] > 0 else "Faltante (-)"

    merged["ESTADO_CUADRE"] = merged.apply(classify_status, axis=1)

    ordered_cols = ["CODIGO", "DESCRIPCION"]
    if "CATEGORIA" in merged.columns:
        ordered_cols.append("CATEGORIA")
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
        "DIFERENCIA",
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
# EXPORTADOR MAESTRO EXCEL (&YA!) CON FORMATO CORPORATIVO
# ==============================================================================
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
    ws_res = wb.active
    ws_res.title = "RESUMEN_EJECUTIVO"

    header_fill = PatternFill(start_color="0F294A", end_color="0F294A", fill_type="solid")
    sub_fill = PatternFill(start_color="1D4ED8", end_color="1D4ED8", fill_type="solid")
    zebra_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    ok_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    err_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

    font_title = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=11, bold=True, color="0F172A")
    font_regular = Font(name="Calibri", size=10, color="1E293B")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    ws_res.merge_cells("A1:E1")
    ws_res["A1"] = "STOCK&YA! — REPORTE EJECUTIVO DE CONSOLIDACIÓN Y AUDITORÍA (HIVIMAR S.A.)"
    ws_res["A1"].fill = header_fill
    ws_res["A1"].font = font_title
    ws_res["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_res.row_dimensions[1].height = 30

    ws_res["A2"] = f"Fecha de generación: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws_res["D2"] = f"Sector detectado: {sector_label}"
    ws_res["A3"] = "Autoría: Derechos de autor reservados por John Aarón Domínguez Jara 01.593.321"
    ws_res["A2"].font = font_bold
    ws_res["D2"].font = font_bold
    ws_res["A3"].font = Font(name="Calibri", size=9, italic=True, color="475569")

    row_cursor = 5
    if df_reconciled is not None and not df_reconciled.empty:
        total_skus = len(df_reconciled)
        cuadrados = int((df_reconciled["ESTADO_CUADRE"] == "Cuadrado").sum())
        descuadrados = total_skus - cuadrados
        exactitud = (cuadrados / total_skus * 100.0) if total_skus > 0 else 0.0
        tot_a = float(df_reconciled[f"STOCK_{name_a}"].sum())
        tot_b = float(df_reconciled[f"STOCK_{name_b}"].sum())

        kpis = [
            ("Total Materiales Únicos Auditados", total_skus),
            ("Materiales Cuadrados Exactos", cuadrados),
            ("Materiales con Diferencia / Descuadrados", descuadrados),
            ("Porcentaje de Exactitud de Inventario (%)", round(exactitud, 2)),
            (f"Unidades Totales ({name_a})", tot_a),
            (f"Unidades Totales ({name_b})", tot_b),
            ("Diferencia Neta Global", tot_b - tot_a),
        ]
        ws_res.cell(row=row_cursor, column=1, value="INDICADOR EJECUTIVO").fill = sub_fill
        ws_res.cell(row=row_cursor, column=1).font = font_header
        ws_res.cell(row=row_cursor, column=2, value="VALOR").fill = sub_fill
        ws_res.cell(row=row_cursor, column=2).font = font_header
        row_cursor += 1
        for k_name, k_val in kpis:
            c1 = ws_res.cell(row=row_cursor, column=1, value=k_name)
            c2 = ws_res.cell(row=row_cursor, column=2, value=k_val)
            c1.font = font_bold
            c2.font = font_regular
            c1.border = thin_border
            c2.border = thin_border
            row_cursor += 1

    def write_df_sheet(title: str, df_data: pd.DataFrame):
        ws = wb.create_sheet(title=title)
        ws.freeze_panes = "A2"
        for c_idx, col_name in enumerate(df_data.columns, start=1):
            cell = ws.cell(row=1, column=c_idx, value=str(col_name))
            cell.fill = header_fill
            cell.font = font_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        for r_idx, row_vals in enumerate(df_data.itertuples(index=False), start=2):
            for c_idx, val in enumerate(row_vals, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.font = font_regular
                cell.border = thin_border
                if r_idx % 2 == 0:
                    cell.fill = zebra_fill
                col_title = str(df_data.columns[c_idx - 1])
                if col_title == "ESTADO_CUADRE":
                    if val == "Cuadrado":
                        cell.fill = ok_fill
                        cell.font = Font(name="Calibri", size=10, bold=True, color="166534")
                    else:
                        cell.fill = err_fill
                        cell.font = Font(name="Calibri", size=10, bold=True, color="991B1B")

        for col_idx in range(1, len(df_data.columns) + 1):
            col_letter = get_column_letter(col_idx)
            max_len = max(
                len(str(df_data.columns[col_idx - 1])),
                int(df_data.iloc[:, col_idx - 1].astype(str).str.len().max()) if len(df_data) > 0 else 10,
            )
            ws.column_dimensions[col_letter].width = min(max_len + 4, 48)

    if df_reconciled is not None and not df_reconciled.empty:
        write_df_sheet("CUADRE_CONSOLIDADO", df_reconciled)
        df_diff_only = df_reconciled[df_reconciled["ESTADO_CUADRE"] != "Cuadrado"]
        if not df_diff_only.empty:
            write_df_sheet("SOLO_DESCUADRADOS", df_diff_only)

    if df_top_sales is not None and not df_top_sales.empty:
        write_df_sheet("TOP20_MAS_MOVIMIENTO", df_top_sales)

    if df_no_move is not None and not df_no_move.empty:
        write_df_sheet("TOP20_SIN_MOVIMIENTO", df_no_move)

    if df_ai is not None and not df_ai.empty:
        write_df_sheet("PROYECCION_IA_12M", df_ai)

    ws_res.column_dimensions["A"].width = 44
    ws_res.column_dimensions["B"].width = 22
    ws_res.column_dimensions["D"].width = 36

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


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

    @st.dialog("Bienvenido a STOCK&YA! — Guía Rápida Ejecutiva", width="large")
    def show_interactive_guide():
        st.markdown(
            """
            ### Plataforma inteligente de consolidación y auditoría para **HIVIMAR S.A.**
            Diseñada para eliminar fórmulas manuales en Excel y unificar reportes de **SAP, Oracle, Odoo, Microsoft Dynamics, CSV o Google Sheets** en segundos.

            #### Módulos operativos de STOCK&YA!
            1. **Cruce y Cuadre (2 Fuentes):** Sube dos archivos (o dos hojas de un mismo libro) y confirma con **¡SIMÓN!**. El motor agrupa materiales repetidos en distintas ubicaciones (ej. `1A + 3A + 4C`) y verifica exactamente qué está **Cuadrado** o **Descuadrado**.
            2. **Top 20 Rotación y Sin Movimiento:** Visualiza al instante los 20 productos con mayor salida y los 20 productos inmovilizados con sus porcentajes reales.
            3. **Termómetro Creciente / Decreciente (%):** Analiza variaciones porcentuales exactas en ventas e inventario por línea y por material.
            4. **Consolidado 12 Meses + IA Predictiva:** Carga hasta **12 archivos mensuales** para consolidar el año completo y proyectar con Inteligencia Artificial cuánto se estima vender el próximo mes.
            5. **Descarga Ejecutiva `&YA!`:** Exporta en un clic todo el reporte ejecutivo en Excel listo para dirección.
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
                  <div style="font-size: 1.5rem; font-weight: 800; color: #FFFFFF; letter-spacing: -0.02em;">
                    STOCK<span style="color:#38BDF8;">&amp;</span><span style="color:#10B981;">YA!</span>
                  </div>
                  <div style="font-size: 0.75rem; font-weight: 700; color: #94A3B8; letter-spacing: 0.08em;">
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
                "Auditoría y Cuadre (Lógico vs Físico)",
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
            <div style="margin-top: 2rem; font-size: 0.72rem; color: #64748B; line-height: 1.45;">
              Compatible con Excel (.xlsx/.xls), CSV y Google Sheets.<br>
              Detección automática de rubro y estructura activada.
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown(
        """
        <div class="hero-banner">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem;">
            <div>
              <span style="background:rgba(56,189,248,0.18); color:#38BDF8; padding:4px 12px; border-radius:999px; font-size:0.75rem; font-weight:700; letter-spacing:0.06em;">
                PLATAFORMA EJECUTIVA DE CONSOLIDACIÓN &bull; HIVIMAR S.A.
              </span>
              <h1 style="margin:0.45rem 0 0.25rem 0; font-size:2rem; font-weight:800; color:#FFFFFF;">
                Bienvenido a STOCK<span style="color:#38BDF8;">&amp;</span><span style="color:#10B981;">YA!</span>
              </h1>
              <p style="margin:0; color:#CBD5E1; font-size:0.95rem;">
                Consolida ubicaciones múltiples, audita diferencias entre sistemas (SAP, Oracle, Odoo, Excel, Google Sheets) y proyecta demanda con IA en un solo clic.
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
            "Carga de Fuentes para Consolidación y Cuadre",
            "El motor identifica automáticamente el rubro según el nombre del archivo y los elementos cargados.",
        )
        col_src1, col_src2 = st.columns(2)

        sheets_1 = {}
        sheets_2 = {}
        fname_1 = ""
        fname_2 = ""

        with col_src1:
            render_source_badge_header("01", "Primera Fuente (Ej. SAP / Oracle / Sistema Lógico)")
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
                        fname_1 = "Google_Sheets_Fuente_1"
                        sheets_1 = load_google_sheet_from_url(gs_url1)
                    except Exception as exc:
                        st.error(f"Error al leer Google Sheets 1: {exc}")

            if sheets_1:
                sel_sheet_1 = st.selectbox(
                    "Pestaña a analizar (Fuente 1):",
                    list(sheets_1.keys()),
                    index=min(1, len(sheets_1) - 1) if "ORACLE" in sheets_1 or "SAP" in sheets_1 else 0,
                    key="sel_s1",
                )
                name_a = st.text_input("Etiqueta Fuente 1:", value=sel_sheet_1.upper(), key="lbl_a")

        with col_src2:
            render_source_badge_header("02", "Segunda Fuente (Ej. Odoo / Microsoft Dynamics / Conteo Físico)")
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
                        fname_2 = "Google_Sheets_Fuente_2"
                        sheets_2 = load_google_sheet_from_url(gs_url2)
                    except Exception as exc:
                        st.error(f"Error al leer Google Sheets 2: {exc}")

            if sheets_2:
                default_idx_2 = 2 if len(sheets_2) > 2 else (1 if len(sheets_2) > 1 else 0)
                sel_sheet_2 = st.selectbox(
                    "Pestaña a analizar (Fuente 2):",
                    list(sheets_2.keys()),
                    index=default_idx_2,
                    key="sel_s2",
                )
                name_b = st.text_input("Etiqueta Fuente 2:", value=sel_sheet_2.upper(), key="lbl_b")

        if sheets_1 and sheets_2:
            df_raw_1 = sheets_1[sel_sheet_1]
            df_raw_2 = sheets_2[sel_sheet_2]
            sig = f"{fname_1}_{sel_sheet_1}_{len(df_raw_1)}_{fname_2}_{sel_sheet_2}_{len(df_raw_2)}_{name_a}_{name_b}"

            if sig != st.session_state.file_signature:
                st.session_state.confirmed_load = False

            if not st.session_state.confirmed_load:
                st.markdown(
                    f"""
                    <div class="aguanta-box">
                      <div style="flex-shrink:0;">{SVG_ICONS["alert"]}</div>
                      <div>
                        <div style="font-size:1.22rem; font-weight:800; color:#FBBF24;">
                          ¡AGUANTA! ¿Seguro que el archivo que quieres cargar es el correcto?
                        </div>
                        <div style="color:#E2E8F0; font-size:0.92rem; margin-top:0.25rem;">
                          Vas a consolidar <strong>{name_a}</strong> ({len(df_raw_1):,} filas) contra <strong>{name_b}</strong> ({len(df_raw_2):,} filas).
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
                    time.sleep(1.7)
                    anim_placeholder.empty()
                    st.session_state.last_animated_sig = sig

                schema_1 = auto_detect_schema(df_raw_1)
                schema_2 = auto_detect_schema(df_raw_2)

                with st.expander("Columnas detectadas automáticamente (solo se incluyen las que existen)"):
                    c_det1, c_det2 = st.columns(2)
                    with c_det1:
                        st.markdown(f"**Esquema {name_a}:**")
                        st.json({k: v for k, v in schema_1.items() if v is not None})
                    with c_det2:
                        st.markdown(f"**Esquema {name_b}:**")
                        st.json({k: v for k, v in schema_2.items() if v is not None})

                df_cons_1 = consolidate_source(df_raw_1, schema_1, name_a)
                df_cons_2 = consolidate_source(df_raw_2, schema_2, name_b)
                df_reconciled = reconcile_two_sources(df_cons_1, df_cons_2, name_a, name_b)

                st.session_state.df_reconciled = df_reconciled
                st.session_state.name_a = name_a
                st.session_state.name_b = name_b
                st.session_state.sector_key = sector_key
                st.session_state.sector_label = sector_label

                st.success(
                    f"Consolidación completada. Rubro detectado: **{sector_label}**. "
                    f"Se agruparon ubicaciones múltiples y se auditaron **{len(df_reconciled):,} materiales únicos**."
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
        descuadre_pct = 100.0 - exactitud_pct

        sum_a = float(df_reconciled[col_a].sum())
        sum_b = float(df_reconciled[col_b].sum())
        var_inv_pct = ((sum_b - sum_a) / sum_a * 100.0) if sum_a > 0 else 0.0

        c_info, c_export = st.columns([3, 1.2])
        with c_info:
            st.caption(
                f"**Archivo Activo:** `{name_a}` vs `{name_b}`  |  **Rubro:** {sector_label}  |  "
                f"**SKUs Únicos:** {total_skus:,}"
            )
        with c_export:
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
                label="Descargar Todo &YA!",
                data=excel_bytes,
                file_name=f"STOCK_Y_YA_HIVIMAR_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        if nav_option in ["Inicio y Carga Inteligente", "Auditoría y Cuadre (Lógico vs Físico)"]:
            render_section_header(
                "audit",
                "Auditoría de Consolidación y Cuadre Exacto",
                "Comparativo consolidado por código de material y suma de ubicaciones.",
            )

            k1, k2, k3, k4 = st.columns(4)
            with k1:
                st.markdown(
                    f"""
                    <div class="kpi-card">
                      <div class="kpi-label">Exactitud de Inventario</div>
                      <div class="kpi-value">{exactitud_pct:.2f}%</div>
                      <div class="kpi-sub badge-ok">{cuadrados_cnt:,} de {total_skus:,} materiales cuadrados</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with k2:
                st.markdown(
                    f"""
                    <div class="kpi-card">
                      <div class="kpi-label">Materiales Descuadrados</div>
                      <div class="kpi-value">{descuadrados_cnt:,}</div>
                      <div class="kpi-sub {'badge-ok' if descuadrados_cnt == 0 else 'badge-err'}">
                        {descuadre_pct:.2f}% del catálogo con novedad
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with k3:
                st.markdown(
                    f"""
                    <div class="kpi-card">
                      <div class="kpi-label">Unidades {name_a} vs {name_b}</div>
                      <div class="kpi-value">{sum_b:,.0f}</div>
                      <div class="kpi-sub badge-info">Base {name_a}: {sum_a:,.0f} uds</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with k4:
                net_diff = sum_b - sum_a
                st.markdown(
                    f"""
                    <div class="kpi-card">
                      <div class="kpi-label">Variación / Diferencia Neta</div>
                      <div class="kpi-value">{net_diff:+,.0f} uds</div>
                      <div class="kpi-sub {'badge-ok' if net_diff >= 0 else 'badge-warn'}">
                        {var_inv_pct:+.2f}% respecto a {name_a}
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            st.markdown("<br>", unsafe_allow_html=True)

            f_col1, f_col2 = st.columns([1.5, 2])
            with f_col1:
                filtro_estado = st.selectbox(
                    "Filtrar auditoría:",
                    [
                        "Todos los materiales",
                        "Solo Descuadrados / Con Diferencia",
                        "Solo Cuadrados Exactos",
                        "Materiales Multi-Ubicación (Agrupados)",
                    ],
                )
            with f_col2:
                busqueda_sku = st.text_input("Buscar por Código o Descripción:", placeholder="Ej. 1001747 o SKF...")

            df_view = df_reconciled.copy()
            if filtro_estado == "Solo Descuadrados / Con Diferencia":
                df_view = df_view[df_view["ESTADO_CUADRE"] != "Cuadrado"]
            elif filtro_estado == "Solo Cuadrados Exactos":
                df_view = df_view[df_view["ESTADO_CUADRE"] == "Cuadrado"]
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
                        color_continuous_scale=["#0284C7", "#38BDF8", "#10B981"],
                        title=f"Top 20 Productos — {mov_label}",
                    )
                    fig_top.update_traces(texttemplate="%{x:,.0f} uds (%{text}%)", textposition="outside")
                    fig_top.update_layout(
                        paper_bgcolor="rgba(0,0,0,0)",
                        plot_bgcolor="rgba(17,34,64,0.65)",
                        font=dict(color="#F8FAFC"),
                        height=620,
                        showlegend=False,
                        coloraxis_showscale=False,
                        xaxis_title="Unidades",
                        yaxis_title="",
                    )
                    st.plotly_chart(fig_top, use_container_width=True)
                    st.dataframe(
                        df_top_sales[["CODIGO", "DESCRIPCION", col_a, col_b, "UNIDADES_MOVIDAS", "PARTICIPACION_%"]],
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
                        color_continuous_scale=["#F59E0B", "#F43F5E"],
                        title="Top 20 Productos Sin Movimiento (Mayor Stock Estancado)",
                    )
                    fig_dead.update_traces(texttemplate="%{x:,.0f} uds (%{text}%)", textposition="outside")
                    fig_dead.update_layout(
                        paper_bgcolor="rgba(0,0,0,0)",
                        plot_bgcolor="rgba(17,34,64,0.65)",
                        font=dict(color="#F8FAFC"),
                        height=620,
                        coloraxis_showscale=False,
                        xaxis_title="Stock Estancado (Unidades)",
                        yaxis_title="",
                    )
                    st.plotly_chart(fig_dead, use_container_width=True)
                    st.dataframe(
                        df_no_move[["CODIGO", "DESCRIPCION", col_b, "PESO_INVENTARIO_%", "ESTADO_CUADRE"]],
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
                        <div class="kpi-card">
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
                        <div class="kpi-card">
                          <div class="kpi-label">Dinámica del Catálogo (SKUs en Alza vs Baja)</div>
                          <div class="kpi-value">+{pct_sub:.1f}% / -{pct_baj:.1f}%</div>
                          <div class="kpi-sub badge-info">
                            {skus_subieron:,} SKUs crecieron &bull; {skus_bajaron:,} SKUs decrecieron
                          </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

            st.markdown("<br>", unsafe_allow_html=True)

            if "CATEGORIA" in df_reconciled.columns:
                cat_grp = (
                    df_reconciled.groupby("CATEGORIA", as_index=False)[[col_a, col_b]]
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
                    y="CATEGORIA",
                    orientation="h",
                    color="TENDENCIA",
                    color_discrete_map={"Creciente (+)": "#10B981", "Decreciente (-)": "#F43F5E"},
                    text="VARIACION_%",
                    title="Variación Porcentual de Inventario por Línea / Categoría (%)",
                )
                fig_cat.update_traces(texttemplate="%{text:+.2f}%", textposition="outside")
                fig_cat.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(17,34,64,0.65)",
                    font=dict(color="#F8FAFC"),
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
                color_discrete_map={"Creciente (+)": "#10B981", "Decreciente (-)": "#F43F5E"},
                text="VAR_%",
                title="Top Productos con Mayor Variación Porcentual (Creciente vs Decreciente)",
            )
            fig_movers.update_traces(texttemplate="%{text:+.1f}%", textposition="outside")
            fig_movers.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(17,34,64,0.65)",
                font=dict(color="#F8FAFC"),
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
            "Sube los **12 archivos mensuales** (Excel o CSV). El motor los ordenará cronológicamente, "
            "consolidará ubicaciones repetidas y entrenará el modelo predictivo para calcular la **expectativa de venta del próximo mes**."
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
                        <div style="font-size:1.2rem; font-weight:800; color:#FBBF24;">
                          ¡AGUANTA! ¿Seguro que el archivo que quieres cargar es el correcto?
                        </div>
                        <div style="color:#E2E8F0; font-size:0.9rem; margin-top:0.25rem;">
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
                    chosen_sheet = (
                        "ORACLE" if "ORACLE" in s_dict
                        else ("SAP" if "SAP" in s_dict else list(s_dict.keys())[0])
                    )
                    raw_m = s_dict[chosen_sheet]
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
                    time.sleep(1.6)
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
                          <div class="kpi-sub badge-info">Proyección sobre {len(monthly_records)} meses</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with a2:
                    st.markdown(
                        f"""
                        <div class="kpi-card">
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
                        <div class="kpi-card">
                          <div class="kpi-label">Productos en Aceleración</div>
                          <div class="kpi-value">{accel_cnt:,} SKUs</div>
                          <div class="kpi-sub badge-ok">Crecimiento proyectado &ge; +10%</div>
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
                        line=dict(color="#38BDF8", width=3.5),
                        marker=dict(size=9, color="#38BDF8"),
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
                        line=dict(color="#10B981", width=3.5, dash="dash"),
                        marker=dict(size=12, color="#10B981", symbol="diamond"),
                    )
                )
                fig_ai.update_layout(
                    title="Curva Histórica Consolidada y Expectativa de Venta Próximo Mes (IA)",
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(17,34,64,0.65)",
                    font=dict(color="#F8FAFC"),
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
