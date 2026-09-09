import re
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

from ucc_streamlit_chrome import hide_streamlit_cloud_toolbar

# =========================
# CONFIG
# =========================

st.set_page_config(
    page_title="Carga de Actas - Consejo de Investigación",
    layout="wide"
)
hide_streamlit_cloud_toolbar()

# =========================
# ESTILO
# =========================

st.markdown("""
<style>

[data-testid="stAppViewContainer"] {
    background-color: #D8EBE2 !important;
}

.block-container {
    background-color: transparent !important;
    opacity: 1 !important;
}

section.main > div {
    opacity: 1 !important;
}

h1, h2, h3, p, li {
    color: #1e293b !important;
}

[data-testid="stSidebar"] {
    background-color: #B5D5C6 !important;
}

[data-testid="stSidebar"] > div:first-child {
    background-color: #B5D5C6 !important;
}

[data-testid="stSidebar"] * {
    color: #1e293b !important;
}

.header-uccuyo {
    background-color: #064a3f;
    padding: 20px;
    border-radius: 10px;
}

.header-uccuyo h2,
.header-uccuyo p {
    color: white !important;
    margin: 0;
}

[data-testid="stSidebarNav"] a {
    background-color: #ffffff !important;
    color: #1e293b !important;
    font-size: 15px !important;
    margin-bottom: 8px;
    padding: 10px 12px;
    border-radius: 8px;
    border: 1px solid #9bbbac !important;
}

[data-testid="stSidebarNav"] a[aria-current="page"] {
    background-color: #9fc9b6 !important;
    color: #064a3f !important;
    font-weight: 600 !important;
}

.card {
    background: #ffffff;
    border: 1px solid #c7dfd5;
    border-radius: 12px;
    padding: 16px 18px;
    margin: 12px 0 16px 0;
    box-shadow: 0 2px 8px rgba(2, 40, 33, 0.06);
}

.card h3 {
    color: #064a3f !important;
    margin: 0 0 8px 0;
    font-size: 1.05rem;
}

.card a {
    color: #0b6b5d !important;
    font-weight: 700;
    text-decoration: none;
}

.card a:hover {
    text-decoration: underline;
}

.card.ucc-card-cvar {
    border: 2px solid #0b6b5d;
    background: linear-gradient(180deg, #f3faf7 0%, #ffffff 100%);
}

.no-link {
    color: #b45309 !important;
    font-weight: 700;
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================

_APP_ROOT = Path(__file__).resolve().parents[1]
_LOGO_PATH = _APP_ROOT / "assets" / "logo_uccuyo.png"

col1, col2 = st.columns([1, 6])

with col1:
    if _LOGO_PATH.exists():
        st.image(str(_LOGO_PATH), width=110)
    else:
        st.warning("Logo no encontrado")

with col2:
    st.markdown("""
    <div class='header-uccuyo'>
        <h2>Universidad Católica de Cuyo</h2>
        <p>Secretaría de Investigación</p>
    </div>
    """, unsafe_allow_html=True)

# TÍTULO
st.title("📂 Carga de Archivos para el Consejo de Investigación")

if st.button("↑ Actas", key="archivos_volver_actas"):
    st.session_state["volver_dashboard_actas"] = True
    st.switch_page("pages/0_Actas.py")

st.markdown("### Seleccione la carpeta correspondiente y cargue el archivo directamente en Google Drive")

st.markdown(
    """
    <div class="card ucc-card-cvar" style="margin-top:8px;">
      <h3>📁 Acceso rápido · Categorización CVar 2026</h3>
      <p>🔗 <a href="https://drive.google.com/drive/folders/1kLCX1LIHYpEH8W0ud-5SW39Z_GGGLXmp" target="_blank">Abrir carpeta CVar</a></p>
    </div>
    """,
    unsafe_allow_html=True,
)

# =========================
# LINKS (TODOS CONFIGURADOS)
# =========================

actas = {
    "Categorización CVar 2026": "https://drive.google.com/drive/folders/1kLCX1LIHYpEH8W0ud-5SW39Z_GGGLXmp",
    "Orden del Día Mayo - Acta 190": "https://drive.google.com/drive/folders/1i6amwbRjPYBqCK0gZ2k5sbcBV6mVM0ki",
    "Orden del Día Junio - Acta 191": "https://drive.google.com/drive/folders/118VKqYfacn5eBH0dYXRvM5uAXv2M4flI",
    "Orden del Día Julio - Acta 192": "https://drive.google.com/drive/folders/1T1GID13Wa5TXmTfibQ0HAMBtNmOvC-DH",
    "Orden del Día Agosto - Acta 193": "https://drive.google.com/drive/folders/1VVdUSebRgH2FdivZ6ZceoK30mmqtPYVs",
    "Orden del Día Septiembre - Acta 194": "https://drive.google.com/drive/folders/1XkVQdpn4zNQhy2pLx0SbfTxfDQQw0eoY",
    "Orden del Día Octubre - Acta 195": "https://drive.google.com/drive/folders/1UhA4EHVhHFlfB75bkbsLGYDm4YzmVZI0",
    "Orden del Día Noviembre - Acta 196": "https://drive.google.com/drive/folders/1zy8FMLNbqRUkv9EaIgpCKebf1kgy1nmv",
    "Orden del Día Diciembre - Acta 197": "https://drive.google.com/drive/folders/1pRFUHhn1hxpGDvEirgOjlUf2mB2Ml0Em",
}

CVAR_NOMBRE = "Categorización CVar 2026"

# =========================
# RENDER
# =========================

_acta_focus = st.session_state.get("acta_archivos")
_ancla_archivos = None
if _acta_focus:
    _tiene_carpeta = any(f"Acta {_acta_focus}" in nombre for nombre in actas)
    if _tiene_carpeta:
        st.info(f"Acta {_acta_focus}: use la carpeta de ese Orden del Día para subir el archivo.")
    else:
        st.warning(f"El Acta {_acta_focus} todavía no tiene carpeta de Google Drive habilitada.")

st.markdown('<div id="archivo-cvar"></div>', unsafe_allow_html=True)

for nombre, link in actas.items():
    m_acta = re.search(r"Acta\s+(\d+)", nombre)
    n_acta = int(m_acta.group(1)) if m_acta else None
    es_cvar = nombre == CVAR_NOMBRE
    destacado = (n_acta is not None and n_acta == _acta_focus) or (es_cvar and not _acta_focus)
    if destacado and n_acta is not None:
        _ancla_archivos = f"archivo-acta-{n_acta}"
    elif destacado and es_cvar:
        _ancla_archivos = "archivo-cvar"

    if "XXXXXXXX" in link:
        link_html = "<p class='no-link'>🔒 Carpeta aún no habilitada</p>"
    else:
        link_html = f'<p>🔗 <a href="{link}" target="_blank">Abrir carpeta</a></p>'

    borde = "border:2px solid #064a3f;" if (destacado and not es_cvar) else ""
    clase = "card ucc-card-cvar" if es_cvar else "card"
    ancla = f'<div id="archivo-acta-{n_acta}"></div>' if n_acta else ""
    titulo_extra = " · carga permanente" if es_cvar else ""
    st.markdown(
        f"""
        {ancla}
        <div class="{clase}" style="{borde}">
        <h3>📁 {nombre}{titulo_extra}</h3>
        {link_html}
        <p><strong>Cómo cargar el archivo:</strong></p>
        <ol>
            <li>Se abrirá Google Drive</li>
            <li><strong>Hacer clic en la Carpeta de su Unidad Académica</strong></li>
            <li>Hacer clic en <strong>Nuevo → Subir archivo</strong></li>
            <li>Seleccionar el archivo correspondiente</li>
            <li>Verificar que quede dentro de la carpeta</li>
        </ol>
        </div>
        """,
        unsafe_allow_html=True,
    )

if _ancla_archivos:
    components.html(
        f"""
        <script>
        (function () {{
            const doc = window.parent.document;
            const id = {_ancla_archivos!r};
            function ir() {{
                const el = doc.getElementById(id);
                if (!el) return;
                try {{ el.scrollIntoView({{ behavior: "smooth", block: "start" }}); }} catch (e) {{}}
            }}
            [0, 200, 600, 1200].forEach(function (ms) {{ setTimeout(ir, ms); }});
        }})();
        </script>
        """,
        height=0,
    )
