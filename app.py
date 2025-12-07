import streamlit as st
import pandas as pd
from io import BytesIO

from normalizer import normalize_columns
from consistency import add_consistency_to_clean, add_consistency_to_resumen


st.set_page_config(
    page_title="Extractor PEI + Consistencia",
    page_icon="📘",
    layout="wide"
)


def build_resumen(clean_df: pd.DataFrame) -> pd.DataFrame:
    """
    Construye una tabla 'larga' con:
    ID - Año - Objetivo específico - Actividad relacionada
    a partir de la plantilla normalizada (Objetivo 1..6, Actividad Obj 1..6).
    """
    filas = []

    for obj in range(1, 7):
        col_id = clean_df["ID"]
        col_year = clean_df["Año"]
        col_obj = clean_df[f"Objetivo {obj}"]
        col_act = clean_df[f"Actividad Obj {obj}"]

        tmp = pd.DataFrame(
            {
                "ID": col_id,
                "Año": col_year,
                "Objetivo específico": col_obj,
                "Actividad relacionada": col_act,
            }
        )

        # Filtrar filas sin actividad
        mask = tmp["Actividad relacionada"].astype(str).str.strip() != ""
        tmp = tmp[mask]

        filas.append(tmp)

    if filas:
        resumen = pd.concat(filas, ignore_index=True)
    else:
        resumen = pd.DataFrame(
            columns=["ID", "Año", "Objetivo específico", "Actividad relacionada"]
        )

    return resumen


def build_indicators(resumen_df: pd.DataFrame) -> pd.DataFrame:
    """
    Crea una tabla pequeña de indicadores globales:
    - Cantidad total de actividades únicas (por ID)
    - Consistencia general promedio (%)
    """
    # 1) Cantidad de actividades únicas por ID
    if "ID" in resumen_df.columns:
        total_actividades_unicas = resumen_df["ID"].nunique()
    else:
        # fallback por si faltara la columna (no debería suceder)
        actividades = resumen_df["Actividad relacionada"].astype(str).str.strip()
        total_actividades_unicas = actividades[actividades != ""].nunique()

    # 2) Promedio de consistencia general
    if "Consistencia (%)" in resumen_df.columns:
        consistencia_general = resumen_df["Consistencia (%)"].mean()
    else:
        consistencia_general = 0.0

    indicadores = pd.DataFrame(
        {
            "Indicador": [
                "Cantidad total de actividades únicas",
                "Consistencia general (%)",
            ],
            "Valor": [
                total_actividades_unicas,
                round(consistencia_general, 2),
            ],
        }
    )

    return indicadores


def to_excel_three_sheets(
    plantilla: pd.DataFrame,
    resumen: pd.DataFrame,
    indicadores: pd.DataFrame
) -> bytes:
    """
    Exporta las tres hojas:
    - PEI_normalizado
    - Resumen
    - Indicadores
    en un único archivo Excel.
    """
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        plantilla.to_excel(writer, index=False, sheet_name="PEI_normalizado")
        resumen.to_excel(writer, index=False, sheet_name="Resumen")
        indicadores.to_excel(writer, index=False, sheet_name="Indicadores")
    return buffer.getvalue()


# =========================
#       INTERFAZ UI
# =========================

st.title("📘 Extractor PEI + Consistencia de Objetivos-Actividades")

uploaded_file = st.file_uploader(
    "Cargar archivo Excel original (Formulario Único / Looker Studio)",
    type=["xlsx"]
)

if uploaded_file is None:
    st.info("Subí un archivo .xlsx para comenzar.")
    st.stop()

# 1) Leer archivo original
try:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
except Exception as e:
    st.error(f"No se pudo leer el archivo: {e}")
    st.stop()

st.subheader("Vista previa del archivo original")
st.dataframe(df.head(), use_container_width=True)

# 2) Normalizar columnas -> plantilla estándar
clean_df = normalize_columns(df)

# 2.b) Agregar consistencia por objetivo a la plantilla
clean_df = add_consistency_to_clean(clean_df)

st.subheader("Hoja 1: PEI_normalizado + Consistencia por objetivo")
st.dataframe(clean_df.head(), use_container_width=True)

# 3) Construir hoja de RESUMEN (ID - Año - Objetivo - Actividad)
resumen_df = build_resumen(clean_df)

# 3.b) Agregar consistencia a la hoja Resumen
resumen_df = add_consistency_to_resumen(resumen_df)

st.subheader("Hoja 2: Resumen (ID - Año - Objetivo - Actividad + Consistencia)")
st.dataframe(resumen_df.head(), use_container_width=True)

# 4) Construir hoja de INDICADORES
indicadores_df = build_indicators(resumen_df)

st.subheader("Hoja 3: Indicadores globales")
st.dataframe(indicadores_df, use_container_width=True)

# 5) Exportar las TRES hojas en un solo Excel
excel_bytes = to_excel_three_sheets(clean_df, resumen_df, indicadores_df)

st.subheader("Descarga del archivo Excel final")

st.download_button(
    label="📥 Descargar Excel con 3 hojas (PEI_normalizado, Resumen, Indicadores)",
    data=excel_bytes,
    file_name="pei_actividades_consistencia.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
