import streamlit as st
import pandas as pd
from io import BytesIO

from utils import compute_consistency_for_df


st.set_page_config(
    page_title="Consistencia PEI (Objetivo ↔ Actividad)",
    page_icon="📘",
    layout="wide"
)

st.title("📘 Calculadora de Consistencia PEI")
st.write(
    """
    Esta herramienta calcula la **consistencia** entre el **objetivo específico**
    elegido por la unidad académica y la **actividad única** cargada, usando
    similitud semántica (TF-IDF) y devolviendo una nota discreta:

    **0, 10, 30, 50, 70, 90 o 100 (%).**
    """
)

uploaded_file = st.file_uploader(
    "Cargar archivo Excel con Objetivos y Actividades",
    type=["xlsx"]
)

if uploaded_file is None:
    st.info("Subí un archivo .xlsx para comenzar.")
    st.stop()

# 1) Leer el archivo
try:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
except Exception as e:
    st.error(f"No se pudo leer el archivo: {e}")
    st.stop()

st.subheader("Vista previa del archivo original")
st.dataframe(df.head(), use_container_width=True)

if df.empty:
    st.error("El archivo no tiene filas.")
    st.stop()

columns = list(df.columns)

st.markdown("### Seleccioná las columnas correspondientes")

col1, col2 = st.columns(2)

with col1:
    col_obj_codigo = st.selectbox(
        "Columna con el **código / identificador** del objetivo específico",
        options=columns,
        index=0
    )

    col_actividad = st.selectbox(
        "Columna con el texto de la **actividad única**",
        options=columns,
        index=1 if len(columns) > 1 else 0
    )

with col2:
    col_obj_texto = st.selectbox(
        "Columna con el **texto completo** del objetivo específico",
        options=columns,
        index=2 if len(columns) > 2 else 0
    )

    col_detalle = st.selectbox(
        "Columna con el **detalle de la actividad** (si existe)",
        options=["(sin detalle)"] + columns,
        index=0
    )

if col_detalle == "(sin detalle)":
    # Creamos una columna vacía para simplificar la lógica
    df["_DETALLE_VACIO_"] = ""
    col_detalle_real = "_DETALLE_VACIO_"
else:
    col_detalle_real = col_detalle

st.markdown("---")

if st.button("Calcular consistencia", type="primary"):
    with st.spinner("Calculando similitud y consistencia..."):
        result_df = compute_consistency_for_df(
            df,
            col_obj_codigo=col_obj_codigo,
            col_obj_texto=col_obj_texto,
            col_actividad=col_actividad,
            col_detalle=col_detalle_real
        )

    st.subheader("Resultados con Consistencia (%)")
    st.dataframe(result_df.head(50), use_container_width=True)

    # Indicadores básicos
    st.markdown("### Indicadores globales")

    # Promedio de consistencia
    mean_consistency = result_df["Consistencia (%)"].mean()
    st.write(f"**Consistencia promedio:** {mean_consistency:.2f} %")

    # Distribución de valores en la escala discreta
    distrib = (
        result_df["Consistencia (%)"]
        .value_counts()
        .sort_index()
        .rename_axis("Consistencia")
        .reset_index(name="Cantidad")
    )

    st.write("**Distribución de actividades por nivel de consistencia:**")
    st.dataframe(distrib, use_container_width=True)

    # Preparar descarga de Excel
    def to_excel_bytes(df_out: pd.DataFrame) -> bytes:
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_out.to_excel(writer, index=False, sheet_name="Consistencia_PEI")
        return buffer.getvalue()

    excel_bytes = to_excel_bytes(result_df)

    st.download_button(
        label="📥 Descargar resultados en Excel",
        data=excel_bytes,
        file_name="consistencia_pei_objetivo_actividad.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Configurá las columnas y hacé clic en **Calcular consistencia**.")
