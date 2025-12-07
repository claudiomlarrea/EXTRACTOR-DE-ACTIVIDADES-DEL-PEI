import streamlit as st
import pandas as pd
from io import BytesIO
from normalizer import normalize_columns

st.set_page_config(page_title="Extractor PEI", page_icon="📘", layout="wide")

st.title("📘 Extractor y Normalizador de Actividades del PEI")

uploaded_file = st.file_uploader("Cargar archivo Excel original", type=["xlsx"])

if not uploaded_file:
    st.stop()

df = pd.read_excel(uploaded_file, engine="openpyxl")

st.subheader("Vista previa del archivo original")
st.dataframe(df.head(), use_container_width=True)

# Normalizamos
clean_df = normalize_columns(df)

st.subheader("Plantilla PEI Generada Automáticamente")
st.dataframe(clean_df.head(), use_container_width=True)

# Descargar Excel
def to_excel(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="PEI_normalizado")
    return buffer.getvalue()

st.download_button(
    label="📥 Descargar plantilla PEI (Excel)",
    data=to_excel(clean_df),
    file_name="pei_actividades_filtradas.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
