import streamlit as st
import pandas as pd
from io import BytesIO

from normalizer import normalize_columns
from consistency import add_consistency_to_clean, add_consistency_to_resumen

st.set_page_config(page_title="Extractor PEI", page_icon="📘", layout="wide")

st.title("📘 Extractor PEI + Consistencia de Objetivos-Actividades")

uploaded_file = st.file_uploader("Cargar archivo Excel original", type=["xlsx"])

if not uploaded_file:
    st.stop()

# 1) Leer archivo original
df = pd.read_excel(uploaded_file, engine="openpyxl")

st.subheader("Vista previa del archivo original")
st.dataframe(df.head(), use_container_width=True)

# 2) Normalizar columnas -> plantilla estándar
clean_df = normalize_columns(df)

# 2.b) Agregar consistencia a la plantilla
clean_df = add_consistency_to_clean(clean_df)

st.subheader("Hoja 1: PEI_normalizado + Consistencia por objetivo")
st.dataframe(clean_df.head(), use_container_width=True)


# 3) Construir hoja de RESUMEN (Año - Objetivo - Actividad)

def build_resumen(clean_df: pd.DataFrame) -> pd.DataFrame:
    filas = []

    for obj in range(1, 7):
        col_year = clean_df["Año"]
        col_obj = clean_df[f"Objetivo {obj}"]
        col_act = clean_df[f"Actividad Obj {obj}"]

        tmp = pd.DataFrame(
            {
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
            columns=["Año", "Objetivo específico", "Actividad relacionada"]
        )

    return resumen


resumen_df = build_resumen(clean_df)

# 3.b) Agregar consistencia a la hoja Resumen
resumen_df = add_consistency_to_resumen(resumen_df)

st.subheader("Hoja 2: Resumen + Índice de consistencia")
st.dataframe(resumen_df.head(), use_container_width=True)


# 4) Exportar las DOS hojas en un solo Excel

def to_excel_with_two_sheets(plantilla: pd.DataFrame, resumen: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        plantilla.to_excel(writer, index=False, sheet_name="PEI_normalizado")
        resumen.to_excel(writer, index=False, sheet_name="Resumen")
    return buffer.getvalue()


excel_bytes = to_excel_with_two_sheets(clean_df, resumen_df)

st.subheader("Descarga del archivo Excel")

st.download_button(
    label="📥 Descargar Excel con consistencia (2 hojas)",
    data=excel_bytes,
    file_name="pei_actividades_consistencia.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
