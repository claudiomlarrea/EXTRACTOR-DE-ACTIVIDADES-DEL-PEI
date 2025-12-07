import streamlit as st
import pandas as pd
from io import BytesIO
from normalizer import normalize_columns

st.set_page_config(page_title="Extractor PEI", page_icon="📘", layout="wide")

st.title("📘 Extractor y Normalizador de Actividades del PEI")

uploaded_file = st.file_uploader("Cargar archivo Excel original", type=["xlsx"])

if not uploaded_file:
    st.stop()

# 1) Leer archivo original
df = pd.read_excel(uploaded_file, engine="openpyxl")

st.subheader("Vista previa del archivo original")
st.dataframe(df.head(), use_container_width=True)

# 2) Normalizar columnas -> plantilla estándar
clean_df = normalize_columns(df)

st.subheader("Hoja 1: PEI_normalizado (plantilla estándar)")
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

        # Filtrar filas donde no haya actividad (ni objetivo) relevante
        mascara = tmp["Actividad relacionada"].astype(str).str.strip() != ""
        tmp = tmp[mascara]

        filas.append(tmp)

    if filas:
        resumen = pd.concat(filas, ignore_index=True)
    else:
        resumen = pd.DataFrame(
            columns=["Año", "Objetivo específico", "Actividad relacionada"]
        )

    return resumen


resumen_df = build_resumen(clean_df)

st.subheader("Hoja 2: Resumen (Año - Objetivo específico - Actividad relacionada)")
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
    label="📥 Descargar Excel con ambas hojas",
    data=excel_bytes,
    file_name="pei_actividades_filtradas_con_resumen.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
