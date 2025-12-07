import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Extractor de columnas PEI",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Extractor de columnas del PEI")
st.write(
    """
    Cargá el archivo Excel exportado desde Looker Studio o desde el Formulario Único
    y seleccioná solo las columnas que querés conservar.  
    Luego podrás descargar un archivo nuevo con esas columnas.
    """
)

# 1) Subir archivo
uploaded_file = st.file_uploader(
    "📁 Cargar archivo Excel (.xlsx)",
    type=["xlsx"],
    help="Usá el archivo descargado desde Looker Studio o desde el Formulario Único para el PEI.",
)

if uploaded_file is None:
    st.info("Subí un archivo Excel para comenzar.")
    st.stop()

# 2) Leer archivo
try:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
except Exception as e:
    st.error(f"❌ No se pudo leer el archivo: {e}")
    st.stop()

st.success(f"Archivo cargado correctamente. Filas: {len(df)}, Columnas: {len(df.columns)}")

with st.expander("👀 Ver primeras filas del archivo original"):
    st.dataframe(df.head(), use_container_width=True)

# 3) Selección de columnas
st.subheader("✔ Selección de columnas a extraer")

all_columns = list(df.columns)

selected_columns = st.multiselect(
    "Elegí las columnas que querés conservar en el nuevo archivo:",
    options=all_columns,
    default=all_columns,  # podés reducir después
)

if not selected_columns:
    st.warning("Seleccioná al menos una columna.")
    st.stop()

df_filtered = df[selected_columns]

st.write(f"El archivo filtrado tendrá **{len(df_filtered.columns)} columnas** y **{len(df_filtered)} filas**.")

with st.expander("👀 Ver vista previa del archivo filtrado"):
    st.dataframe(df_filtered.head(), use_container_width=True)

# 4) Funciones auxiliares para descarga
def to_excel_bytes(dataframe: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Actividades_filtradas")
    return buffer.getvalue()

def to_csv_bytes(dataframe: pd.DataFrame) -> bytes:
    return dataframe.to_csv(index=False).encode("utf-8-sig")


# 5) Botones de descarga
st.subheader("⬇ Descargar archivo filtrado")

col1, col2 = st.columns(2)

with col1:
    excel_bytes = to_excel_bytes(df_filtered)
    st.download_button(
        label="📥 Descargar Excel (.xlsx)",
        data=excel_bytes,
        file_name="pei_actividades_filtradas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with col2:
    csv_bytes = to_csv_bytes(df_filtered)
    st.download_button(
        label="📥 Descargar CSV (.csv)",
        data=csv_bytes,
        file_name="pei_actividades_filtradas.csv",
        mime="text/csv",
    )

st.success("Listo. Podés subir otro archivo o cambiar la selección de columnas cuando quieras.")
