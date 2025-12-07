import io
import datetime as dt

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches


# ---------------------------------------------------------
# Configuración básica de la app
# ---------------------------------------------------------
st.set_page_config(
    page_title="Consistencia PEI - Objetivos vs Actividades",
    layout="wide",
)

st.title("Análisis de consistencia entre Objetivos Específicos y Actividades Únicas")
st.write(
    "Esta herramienta analiza la consistencia entre los objetivos específicos del PEI "
    "y las actividades únicas cargadas por las unidades académicas."
)


# ---------------------------------------------------------
# Funciones auxiliares
# ---------------------------------------------------------
def detectar_columna_consistencia(df: pd.DataFrame) -> str | None:
    """
    Intenta encontrar la columna que contiene los valores de consistencia (%).
    Busca por patrones frecuentes en el nombre de la columna.
    """
    posibles = [
        "consistencia (%)",
        "consistencia%",
        "consistencia",
        "consistency",
        "consistency (%)",
    ]

    lower_cols = {c.lower(): c for c in df.columns}
    for patron in posibles:
        for col_lower, col_original in lower_cols.items():
            if patron in col_lower:
                return col_original
    return None


def detectar_columna_anio(df: pd.DataFrame) -> str | None:
    posibles = ["año", "anio", "ano", "year"]
    lower_cols = {c.lower(): c for c in df.columns}
    for patron in posibles:
        for col_lower, col_original in lower_cols.items():
            if patron == col_lower or patron in col_lower:
                return col_original
    return None


def detectar_columna_objetivo(df: pd.DataFrame) -> str | None:
    posibles = ["objetivo específico", "objetivo especifico", "objetivos específicos",
                "objetivos especificos", "objetivo", "objetivos"]
    lower_cols = {c.lower(): c for c in df.columns}
    for patron in posibles:
        for col_lower, col_original in lower_cols.items():
            if patron in col_lower:
                return col_original
    return None


def detectar_columna_actividad(df: pd.DataFrame) -> str | None:
    posibles = ["actividad", "actividad única", "actividad unica",
                "actividad obj", "actividad objetivo"]
    lower_cols = {c.lower(): c for c in df.columns}
    for patron in posibles:
        for col_lower, col_original in lower_cols.items():
            if patron in col_lower:
                return col_original
    return None


def categorizar_nivel_consistencia(valor: float) -> int:
    """
    Mapea un porcentaje de consistencia (0–100) a niveles discretos:
    0, 10, 30, 50, 70, 90, 100.
    """
    if pd.isna(valor):
        return 0

    if valor < 5:
        return 0
    elif valor < 20:
        return 10
    elif valor < 40:
        return 30
    elif valor < 60:
        return 50
    elif valor < 80:
        return 70
    elif valor < 95:
        return 90
    else:
        return 100


def generar_excel_para_descarga(df: pd.DataFrame) -> bytes:
    """
    Devuelve un archivo Excel en memoria a partir del DataFrame.
    """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Consistencia")
    buffer.seek(0)
    return buffer.getvalue()


def generar_informe_word(
    df: pd.DataFrame,
    col_consistencia: str,
    col_anio: str | None,
    promedio_global: float,
    distribucion_niveles: pd.Series,
) -> bytes:
    """
    Genera un informe en Word con:
    - Resumen numérico global
    - Distribución por niveles de consistencia
    - Interpretación
    - Conclusiones y recomendaciones
    Devuelve el archivo .docx como bytes.
    """
    doc = Document()

    # Portada / título
    doc.add_heading(
        "Informe de consistencia entre Objetivos Específicos y Actividades Únicas",
        level=1,
    )
    doc.add_paragraph(
        f"Fecha de generación del informe: {dt.datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )
    doc.add_paragraph(
        "Unidad responsable: Secretaría de Investigación / Observatorio de IA - UCCuyo"
    )
    doc.add_paragraph("")

    # Datos básicos
    total_actividades = len(df)
    doc.add_heading("1. Resumen general", level=2)
    p = doc.add_paragraph()
    p.add_run("Cantidad total de actividades únicas analizadas: ").bold = True
    p.add_run(f"{total_actividades}")

    p = doc.add_paragraph()
    p.add_run("Consistencia promedio global: ").bold = True
    p.add_run(f"{promedio_global:.2f} %")

    if col_anio is not None:
        anios = sorted(df[col_anio].dropna().unique())
        if len(anios) > 0:
            p = doc.add_paragraph()
            p.add_run("Años considerados en el análisis: ").bold = True
            p.add_run(", ".join(str(a) for a in anios))

    doc.add_paragraph("")

    # Tabla de distribución por niveles
    doc.add_heading("2. Distribución por niveles de consistencia", level=2)
    doc.add_paragraph(
        "La siguiente tabla muestra cuántas actividades se ubican en cada nivel de "
        "consistencia (0, 10, 30, 50, 70, 90, 100), donde 0 indica ausencia de "
        "alineación y 100 indica una coincidencia plena entre actividad y objetivo."
    )

    tabla = doc.add_table(rows=1 + len(distribucion_niveles), cols=2)
    hdr_cells = tabla.rows[0].cells
    hdr_cells[0].text = "Nivel de consistencia (%)"
    hdr_cells[1].text = "Cantidad de actividades"

    for i, (nivel, cantidad) in enumerate(distribucion_niveles.items(), start=1):
        row_cells = tabla.rows[i].cells
        row_cells[0].text = str(int(nivel))
        row_cells[1].text = str(int(cantidad))

    doc.add_paragraph("")

    # 3. Interpretación de resultados
    doc.add_heading("3. Interpretación de los resultados", level=2)

    if promedio_global < 20:
        nivel_texto = "muy bajo"
    elif promedio_global < 40:
        nivel_texto = "bajo"
    elif promedio_global < 60:
        nivel_texto = "medio"
    elif promedio_global < 80:
        nivel_texto = "aceptable/alto"
    else:
        nivel_texto = "muy alto"

    doc.add_paragraph(
        f"El índice de consistencia promedio obtenido es de {promedio_global:.2f} %, "
        f"lo que se interpreta como un nivel **{nivel_texto}** de concordancia entre "
        "las actividades reportadas por las unidades académicas y los objetivos "
        "específicos del Plan Estratégico Institucional (PEI)."
    )

    doc.add_paragraph(
        "La distribución por niveles permite identificar en qué tramo se concentra la "
        "mayor parte de las actividades. Una alta proporción en niveles de 0–10 % "
        "indica problemas de alineación o errores de clasificación de las acciones en "
        "los objetivos. En cambio, una mayor presencia en niveles de 70–100 % sugiere "
        "un uso más criterioso del PEI como marco orientador."
    )

    # 4. Conclusiones
    doc.add_heading("4. Conclusiones principales", level=2)
    doc.add_paragraph(
        "1. El valor promedio global sintetiza el grado de alineación efectiva entre "
        "la planificación estratégica y la ejecución reportada. Esto permite estimar "
        "en qué medida el PEI está siendo utilizado como guía real de la gestión."
    )
    doc.add_paragraph(
        "2. La presencia de actividades en niveles bajos de consistencia puede deberse "
        "a dos fenómenos: (a) acciones que efectivamente no responden al objetivo en "
        "el que fueron cargadas, o (b) objetivos mal seleccionados en el formulario "
        "de reporte."
    )
    doc.add_paragraph(
        "3. Los niveles altos de consistencia evidencian buenas prácticas de "
        "planificación y seguimiento, donde cada acción se vincula claramente con el "
        "resultado esperado del PEI."
    )

    # 5. Recomendaciones
    doc.add_heading("5. Recomendaciones para la gestión institucional", level=2)
    doc.add_paragraph(
        "• Devolver a cada unidad académica un resumen de su propio índice de "
        "consistencia, para fomentar la autoevaluación y el ajuste de futuras "
        "cargas de actividades."
    )
    doc.add_paragraph(
        "• Revisar las descripciones de los objetivos específicos en las "
        "comunicaciones operativas, de modo que sean más claras y fácilmente "
        "identificables por quienes completan los formularios."
    )
    doc.add_paragraph(
        "• Incorporar instancias de capacitación breves (microtalleres o cápsulas "
        "virtuales) sobre cómo vincular correctamente cada actividad con el objetivo "
        "correspondiente."
    )
    doc.add_paragraph(
        "• Utilizar este indicador de consistencia como una métrica periódica del "
        "Sistema de Aseguramiento de la Calidad y del seguimiento del PEI, "
        "integrándolo en los tableros de control (Power BI / Looker Studio)."
    )

    doc.add_paragraph("")
    doc.add_paragraph(
        "Este informe puede complementarse con análisis cualitativos de ejemplos de "
        "actividades con alta y baja consistencia, para retroalimentar las prácticas "
        "de gestión de cada unidad."
    )

    # Guardar a memoria
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------
# Carga de archivo
# ---------------------------------------------------------
st.sidebar.header("1. Subir archivo de consistencia")
uploaded_file = st.sidebar.file_uploader(
    "Suba el archivo Excel con la columna 'Consistencia (%)'",
    type=["xlsx", "xls"],
)

if uploaded_file is None:
    st.info("Subí un archivo Excel para comenzar el análisis.")
    st.stop()

# Leer el Excel
df = pd.read_excel(uploaded_file)

if df.empty:
    st.error("El archivo está vacío o no se pudo leer correctamente.")
    st.stop()

# Detectar columnas clave
col_consistencia = detectar_columna_consistencia(df)
col_anio = detectar_columna_anio(df)
col_obj = detectar_columna_objetivo(df)
col_act = detectar_columna_actividad(df)

if col_consistencia is None:
    st.error(
        "No se encontró ninguna columna de consistencia. "
        "Asegurate de que exista una columna llamada, por ejemplo, "
        "'Consistencia (%)'."
    )
    st.stop()

# Asegurar que los valores sean numéricos
df[col_consistencia] = pd.to_numeric(df[col_consistencia], errors="coerce")

# Crear columna de nivel discreto, si no existe
if "Nivel consistencia" not in df.columns:
    df["Nivel consistencia"] = df[col_consistencia].apply(categorizar_nivel_consistencia)

# ---------------------------------------------------------
# Cálculo de indicadores globales
# ---------------------------------------------------------
total_actividades = len(df)
promedio_global = df[col_consistencia].mean()

distribucion_niveles = (
    df["Nivel consistencia"]
    .value_counts()
    .sort_index()
)

st.subheader("Indicadores globales")

col1, col2 = st.columns(2)
with col1:
    st.metric("Cantidad total de actividades únicas", total_actividades)
with col2:
    st.metric("Consistencia promedio global (%)", f"{promedio_global:.2f}")

st.write("### Distribución de actividades por nivel de consistencia (%)")
st.dataframe(
    pd.DataFrame(
        {
            "Nivel de consistencia (%)": distribucion_niveles.index.astype(int),
            "Cantidad de actividades": distribucion_niveles.values.astype(int),
        }
    ),
    use_container_width=True,
)

# ---------------------------------------------------------
# Descarga de Excel procesado
# ---------------------------------------------------------
st.subheader("Descargar resultados")

excel_bytes = generar_excel_para_descarga(df)
st.download_button(
    label="📊 Descargar resultados en Excel",
    data=excel_bytes,
    file_name="consistencia_pei_resultados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ---------------------------------------------------------
# Generar y descargar informe en Word
# ---------------------------------------------------------
word_bytes = generar_informe_word(
    df=df,
    col_consistencia=col_consistencia,
    col_anio=col_anio,
    promedio_global=promedio_global,
    distribucion_niveles=distribucion_niveles,
)

st.download_button(
    label="📄 Descargar informe de consistencia en Word",
    data=word_bytes,
    file_name="informe_consistencia_pei.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
)
