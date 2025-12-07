import pandas as pd

def normalize_columns(df: pd.DataFrame):
    """
    Detecta columnas por patrón y genera una plantilla con:
    ID, Año, Obj1, Act1, Det1, Obj2, Act2, Det2, … Obj6, Act6, Det6
    """

    clean = pd.DataFrame()

    # ID y Año
    clean["ID"] = df.filter(regex="(?i)^id$|id ", axis=1).iloc[:, 0]
    clean["Año"] = df.filter(regex="(?i)año|anio|year", axis=1).iloc[:, 0]

    for obj in range(1, 7):
        # Objetivo
        col_obj = df.filter(regex=fr"(?i)objetiv[oa].*{obj}", axis=1)
        clean[f"Objetivo {obj}"] = col_obj.iloc[:, 0] if not col_obj.empty else ""

        # Actividad
        col_act = df.filter(regex=fr"(?i)activ.*{obj}", axis=1)
        clean[f"Actividad Obj {obj}"] = col_act.iloc[:, 0] if not col_act.empty else ""

        # Detalle
        col_det = df.filter(regex=fr"(?i)detall.*{obj}", axis=1)
        clean[f"Detalle Obj {obj}"] = col_det.iloc[:, 0] if not col_det.empty else ""

    return clean
