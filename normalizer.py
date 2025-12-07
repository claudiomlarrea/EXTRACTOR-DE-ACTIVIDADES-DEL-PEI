import pandas as pd


def _safe_first_match(df: pd.DataFrame, regex: str, default):
    """
    Busca la primera columna que matchee el regex.
    Si no encuentra ninguna, devuelve la serie 'default'.
    """
    cols = df.filter(regex=regex, axis=1)

    if cols.shape[1] > 0:
        return cols.iloc[:, 0]
    else:
        # default puede ser una Serie o un escalar
        if isinstance(default, pd.Series):
            return default
        else:
            return pd.Series([default] * len(df), index=df.index)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Genera SIEMPRE una plantilla con:
    ID, Año,
    Objetivo 1, Actividad Obj 1, Detalle Obj 1,
    ...
    Objetivo 6, Actividad Obj 6, Detalle Obj 6
    """

    n = len(df)
    index = df.index

    # Valores por defecto
    default_id = pd.Series(range(1, n + 1), index=index)      # 1..n si no hay ID
    default_year = pd.Series([""] * n, index=index)           # vacío si no hay Año
    default_blank = pd.Series([""] * n, index=index)          # para objetivos/actividades/detalles

    clean = pd.DataFrame(index=index)

    # --- ID ---
    clean["ID"] = _safe_first_match(
        df,
        regex=r"(?i)^id$|^id | id$| id ",   # más tolerante
        default=default_id,
    )

    # --- Año ---
    clean["Año"] = _safe_first_match(
        df,
        regex=r"(?i)año|anio|year",
        default=default_year,
    )

    # --- Objetivos 1 a 6 ---
    for obj in range(1, 7):
        # Objetivo específico
        objetivo = _safe_first_match(
            df,
            regex=fr"(?i)objetiv[oa].*{obj}",   # cualquier cosa que contenga 'objetivo' y el número
            default=default_blank,
        )
        actividad = _safe_first_match(
            df,
            regex=fr"(?i)activ.*{obj}",         # 'Actividades Objetivo 1', etc.
            default=default_blank,
        )
        detalle = _safe_first_match(
            df,
            regex=fr"(?i)detall.*{obj}",        # 'Detalle de la Actividad Objetivo 1', etc.
            default=default_blank,
        )

        clean[f"Objetivo {obj}"] = objetivo
        clean[f"Actividad Obj {obj}"] = actividad
        clean[f"Detalle Obj {obj}"] = detalle

    return clean
