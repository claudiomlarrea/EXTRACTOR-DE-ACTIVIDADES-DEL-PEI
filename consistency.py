import pandas as pd


def _score_match(activity: str, keywords):
    """Devuelve 1 / 0.5 / 0 según cuántas palabras clave aparecen en la actividad."""
    text = str(activity).lower()
    matches = sum(1 for k in keywords if k in text)

    if matches >= 2:
        return 1.0
    elif matches == 1:
        return 0.5
    else:
        return 0.0


def evaluate_consistency(objective: str, activity: str) -> float:
    """
    Calcula un índice de consistencia 0–100 entre
    el texto del objetivo específico y la actividad.
    """

    obj = str(objective).lower()
    act = str(activity).lower()

    # 1) Grupos temáticos generales del PEI (ajustables)
    thematic_groups = {
        "calidad": ["monitoreo", "seguimiento", "evaluación", "mejora", "indicador"],
        "convenios": ["convenio", "articulación", "alianza", "acuerdo"],
        "docencia": ["capacitación", "formación", "docente", "curso", "clase"],
        "investigación": ["proyecto", "investigación", "paper", "publicación"],
        "extensión": ["extensión", "comunidad", "social", "vinculación"],
        "gestión": ["reunión", "planificación", "comisión", "organización", "gestión"],
    }

    # 1.a) Detectar categoría del objetivo
    category = "otros"
    for cat, words in thematic_groups.items():
        if any(w in obj for w in words):
            category = cat
            break

    # 2) Coincidencia semántica básica (40%)
    base_keywords = [w for w in obj.split() if len(w) > 3]
    semantic = _score_match(act, base_keywords)

    # 3) Pertinencia temática (40%)
    thematic_keywords = thematic_groups.get(category, [])
    thematic = _score_match(act, thematic_keywords)

    # 4) Pertinencia operativa (20%)
    action_verbs = [
        "realizar", "implementar", "evaluar", "crear", "organizar",
        "desarrollar", "monitorear", "diseñar", "consolidar", "actualizar"
    ]
    is_action = any(v in act for v in action_verbs)
    if is_action:
        operational = 1.0
    elif len(act.strip()) > 3:
        operational = 0.5
    else:
        operational = 0.0

    IC = semantic * 0.40 + thematic * 0.40 + operational * 0.20
    return round(IC * 100, 1)


def add_consistency_to_clean(clean_df: pd.DataFrame) -> pd.DataFrame:
    """
    Agrega, para cada objetivo 1–6, columnas:
      - Consistencia Obj n (%)
      - Nivel Obj n  (Baja/Media/Alta)
    """
    df = clean_df.copy()

    for obj in range(1, 7):
        col_obj = f"Objetivo {obj}"
        col_act = f"Actividad Obj {obj}"
        cons_col = f"Consistencia Obj {obj} (%)"
        level_col = f"Nivel Obj {obj}"

        if col_obj in df.columns and col_act in df.columns:
            df[cons_col] = df.apply(
                lambda row: evaluate_consistency(row[col_obj], row[col_act])
                if str(row[col_act]).strip() != "" else 0.0,
                axis=1,
            )

            df[level_col] = df[cons_col].apply(
                lambda x: "Alta" if x >= 71 else "Media" if x >= 31 else "Baja"
            )
        else:
            # Si faltaran columnas, las creamos vacías para no romper estructura
            df[cons_col] = 0.0
            df[level_col] = ""

    return df


def add_consistency_to_resumen(resumen_df: pd.DataFrame) -> pd.DataFrame:
    """
    Agrega a la hoja Resumen:
      - Consistencia (%)
      - Nivel de consistencia
    Asume columnas:
      - 'Objetivo específico'
      - 'Actividad relacionada'
    """
    df = resumen_df.copy()

    df["Consistencia (%)"] = df.apply(
        lambda row: evaluate_consistency(
            row["Objetivo específico"], row["Actividad relacionada"]
        ),
        axis=1,
    )

    df["Nivel de consistencia"] = df["Consistencia (%)"].apply(
        lambda x: "Alta" if x >= 71 else "Media" if x >= 31 else "Baja"
    )

    return df
