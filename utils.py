from typing import List, Tuple
import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


# Stopwords muy básicas en español; podés ampliarlas si querés
STOPWORDS_ES = [
    "de", "la", "el", "los", "las", "y", "en", "del", "para", "con",
    "a", "por", "una", "un", "al", "que", "se", "su", "sus"
]


def build_tfidf_model(
    textos_objetivos: List[str],
    textos_actividades: List[str]
) -> Tuple[TfidfVectorizer, np.ndarray, np.ndarray]:
    """
    Construye un modelo TF-IDF conjunto para objetivos y actividades.
    Devuelve:
    - vectorizer
    - matriz TF-IDF de objetivos (X_obj)
    - matriz TF-IDF de actividades (X_act)
    """
    corpus = textos_objetivos + textos_actividades

    vectorizer = TfidfVectorizer(
        lowercase=True,
        ngram_range=(1, 2),          # unigrama + bigrama
        stop_words=STOPWORDS_ES,
        min_df=1
    )

    X = vectorizer.fit_transform(corpus)
    n_obj = len(textos_objetivos)

    X_obj = X[:n_obj, :]
    X_act = X[n_obj:, :]

    return vectorizer, X_obj, X_act


def score_consistency(sim_selected: float, sim_best: float, rank_selected: int) -> int:
    """
    Mapea la similitud del objetivo elegido vs. el mejor posible
    a la escala discreta {0, 10, 30, 50, 70, 90, 100}.
    """

    # Texto muy genérico: ninguna similitud relevante con ningún objetivo
    if sim_best < 0.10:
        # texto vago: no "castigamos" con 0, pero es poco informativo
        return 30

    ratio = sim_selected / (sim_best + 1e-6)

    # Caso 1: el objetivo elegido es el mejor (rank 1)
    if rank_selected == 1:
        # Match muy fuerte y nítido
        if sim_selected >= 0.40 and ratio >= 0.95:
            return 100
        # Muy buena coherencia
        if sim_selected >= 0.30 and ratio >= 0.90:
            return 90
        # Alta consistencia, pero algo menos clara
        if sim_selected >= 0.20 and ratio >= 0.80:
            return 70
        # Es el mejor, pero el texto es flojo / vago
        return 50

    # Caso 2: el objetivo elegido es 2º o 3º, pero bastante parecido al mejor
    if rank_selected in (2, 3) and ratio >= 0.70:
        return 30

    # Caso 3: cierta relación, pero floja
    if ratio >= 0.40:
        return 10

    # Caso 4: claramente inconsistente
    return 0


def compute_consistency_for_df(
    df: pd.DataFrame,
    col_obj_codigo: str,
    col_obj_texto: str,
    col_actividad: str,
    col_detalle: str
) -> pd.DataFrame:
    """
    Calcula la consistencia entre la ACTIVIDAD y el OBJETIVO elegido.

    df debe tener como mínimo estas columnas:
      - col_obj_codigo: código o identificador del objetivo elegido
      - col_obj_texto: texto completo del objetivo específico
      - col_actividad: actividad única
      - col_detalle: detalle de la actividad (opcional, puede estar vacío)

    Devuelve un nuevo DataFrame con:
      - Consistencia (%) ∈ {0,10,30,50,70,90,100}
      - Objetivo_más_parecido
      - Similitud_obj_elegido
      - Similitud_mejor_obj
      - Rank_obj_elegido
    """

    # Reset index para que coincida con la matriz de similitud
    df_work = df.reset_index(drop=True).copy()

    # Catálogo de objetivos únicos
    catalog = df_work[[col_obj_codigo, col_obj_texto]].drop_duplicates()
    codigos = catalog[col_obj_codigo].astype(str).tolist()
    textos_obj = catalog[col_obj_texto].astype(str).tolist()
    idx_by_codigo = {c: i for i, c in enumerate(codigos)}

    # Texto enriquecido de actividades: actividad + detalle
    textos_act = (
        df_work[col_actividad].fillna("").astype(str) + " " +
        df_work[col_detalle].fillna("").astype(str)
    ).tolist()

    # Modelo TF-IDF
    _, X_obj, X_act = build_tfidf_model(textos_obj, textos_act)

    # Matriz de similitud [n_act x n_obj]
    sim_matrix = cosine_similarity(X_act, X_obj)

    scores = []
    best_codes = []
    sim_selected_list = []
    sim_best_list = []
    rank_list = []

    n_rows = df_work.shape[0]

    for i in range(n_rows):
        codigo_elegido = str(df_work.at[i, col_obj_codigo])

        # Si el código elegido no está en el catálogo, no podemos evaluar bien
        if codigo_elegido not in idx_by_codigo:
            scores.append(0)
            best_codes.append(None)
            sim_selected_list.append(0.0)
            sim_best_list.append(0.0)
            rank_list.append(None)
            continue

        j_sel = idx_by_codigo[codigo_elegido]
        sims = sim_matrix[i, :]

        sim_sel = float(sims[j_sel])
        sim_best = float(sims.max())

        # ranking (1 = similitud más alta)
        order = (-sims).argsort()
        rank_sel = int(np.where(order == j_sel)[0][0]) + 1
        codigo_best = codigos[int(order[0])]

        score = score_consistency(sim_sel, sim_best, rank_sel)

        scores.append(score)
        best_codes.append(codigo_best)
        sim_selected_list.append(sim_sel)
        sim_best_list.append(sim_best)
        rank_list.append(rank_sel)

    df_work["Consistencia (%)"] = scores
    df_work["Objetivo_más_parecido"] = best_codes
    df_work["Similitud_obj_elegido"] = sim_selected_list
    df_work["Similitud_mejor_obj"] = sim_best_list
    df_work["Rank_obj_elegido"] = rank_list

    return df_work
