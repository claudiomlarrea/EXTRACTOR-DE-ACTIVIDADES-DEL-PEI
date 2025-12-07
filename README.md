# Calculadora de Consistencia PEI (Objetivo ↔ Actividad)

Esta herramienta calcula un **índice de consistencia** entre el objetivo específico
elegido por una unidad académica y la **actividad única** que cargó en el formulario PEI.

La consistencia se expresa en la escala discreta:

> 0, 10, 30, 50, 70, 90, 100 (%)

## ¿Cómo se calcula?

1. Se construye un modelo TF-IDF con:
   - Los textos de todos los objetivos específicos.
   - El texto de cada actividad (actividad + detalle).

2. Para cada actividad, se calcula la **similitud coseno** con todos los objetivos.

3. Se compara:
   - La similitud con el objetivo que la unidad **declaró**.
   - La similitud con el mejor objetivo posible (Top-1).
   - La posición (rank) que ocupa el objetivo elegido (1º, 2º, 3º, etc.).

4. A partir de esos valores se aplica una regla determinista que asigna
   una nota de consistencia en la escala {0, 10, 30, 50, 70, 90, 100}.

## Uso

1. Crear y activar entorno virtual (opcional):

```bash
python -m venv .venv
source .venv/bin/activate           # en Windows: .venv\Scripts\activate
