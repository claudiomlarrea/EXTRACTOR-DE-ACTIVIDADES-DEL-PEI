# Extractor de columnas del PEI 📊

Pequeña aplicación en **Streamlit** para trabajar con los archivos Excel
descargados desde **Looker Studio** o desde el **Formulario Único para el PEI**.

Permite:
- Subir el Excel original.
- Elegir qué columnas conservar.
- Previsualizar el resultado.
- Descargar un nuevo Excel o CSV solo con las columnas seleccionadas.

---

## Requisitos

- Python 3.9 o superior.
- `pip` instalado.

## Instalación

```bash
git clone https://github.com/TU_USUARIO/pei-column-extractor.git
cd pei-column-extractor
python -m venv .venv
source .venv/bin/activate   # En Windows: .venv\Scripts\activate
pip install -r requirements.txt

