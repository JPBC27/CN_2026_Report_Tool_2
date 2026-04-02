# Consolidación y Reporte CN 2026

Este proyecto es una herramienta de automatización para la consolidación de datos de RR.HH. (Dotación y Censo Perú), desarrollada en Python utilizando Streamlit y Pandas.

## 🚀 Cómo Subir a GitHub

Sigue estos pasos para subir este proyecto a tu propia cuenta de GitHub:

1.  **Crea un repositorio en GitHub**:
    *   Inicia sesión en [GitHub](https://github.com/).
    *   Haz clic en el botón "+" y selecciona "New repository".
    *   Dale un nombre (ej: `CN-2026-Report-Tool`) y haz clic en "Create repository".
    *   **NO** selecciones la opción de inicializar con un README o .gitignore (ya los hemos creado aquí).

2.  **Conecta tu repositorio local con GitHub**:
    Abre una terminal en la carpeta `CN_2026_Report_Tool` y ejecuta:
    ```bash
    git remote add origin https://github.com/TU_USUARIO/TU_REPOSITORIO.git
    git branch -M main
    git push -u origin main
    ```

## 🛠️ Requisitos e Instalación

1.  Asegúrate de tener Python 3.8 o superior.
2.  Instala las dependencias:
    ```bash
    pip install -r requirements.txt
    ```
3.  Ejecuta la plataforma:
    ```bash
    streamlit run app.py
    ```

## 📁 Estructura del Proyecto

- `app.py`: Interfaz de usuario (Streamlit).
- `reporte_cn_2026.py`: Lógica de procesamiento de datos (Pandas).
- `requirements.txt`: Lista de librerías necesarias.
- `.gitignore`: Archivos que Git debe ignorar.
