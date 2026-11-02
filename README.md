# Herramientas de clase: renombrar archivos y pasar notas a iDoceo

Aplicación Streamlit con dos herramientas independientes.

## 1. Renombrar archivos individuales
Sube uno o varios archivos. Elimina el prefijo `Student_YYYY_MM_DD__HH_MM_` y conserva el resto del nombre, incluida la extensión. Descarga un ZIP con los archivos renombrados.

## 2. Pasar notas a iDoceo
- Admite exportaciones de iDoceo `.xls` y `.xlsx`.
- Admite uno o varios resultados de Socrative `.xlsx`, `.xls` y `.csv` simultáneamente.
- Detecta la fila de encabezados de Socrative aunque haya información previa.
- Utiliza `Nombre` y `Apellidos` de iDoceo para identificar al alumnado.
- Tolera tildes, diferencias pequeñas y nombres parcialmente distintos.
- Si Socrative incluye `Score (%)`, convierte directamente el porcentaje a nota sobre 10.
- Si se utiliza una puntuación bruta, permite indicar la puntuación máxima.
- Añade una nueva columna con el nombre que se indique, sin modificar la exportación original.
- Muestra coincidencias, alumnos sin nota y nombres no identificados antes de descargar.

## Ejecutar en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicar en GitHub y Streamlit Community Cloud

1. Crea un repositorio de GitHub y sube `app.py`, `requirements.txt` y `README.md`.
2. En https://share.streamlit.io/ crea una aplicación.
3. Selecciona el repositorio, la rama y `app.py`.

## Nota sobre el archivo de iDoceo
Si la exportación original es `.xls`, la aplicación genera el resultado como `.xlsx` para poder añadir la columna de notas. Comprueba que la versión de iDoceo que utilizas permite importar ese Excel. La hoja de datos personales se conserva, junto con sus columnas y el orden del alumnado.
