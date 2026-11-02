# Gestor de notas para Streamlit

Aplicación para:
- Quitar de los nombres de archivo el prefijo `Student_YYYY_MM_DD__HH_MM_`, conservando el resto del nombre.
- Leer una plantilla Excel con el listado oficial del alumnado.
- Leer uno o varios archivos de calificaciones (Excel o CSV).
- Buscar coincidencias aproximadas entre nombres aunque haya diferencias de tildes, guiones o pequeñas variaciones.
- Convertir las puntuaciones a base 10.
- Escribir las notas en la columna indicada de la plantilla.
- Mostrar nombres no reconocidos, alumnos sin nota y posibles notas duplicadas.
- Descargar la plantilla completada y un ZIP con los archivos renombrados.

## Ejecutar en local

1. Instala Python 3.10 o superior.
2. Abre una terminal en esta carpeta.
3. Ejecuta:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicar en GitHub y Streamlit Community Cloud

1. Crea un repositorio en GitHub y sube `app.py`, `requirements.txt` y `README.md`.
2. En Streamlit Community Cloud, pulsa **Create app**.
3. Selecciona el repositorio, la rama y `app.py`.
4. Despliega la aplicación.

## Uso

1. Sube la plantilla `.xlsx`.
2. Sube todos los archivos con notas.
3. Selecciona la hoja de la plantilla.
4. Indica la columna de nombres, la columna donde escribir la nota y la primera fila del alumnado.
5. Selecciona las columnas de nombre y puntuación de los archivos de notas.
6. Indica la puntuación máxima del ejercicio (por ejemplo, 25 si el test está sobre 25).
7. Pulsa **Procesar notas** y revisa las coincidencias antes de descargar.

## Importante

- El prefijo se elimina mediante un patrón de fecha y hora, por lo que no depende de los números concretos.
- Si hay varias calificaciones para el mismo alumno, la aplicación conserva la última encontrada y lo señala para revisión.
- La coincidencia aproximada no es infalible. Revisa especialmente los nombres sin coincidencia y las similitudes bajas.
- La plantilla original no se modifica: se genera una copia rellenada.
- La versión inicial supone que los archivos de notas comparten las mismas columnas. Si tus exportaciones cambian de formato, se puede ampliar para configurar cada archivo por separado.
