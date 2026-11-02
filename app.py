import io
import re
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Gestor de notas", page_icon="📚", layout="wide")

st.title("Gestor de notas")
st.caption("Renombra exportaciones, convierte calificaciones a base 10 y las coloca en la plantilla del alumnado.")

PREFIX_RE = re.compile(r"^Student_\d{4}_\d{2}_\d{2}__\d{2}_\d{2}_", re.IGNORECASE)

def clean_filename(name: str) -> str:
    """Elimina el prefijo Student_fecha__hora_ sin depender de números concretos."""
    return PREFIX_RE.sub("", name, count=1)

def norm_name(value) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip().lower()
    s = re.sub(r"\.(xlsx|xls|csv)$", "", s)
    s = re.sub(r"[_\-]+", " ", s)
    s = re.sub(r"[^a-záéíóúüñ\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def read_table(uploaded, sheet_name=None):
    data = uploaded.getvalue()
    suffix = Path(uploaded.name).suffix.lower()
    if suffix == ".csv":
        for enc in ("utf-8-sig", "latin1"):
            try:
                return pd.read_csv(io.BytesIO(data), encoding=enc)
            except Exception:
                pass
        raise ValueError(f"No se ha podido leer {uploaded.name} como CSV.")
    if suffix in (".xlsx", ".xlsm"):
        return pd.read_excel(io.BytesIO(data), sheet_name=sheet_name or 0)
    if suffix == ".xls":
        return pd.read_excel(io.BytesIO(data), sheet_name=sheet_name or 0)
    raise ValueError(f"Formato no compatible: {suffix}")

def best_column(columns, hints):
    for c in columns:
        if any(h in str(c).lower() for h in hints):
            return c
    return columns[0] if len(columns) else None

def numeric_grade(value):
    if pd.isna(value):
        return None
    if isinstance(value, str):
        value = value.strip().replace("%", "").replace(",", ".")
        value = re.sub(r"[^0-9.\-]", "", value)
    try:
        return float(value)
    except (TypeError, ValueError):
        return None

def match_student(name, student_names, threshold):
    target = norm_name(name)
    if not target:
        return None, 0
    normalized = {n: norm_name(n) for n in student_names}
    choices = list(normalized.values())
    hit = process.extractOne(target, choices, scorer=fuzz.WRatio)
    if not hit:
        return None, 0
    _, score, idx = hit
    original = student_names[idx]
    return (original if score >= threshold else None), score

st.subheader("1. Carga tus archivos")
left, right = st.columns(2)
with left:
    template_file = st.file_uploader("Plantilla con nombres y huecos de notas (.xlsx)", type=["xlsx", "xlsm"])
with right:
    grade_files = st.file_uploader("Archivos con las notas (.xlsx, .xls o .csv)", type=["xlsx", "xls", "xlsm", "csv"], accept_multiple_files=True)

if template_file and grade_files:
    try:
        template_bytes = template_file.getvalue()
        wb = load_workbook(io.BytesIO(template_bytes))
        sheet_names = wb.sheetnames
        selected_sheet = st.selectbox("Hoja de la plantilla", sheet_names)
        ws = wb[selected_sheet]

        # Preview template as dataframe to select columns/rows.
        preview = pd.read_excel(io.BytesIO(template_bytes), sheet_name=selected_sheet, header=None)
        st.markdown("### 2. Indica dónde están los nombres y las notas en la plantilla")
        c1, c2, c3 = st.columns(3)
        with c1:
            header_row = st.number_input("Fila de encabezados (0 si no hay)", min_value=0, value=0, step=1)
        with c2:
            name_col_idx = st.number_input("Columna de nombres (A=0, B=1…)", min_value=0, value=0, step=1)
        with c3:
            grade_col_idx = st.number_input("Columna donde escribir la nota (A=0, B=1…)", min_value=0, value=1, step=1)
        first_student_row = st.number_input("Primera fila del alumnado en Excel (empieza en 1)", min_value=1, value=int(header_row)+2, step=1)

        students = []
        excel_rows = []
        for r in range(int(first_student_row), ws.max_row + 1):
            val = ws.cell(r, int(name_col_idx)+1).value
            if val is not None and str(val).strip():
                students.append(str(val).strip())
                excel_rows.append(r)

        st.write(f"Alumnado detectado en la plantilla: **{len(students)}**")
        if not students:
            st.error("No se han encontrado nombres en esa columna. Comprueba la hoja y la columna seleccionadas.")
            st.stop()

        st.markdown("### 3. Configura cómo leer las notas")
        st.caption("Se asume que todos los archivos de notas tienen las mismas columnas. Puedes ajustar los selectores tras cargar una muestra.")
        sample = read_table(grade_files[0])
        if sample.empty:
            st.error("El primer archivo de notas está vacío.")
            st.stop()
        cols = list(sample.columns)
        g1, g2, g3 = st.columns(3)
        with g1:
            source_name_col = st.selectbox("Columna del nombre en los archivos de notas", cols, index=cols.index(best_column(cols, ["student", "name", "nombre", "alumno"])) if best_column(cols, ["student", "name", "nombre", "alumno"]) in cols else 0)
        with g2:
            source_grade_col = st.selectbox("Columna de la puntuación obtenida", cols, index=cols.index(best_column(cols, ["score", "grade", "nota", "puntuación", "points"])) if best_column(cols, ["score", "grade", "nota", "puntuación", "points"]) in cols else min(1, len(cols)-1))
        with g3:
            max_score = st.number_input("Puntuación máxima del ejercicio", min_value=0.01, value=10.0, step=0.5, help="Por ejemplo, escribe 25 si el resultado está sobre 25. La aplicación lo convertirá a base 10.")

        threshold = st.slider("Exigencia para considerar que dos nombres coinciden", min_value=50, max_value=100, value=82, step=1)
        st.caption("La coincidencia aproximada ayuda con tildes, guiones, segundos apellidos o pequeñas diferencias. Revisa siempre los casos dudosos antes de descargar.")

        if st.button("Procesar notas", type="primary", use_container_width=True):
            collected = []
            rename_map = []
            errors = []
            for up in grade_files:
                cleaned = clean_filename(up.name)
                rename_map.append((up.name, cleaned))
                try:
                    df = read_table(up)
                    if source_name_col not in df.columns or source_grade_col not in df.columns:
                        errors.append(f"{up.name}: no contiene las columnas elegidas.")
                        continue
                    for _, row in df.iterrows():
                        nm = row[source_name_col]
                        raw = numeric_grade(row[source_grade_col])
                        if raw is None:
                            continue
                        matched, similarity = match_student(nm, students, threshold)
                        if matched:
                            score10 = raw / float(max_score) * 10
                            collected.append({"Alumno plantilla": matched, "Nombre archivo": str(nm), "Nota original": raw, "Nota /10": round(score10, 2), "Coincidencia": round(similarity)})
                        else:
                            collected.append({"Alumno plantilla": None, "Nombre archivo": str(nm), "Nota original": raw, "Nota /10": round(raw / float(max_score) * 10, 2), "Coincidencia": round(similarity)})
                except Exception as e:
                    errors.append(f"{up.name}: {e}")

            results = pd.DataFrame(collected)
            if results.empty:
                st.error("No se han podido extraer calificaciones. Revisa las columnas seleccionadas.")
                st.stop()

            # If an individual appears more than once, keep the latest encountered result and flag it.
            matched_rows = results[results["Alumno plantilla"].notna()].copy()
            duplicates = matched_rows[matched_rows.duplicated("Alumno plantilla", keep=False)]["Alumno plantilla"].unique().tolist()
            final_scores = matched_rows.drop_duplicates("Alumno plantilla", keep="last").set_index("Alumno plantilla")["Nota /10"].to_dict()

            for student, row_num in zip(students, excel_rows):
                if student in final_scores:
                    ws.cell(row_num, int(grade_col_idx)+1).value = final_scores[student]

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            st.session_state["output_xlsx"] = output.getvalue()
            st.session_state["results"] = results
            st.session_state["rename_map"] = rename_map
            st.session_state["errors"] = errors
            st.session_state["duplicates"] = duplicates
            st.session_state["missing"] = [s for s in students if s not in final_scores]

    except Exception as e:
        st.error(f"No se han podido preparar los archivos: {e}")

if "output_xlsx" in st.session_state:
    st.divider()
    st.subheader("4. Revisión de resultados")
    results = st.session_state["results"]
    unmatched = results[results["Alumno plantilla"].isna()]
    a, b, c = st.columns(3)
    a.metric("Notas procesadas", len(results))
    b.metric("Nombres sin coincidencia", len(unmatched))
    c.metric("Alumnado sin nota", len(st.session_state["missing"]))

    if st.session_state["duplicates"]:
        st.warning("Hay alumnos con más de una nota. Se ha colocado la última encontrada; comprueba si es la correcta: " + ", ".join(st.session_state["duplicates"]))
    if st.session_state["errors"]:
        st.warning("Algunos archivos no se han podido procesar:\n\n" + "\n".join(st.session_state["errors"]))
    if not unmatched.empty:
        st.markdown("**Nombres de los archivos que no han encontrado coincidencia:**")
        st.dataframe(unmatched[["Nombre archivo", "Nota original", "Nota /10", "Coincidencia"]], use_container_width=True, hide_index=True)
    if st.session_state["missing"]:
        st.markdown("**Alumnos de la plantilla que se han quedado sin nota:**")
        st.write(", ".join(st.session_state["missing"]))

    st.dataframe(results, use_container_width=True, hide_index=True)
    st.download_button(
        "Descargar plantilla rellenada",
        data=st.session_state["output_xlsx"],
        file_name="plantilla_notas_rellenada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    # A separate ZIP containing renamed source files (original contents preserved).
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for up in grade_files:
            zf.writestr(clean_filename(up.name), up.getvalue())
    zip_buffer.seek(0)
    st.download_button(
        "Descargar archivos con nombres corregidos (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="archivos_renombrados.zip",
        mime="application/zip",
        use_container_width=True
    )
else:
    st.info("Sube la plantilla y al menos un archivo de notas para empezar.")
