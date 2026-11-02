import io
import re
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Herramientas de clase", page_icon="📚", layout="wide")
st.title("Herramientas de clase")

# ---------------------------------------------------------------------
# Utilidades compartidas
# ---------------------------------------------------------------------
PREFIX_RE = re.compile(
    r"^Student_\d{4}_\d{2}_\d{2}__\d{2}_\d{2}_",
    re.IGNORECASE
)

def clean_filename(filename: str) -> str:
    """Quita Student_fecha__hora_ sin depender de los números concretos."""
    return PREFIX_RE.sub("", filename, count=1)

def normalize_name(value) -> str:
    if pd.isna(value):
        return ""
    value = str(value).lower().strip()
    value = re.sub(r"[^a-záéíóúüñ\s]", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()

def name_score(a, b):
    """Combina coincidencia de palabras y coincidencia parcial."""
    a, b = normalize_name(a), normalize_name(b)
    if not a or not b:
        return 0
    return max(
        fuzz.token_set_ratio(a, b),
        fuzz.WRatio(a, b),
        fuzz.partial_ratio(a, b) * 0.92
    )

def read_excel_any(uploaded, sheet_name=0, header=0):
    raw = uploaded.getvalue()
    suffix = Path(uploaded.name).suffix.lower()
    if suffix == ".csv":
        try:
            return pd.read_csv(io.BytesIO(raw), encoding="utf-8-sig", header=header)
        except UnicodeDecodeError:
            return pd.read_csv(io.BytesIO(raw), encoding="latin1", header=header)
    if suffix in (".xlsx", ".xlsm"):
        return pd.read_excel(io.BytesIO(raw), sheet_name=sheet_name, header=header)
    if suffix == ".xls":
        return pd.read_excel(io.BytesIO(raw), sheet_name=sheet_name, header=header, engine="xlrd")
    raise ValueError("Formato no compatible. Utiliza XLS, XLSX o CSV.")

def read_idoceo_as_workbook(uploaded):
    """Carga exportación iDoceo XLS/XLSX y la convierte a un libro XLSX editable."""
    raw = uploaded.getvalue()
    suffix = Path(uploaded.name).suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        return load_workbook(io.BytesIO(raw))
    if suffix == ".xls":
        # Se lee con pandas/xlrd y se reconstruye el libro para descargarlo en XLSX.
        xls = pd.ExcelFile(io.BytesIO(raw), engine="xlrd")
        wb = load_workbook(io.BytesIO(_empty_xlsx_bytes()))
        # Reemplazar hoja inicial por cada hoja original, conservando datos y columnas.
        first = True
        for sheet in xls.sheet_names:
            df = pd.read_excel(io.BytesIO(raw), sheet_name=sheet, header=None, engine="xlrd")
            ws = wb.active if first else wb.create_sheet()
            first = False
            ws.title = str(sheet)[:31]
            for row in df.itertuples(index=False, name=None):
                ws.append([None if pd.isna(v) else v for v in row])
        return wb
    raise ValueError("La plantilla de iDoceo debe ser XLS o XLSX.")

def _empty_xlsx_bytes():
    from openpyxl import Workbook
    b = io.BytesIO()
    Workbook().save(b)
    b.seek(0)
    return b.getvalue()

def to_float(value):
    if pd.isna(value):
        return None
    if isinstance(value, str):
        value = value.strip().replace(",", ".").replace("%", "")
        value = re.sub(r"[^0-9.\-]", "", value)
    try:
        return float(value)
    except (TypeError, ValueError):
        return None

def best_match(source_name, roster_names):
    scores = [(name, name_score(source_name, name)) for name in roster_names]
    scores.sort(key=lambda x: x[1], reverse=True)
    if not scores:
        return None, 0, 0
    second = scores[1][1] if len(scores) > 1 else 0
    return scores[0][0], scores[0][1], second

# ---------------------------------------------------------------------
# Pestañas independientes
# ---------------------------------------------------------------------
tab_rename, tab_grades = st.tabs(["1 · Renombrar archivos", "2 · Pasar notas a iDoceo"])

with tab_rename:
    st.header("Renombrar archivos individuales")
    st.write(
        "Sube los archivos exportados. Se eliminará únicamente el prefijo "
        "`Student_fecha__hora_`; el resto del nombre y la extensión se conservarán."
    )
    files_to_rename = st.file_uploader(
        "Selecciona uno o varios archivos",
        accept_multiple_files=True,
        key="rename_files"
    )
    if files_to_rename:
        preview = pd.DataFrame([
            {"Nombre original": f.name, "Nombre nuevo": clean_filename(f.name)}
            for f in files_to_rename
        ])
        st.dataframe(preview, use_container_width=True, hide_index=True)
        zip_out = io.BytesIO()
        with zipfile.ZipFile(zip_out, "w", zipfile.ZIP_DEFLATED) as archive:
            used = set()
            for f in files_to_rename:
                new_name = clean_filename(f.name)
                # Evita sobrescribir si dos archivos acaban con el mismo nombre.
                candidate = new_name
                stem, suffix = Path(new_name).stem, Path(new_name).suffix
                counter = 2
                while candidate.lower() in used:
                    candidate = f"{stem}_{counter}{suffix}"
                    counter += 1
                used.add(candidate.lower())
                archive.writestr(candidate, f.getvalue())
        st.download_button(
            "Descargar archivos renombrados (ZIP)",
            data=zip_out.getvalue(),
            file_name="archivos_renombrados.zip",
            mime="application/zip",
            use_container_width=True
        )
    else:
        st.info("Esta herramienta funciona de forma independiente: no necesitas cargar una plantilla ni notas.")

with tab_grades:
    st.header("Pasar calificaciones a una exportación de iDoceo")
    st.write(
        "Carga la exportación de la clase de iDoceo y el Excel de resultados "
        "de Socrative. Se añadirá una columna con la nota sobre 10, respetando "
        "el orden del alumnado de iDoceo."
    )

    col_a, col_b = st.columns(2)
    with col_a:
        idoceo_file = st.file_uploader(
            "1. Plantilla/exportación de iDoceo (.xls o .xlsx)",
            type=["xls", "xlsx", "xlsm"],
            key="idoceo_file"
        )
    with col_b:
        socrative_file = st.file_uploader(
            "2. Resultados de Socrative (.xlsx, .xls o .csv)",
            type=["xlsx", "xls", "csv"],
            key="socrative_file"
        )

    if idoceo_file and socrative_file:
        try:
            wb = read_idoceo_as_workbook(idoceo_file)
            sheet_name = st.selectbox("Hoja de iDoceo", wb.sheetnames, key="idoceo_sheet")
            ws = wb[sheet_name]

            st.markdown("#### Configuración de iDoceo")
            h1, h2, h3 = st.columns(3)
            with h1:
                header_row = st.number_input("Fila de encabezados", min_value=1, value=1, step=1, key="idoceo_header")
            with h2:
                first_row = st.number_input("Primera fila de alumnos", min_value=2, value=2, step=1, key="idoceo_first")
            with h3:
                new_column = st.text_input("Nombre de la nueva columna", value="Dicc lista 1", key="idoceo_newcol")

            headers = [ws.cell(int(header_row), c).value for c in range(1, ws.max_column + 1)]
            header_labels = [f"{get_column_letter(i+1)} · {str(v) if v is not None else '(vacía)'}" for i, v in enumerate(headers)]
            default_first = next((i for i, v in enumerate(headers) if str(v).strip().lower() == "nombre"), 0)
            default_last = next((i for i, v in enumerate(headers) if str(v).strip().lower() == "apellidos"), min(1, len(headers)-1))
            p1, p2 = st.columns(2)
            with p1:
                first_name_col = st.selectbox("Columna Nombre", range(len(headers)), index=default_first, format_func=lambda i: header_labels[i])
            with p2:
                surname_col = st.selectbox("Columna Apellidos", range(len(headers)), index=default_last, format_func=lambda i: header_labels[i])

            roster = []
            row_map = {}
            for r in range(int(first_row), ws.max_row + 1):
                first = ws.cell(r, first_name_col + 1).value
                last = ws.cell(r, surname_col + 1).value
                if first is None and last is None:
                    continue
                full = " ".join(str(x).strip() for x in (first, last) if x is not None and str(x).strip())
                if full:
                    roster.append(full)
                    row_map[full] = r

            st.caption(f"Alumnado localizado en iDoceo: {len(roster)}")

            # Leer el archivo de Socrative. Sus exportaciones suelen llevar metadatos
            # arriba y la fila Student Name / Score (%) como encabezado.
            raw_socrative = socrative_file.getvalue()
            suffix = Path(socrative_file.name).suffix.lower()
            if suffix in (".xlsx", ".xlsm"):
                raw_preview = pd.read_excel(io.BytesIO(raw_socrative), header=None)
            elif suffix == ".xls":
                raw_preview = pd.read_excel(io.BytesIO(raw_socrative), header=None, engine="xlrd")
            else:
                raw_preview = pd.read_csv(io.BytesIO(raw_socrative), header=None, encoding="utf-8-sig")

            header_idx = None
            for i in range(min(len(raw_preview), 30)):
                vals = [str(v).strip().lower() for v in raw_preview.iloc[i].tolist()]
                if any(v in ("student name", "student", "nombre", "nombre del alumno") for v in vals):
                    header_idx = i
                    break
            if header_idx is None:
                header_idx = 0

            if suffix in (".xlsx", ".xlsm"):
                source_df = pd.read_excel(io.BytesIO(raw_socrative), header=header_idx)
            elif suffix == ".xls":
                source_df = pd.read_excel(io.BytesIO(raw_socrative), header=header_idx, engine="xlrd")
            else:
                source_df = pd.read_csv(io.BytesIO(raw_socrative), header=header_idx, encoding="utf-8-sig")

            source_cols = list(source_df.columns)
            name_default = next((i for i, c in enumerate(source_cols) if str(c).strip().lower() in ("student name", "student", "nombre", "nombre del alumno")), 0)
            score_default = next((i for i, c in enumerate(source_cols) if "score (%)" in str(c).lower() or "porcentaje" in str(c).lower()), None)
            score_mode_default = score_default is not None

            st.markdown("#### Configuración de Socrative")
            s1, s2, s3 = st.columns(3)
            with s1:
                source_name_col = st.selectbox("Columna con el nombre del alumno", source_cols, index=name_default, key="soc_name")
            with s2:
                source_score_col = st.selectbox("Columna con la puntuación", source_cols, index=score_default if score_default is not None else min(2, len(source_cols)-1), key="soc_score")
            with s3:
                use_percent = st.checkbox("La puntuación está expresada en porcentaje", value=score_mode_default, key="soc_percent")

            max_score = 10.0
            if not use_percent:
                max_score = st.number_input("Puntuación máxima del ejercicio", min_value=0.01, value=19.0, step=1.0, key="soc_max")
            threshold = st.slider("Umbral mínimo de coincidencia de nombres", 50, 100, 78, key="match_threshold")
            st.caption("El programa compara nombre y apellidos sin distinguir mayúsculas, tildes ni pequeños cambios. Las coincidencias dudosas se muestran para que puedas revisarlas.")

            if st.button("Preparar notas para iDoceo", type="primary", use_container_width=True, key="process_grades"):
                score_by_student = {}
                audit = []
                unmatched = []
                ambiguous = []

                for _, record in source_df.iterrows():
                    source_name = record.get(source_name_col)
                    raw_score = to_float(record.get(source_score_col))
                    if pd.isna(source_name) or raw_score is None:
                        continue

                    if use_percent:
                        grade10 = raw_score / 10.0
                    else:
                        grade10 = raw_score / float(max_score) * 10.0
                    grade10 = round(max(0, min(10, grade10)), 2)

                    matched, top_score, second_score = best_match(source_name, roster)
                    if matched and top_score >= threshold:
                        score_by_student[matched] = grade10
                        audit.append({
                            "Nombre en Socrative": str(source_name),
                            "Alumno de iDoceo": matched,
                            "Nota /10": grade10,
                            "Coincidencia": round(top_score, 1)
                        })
                        if top_score - second_score < 5:
                            ambiguous.append(str(source_name) + " → " + matched)
                    else:
                        unmatched.append({
                            "Nombre en Socrative": str(source_name),
                            "Nota /10": grade10,
                            "Mejor coincidencia": matched or "",
                            "Puntuación de coincidencia": round(top_score, 1)
                        })

                output_col = ws.max_column + 1
                ws.cell(int(header_row), output_col).value = new_column.strip() or "Nota"
                ws.cell(int(header_row), output_col).font = Font(bold=True)
                ws.cell(int(header_row), output_col).fill = PatternFill("solid", fgColor="D9EAF7")

                for student, row_num in row_map.items():
                    if student in score_by_student:
                        ws.cell(row_num, output_col).value = score_by_student[student]
                        ws.cell(row_num, output_col).number_format = "0.00"

                out = io.BytesIO()
                wb.save(out)
                out.seek(0)

                st.session_state["grade_output"] = out.getvalue()
                st.session_state["grade_audit"] = pd.DataFrame(audit)
                st.session_state["grade_unmatched"] = pd.DataFrame(unmatched)
                st.session_state["grade_missing"] = [s for s in roster if s not in score_by_student]
                st.session_state["grade_ambiguous"] = ambiguous
                st.session_state["grade_count"] = len(score_by_student)
                st.session_state["grade_colname"] = new_column.strip() or "Nota"

        except Exception as exc:
            st.error(f"No se han podido leer los archivos: {exc}")
            st.info("Si el archivo XLS de iDoceo no se abre, vuelve a exportarlo desde iDoceo o utiliza su versión XLSX.")

    if "grade_output" in st.session_state:
        st.divider()
        st.subheader("Revisión antes de descargar")
        m1, m2, m3 = st.columns(3)
        m1.metric("Alumnos con nota", st.session_state["grade_count"])
        m2.metric("Sin coincidencia", len(st.session_state["grade_unmatched"]))
        m3.metric("Alumnos sin nota", len(st.session_state["grade_missing"]))

        if st.session_state["grade_ambiguous"]:
            st.warning("Hay coincidencias con nombres parecidos. Revísalas: " + "; ".join(st.session_state["grade_ambiguous"]))
        if not st.session_state["grade_unmatched"].empty:
            st.markdown("**Nombres de Socrative que no se han podido asignar:**")
            st.dataframe(st.session_state["grade_unmatched"], use_container_width=True, hide_index=True)
        if st.session_state["grade_missing"]:
            st.markdown("**Alumnos de iDoceo que no han recibido nota:**")
            st.write(", ".join(st.session_state["grade_missing"]))
        if not st.session_state["grade_audit"].empty:
            st.markdown("**Cruce de nombres y notas:**")
            st.dataframe(st.session_state["grade_audit"], use_container_width=True, hide_index=True)

        st.download_button(
            "Descargar archivo de iDoceo con las notas",
            data=st.session_state["grade_output"],
            file_name="iDoceo_con_notas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.info("Este apartado es independiente del renombrado. Aquí solo necesitas la exportación de iDoceo y el archivo de resultados.")
