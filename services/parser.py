from __future__ import annotations

import html
import re
import unicodedata
from pathlib import Path

import pandas as pd


class ParsingError(Exception):
    pass


# Canonical keys used internally.
COLUMN_ALIASES: dict[str, list[str]] = {
    "employee_id": ["id de persona", "person id", "employee id", "legajo", "codigo"],
    "employee_name": ["nombre", "name", "employee name", "person name"],
    "department": ["departamento", "department", "dept"],
    "position": ["posicion", "position", "puesto"],
    "work_date": ["fecha", "date", "work date"],
    "schedule": ["horario", "schedule", "shift"],
    "entry_time": ["registro de entrada", "entrada", "check in", "in time"],
    "exit_time": ["registro de salida", "salida", "check out", "out time"],
    "worked": ["trabajo", "worked"],
    "overtime": ["horas extra", "overtime"],
    "attended": ["asistio", "attended"],
    "late": ["entrada con retraso", "late", "llegada tarde"],
    "left_early": ["temprano", "left early"],
    "absent": ["ausente", "absent"],
    "work_permission": ["permiso laboral", "work permission"],
}

REQUIRED_COLUMNS = ("employee_id", "employee_name", "work_date", "entry_time", "exit_time")


def _normalize(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    return " ".join(text.lower().replace("_", " ").split())


def _clean_cell(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "nat", "-"}:
        return ""
    return text


def _column_match_score(column: str, aliases: list[str]) -> int:
    normalized = _normalize(column)
    best = 0
    for alias in aliases:
        normalized_alias = _normalize(alias)
        if normalized == normalized_alias:
            return 100
        if normalized_alias in normalized:
            best = max(best, 80)
        if normalized in normalized_alias:
            best = max(best, 60)
    return best


def _best_column(columns: list[str], aliases: list[str]) -> str | None:
    ranked = sorted(columns, key=lambda c: _column_match_score(c, aliases), reverse=True)
    if ranked and _column_match_score(ranked[0], aliases) > 0:
        return ranked[0]
    return None


def detect_column_mapping(columns: list[str]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for canonical_name, aliases in COLUMN_ALIASES.items():
        match = _best_column(columns, aliases)
        if match:
            mapping[canonical_name] = match

    missing = [col for col in REQUIRED_COLUMNS if col not in mapping]
    if missing:
        raise ParsingError(
            "No se pudieron detectar columnas obligatorias: "
            + ", ".join(missing)
            + ". Revisa encabezados del reporte Hikvision."
        )

    return mapping


def _looks_like_hikvision_html_xls(path: Path) -> bool:
    header = path.read_bytes()[:512].lstrip().lower()
    return header.startswith(b"<html")


def _extract_hikvision_table(html_text: str, table_class: str) -> str:
    pattern = re.compile(
        rf"<table[^>]*class=['\"]{re.escape(table_class)}['\"][^>]*>(.*?)</table>",
        re.IGNORECASE | re.DOTALL,
    )
    match = pattern.search(html_text)
    if not match:
        raise ParsingError(f"No se encontro la tabla '{table_class}' en el reporte.")
    return match.group(1)


def _strip_html_cell(cell_html: str) -> str:
    text = re.sub(r"<[^>]+>", "", cell_html)
    text = html.unescape(text)
    return " ".join(text.strip().split())


def _parse_table_rows(table_html: str) -> list[list[str]]:
    rows: list[list[str]] = []
    current: list[str] = []
    token_pattern = re.compile(r"(?is)<td\b[^>]*>(.*?)</td>|</tr>")

    for match in token_pattern.finditer(table_html):
        cell_html = match.group(1)
        if cell_html is not None:
            current.append(_strip_html_cell(cell_html))
            continue

        if current:
            rows.append(current)
            current = []

    if current:
        rows.append(current)

    return rows


def _read_hikvision_html_report(path: Path) -> pd.DataFrame:
    raw_bytes = path.read_bytes()
    try:
        html_text = raw_bytes.decode("utf-8")
    except UnicodeDecodeError:
        html_text = raw_bytes.decode("latin-1")

    detail_html = _extract_hikvision_table(html_text, "Detail2")
    report_html = _extract_hikvision_table(html_text, "Daily_Report")

    header_rows = _parse_table_rows(detail_html)
    if not header_rows:
        raise ParsingError("No se pudo leer encabezado del reporte Hikvision.")

    headers = max(header_rows, key=len)
    if len(headers) < 5:
        raise ParsingError("Encabezado invalido en el reporte Hikvision.")

    data_rows = _parse_table_rows(report_html)
    normalized_rows: list[list[str]] = []
    for row in data_rows:
        if not any(cell.strip() for cell in row):
            continue
        if len(row) < len(headers):
            row = row + [""] * (len(headers) - len(row))
        normalized_rows.append(row[: len(headers)])

    if not normalized_rows:
        raise ParsingError("La tabla Daily_Report no contiene filas de datos.")

    return pd.DataFrame(normalized_rows, columns=headers)


def _read_hikvision_xlsx(path: Path) -> pd.DataFrame:
    # Hikvision files can shift header row by a few lines.
    for header in (0, 1, 2, 3):
        df = pd.read_excel(path, header=header)
        if df.empty:
            continue
        columns = [str(c) for c in df.columns]
        try:
            detect_column_mapping(columns)
            return df
        except ParsingError:
            continue

    raise ParsingError(
        "No se pudo detectar una fila de encabezados compatible en el Excel (filas 1 a 4)."
    )


def _read_source_dataframe(path: Path) -> pd.DataFrame:
    if _looks_like_hikvision_html_xls(path):
        return _read_hikvision_html_report(path)
    return _read_hikvision_xlsx(path)


def load_hikvision_excel(file_path: str | Path) -> pd.DataFrame:
    path = Path(file_path)
    if not path.exists():
        raise ParsingError(f"No existe el archivo: {path}")

    source_df = _read_source_dataframe(path)
    if source_df.empty:
        raise ParsingError("El archivo no contiene filas.")

    source_columns = [str(c) for c in source_df.columns]
    mapping = detect_column_mapping(source_columns)

    canonical_df = pd.DataFrame()
    canonical_df["employee_id"] = source_df[mapping["employee_id"]].map(_clean_cell)
    canonical_df["employee_name"] = source_df[mapping["employee_name"]].map(_clean_cell)
    canonical_df["department"] = (
        source_df[mapping["department"]].map(_clean_cell) if "department" in mapping else ""
    )
    canonical_df["schedule"] = (
        source_df[mapping["schedule"]].map(_clean_cell) if "schedule" in mapping else ""
    )
    canonical_df["work_date_raw"] = source_df[mapping["work_date"]].map(_clean_cell)
    canonical_df["entry_time_raw"] = source_df[mapping["entry_time"]].map(_clean_cell)
    canonical_df["exit_time_raw"] = source_df[mapping["exit_time"]].map(_clean_cell)

    # Keep source row index for traceability in inconsistencies.
    canonical_df["source_row"] = source_df.index.to_series().astype(int) + 2

    # Remove lines that are completely empty in critical fields.
    critical = ["employee_id", "employee_name", "work_date_raw", "entry_time_raw", "exit_time_raw"]
    mask_non_empty = canonical_df[critical].apply(lambda col: col.astype(str).str.strip() != "").any(axis=1)
    canonical_df = canonical_df[mask_non_empty].reset_index(drop=True)

    if canonical_df.empty:
        raise ParsingError("No se detectaron filas utiles en el archivo de entrada.")

    return canonical_df
