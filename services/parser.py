from __future__ import annotations

import html
import re
import unicodedata
from pathlib import Path

import pandas as pd

from domain.column_aliases import COLUMN_CANDIDATES

REQUIRED_COLUMNS = ("employee_id", "employee_name", "punch_datetime")


class ParsingError(Exception):
    pass


def _normalize(value: str) -> str:
    text = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    return " ".join(text.lower().replace("_", " ").split())


def _looks_like_hikvision_html_xls(path: Path) -> bool:
    header = path.read_bytes()[:512].lstrip().lower()
    return header.startswith(b"<html")


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


def _marker_is_positive(value: object) -> bool:
    text = str(value).strip()
    if text in {"", "-", "0", "0.0", "0:0", "0:00", "nan", "NaN", "None"}:
        return False

    if ":" in text:
        parts = [p.strip() for p in text.split(":")]
        try:
            numbers = [int(part) for part in parts if part != ""]
            return any(number > 0 for number in numbers)
        except ValueError:
            return True

    try:
        return float(text.replace(",", ".")) > 0
    except ValueError:
        return True


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


def _build_canonical_from_entry_exit(df: pd.DataFrame) -> pd.DataFrame:
    columns = [str(c) for c in df.columns]
    entry_exit_aliases = {
        "employee_id": ["id de persona", "employee id", "person id", "legajo", "codigo"],
        "employee_name": ["nombre", "employee name", "name", "person name"],
        "work_date": ["fecha", "date", "work date"],
        "entry_time": ["registro de entrada", "entrada", "check in", "in time"],
        "exit_time": ["registro de salida", "salida", "check out", "out time"],
        "late_marker": ["entrada con retraso", "late", "llegada tarde", "retardo"],
        "absent_marker": ["ausente", "absent", "falta", "faltas"],
    }

    mapping: dict[str, str] = {}
    for canonical_name, aliases in entry_exit_aliases.items():
        ranked = sorted(columns, key=lambda c: _column_match_score(c, aliases), reverse=True)
        if ranked and _column_match_score(ranked[0], aliases) > 0:
            mapping[canonical_name] = ranked[0]

    required = ("employee_id", "employee_name", "work_date", "entry_time", "exit_time")
    missing = [key for key in required if key not in mapping]
    if missing:
        raise ParsingError(
            "No se pudieron detectar columnas de entrada/salida: " + ", ".join(missing)
        )

    base = pd.DataFrame()
    base["employee_id"] = df[mapping["employee_id"]].astype(str).str.strip()
    base["employee_name"] = df[mapping["employee_name"]].astype(str).str.strip()
    base["work_date_raw"] = pd.to_datetime(df[mapping["work_date"]], errors="coerce")
    base["late_flag"] = (
        df[mapping["late_marker"]].map(_marker_is_positive) if "late_marker" in mapping else False
    )
    base["absent_flag"] = (
        df[mapping["absent_marker"]].map(_marker_is_positive)
        if "absent_marker" in mapping
        else False
    )
    base["work_date"] = base["work_date_raw"].dt.date

    valid_base = base.dropna(subset=["work_date_raw"]).copy()
    if valid_base.empty:
        raise ParsingError("No hay fechas validas en el reporte.")

    date_text = valid_base["work_date_raw"].dt.strftime("%Y-%m-%d")
    entry_text = (
        df.loc[valid_base.index, mapping["entry_time"]]
        .astype(str)
        .str.strip()
        .replace({"-": "", "nan": "", "NaN": "", "None": ""})
    )
    exit_text = (
        df.loc[valid_base.index, mapping["exit_time"]]
        .astype(str)
        .str.strip()
        .replace({"-": "", "nan": "", "NaN": "", "None": ""})
    )

    entry_dt = pd.to_datetime(date_text + " " + entry_text, errors="coerce")
    exit_dt = pd.to_datetime(date_text + " " + exit_text, errors="coerce")

    punch_rows: list[pd.DataFrame] = []

    entry_rows = valid_base.copy()
    entry_rows["punch_datetime"] = entry_dt
    entry_rows = entry_rows.dropna(subset=["punch_datetime"])
    if not entry_rows.empty:
        punch_rows.append(entry_rows)

    exit_rows = valid_base.copy()
    exit_rows["punch_datetime"] = exit_dt
    exit_rows = exit_rows.dropna(subset=["punch_datetime"])
    if not exit_rows.empty:
        punch_rows.append(exit_rows)

    no_punch_status_rows = valid_base[
        entry_dt.isna()
        & exit_dt.isna()
        & (valid_base["late_flag"].astype(bool) | valid_base["absent_flag"].astype(bool))
    ].copy()
    if not no_punch_status_rows.empty:
        no_punch_status_rows["punch_datetime"] = pd.NaT
        punch_rows.append(no_punch_status_rows)

    if not punch_rows:
        raise ParsingError("No se pudieron construir fichadas validas desde entrada/salida.")

    canonical_df = pd.concat(punch_rows, ignore_index=True)
    canonical_df = canonical_df[
        (canonical_df["employee_id"] != "") & (canonical_df["employee_name"] != "")
    ]
    canonical_df = canonical_df[
        [
            "employee_id",
            "employee_name",
            "work_date",
            "punch_datetime",
            "late_flag",
            "absent_flag",
        ]
    ]
    canonical_df = canonical_df.sort_values(
        by=["employee_id", "employee_name", "work_date", "punch_datetime"],
        na_position="last",
    ).reset_index(drop=True)
    return canonical_df


def detect_column_mapping(columns: list[str]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for canonical_name, aliases in COLUMN_CANDIDATES.items():
        ranked = sorted(columns, key=lambda c: _column_match_score(c, aliases), reverse=True)
        if ranked and _column_match_score(ranked[0], aliases) > 0:
            mapping[canonical_name] = ranked[0]

    missing = [col for col in REQUIRED_COLUMNS if col not in mapping]
    if missing:
        raise ParsingError(
            "No se pudieron detectar columnas obligatorias: "
            + ", ".join(missing)
            + ". Revisa encabezados del Excel de Hikvision."
        )

    return mapping


def load_hikvision_excel(file_path: str | Path) -> pd.DataFrame:
    path = Path(file_path)
    if not path.exists():
        raise ParsingError(f"No existe el archivo: {path}")

    if _looks_like_hikvision_html_xls(path):
        source_df = _read_hikvision_html_report(path)
        canonical_df = _build_canonical_from_entry_exit(source_df)
    else:
        source_df = pd.read_excel(path)
        if source_df.empty:
            raise ParsingError("El archivo no contiene filas.")

        mapping = detect_column_mapping([str(c) for c in source_df.columns])

        canonical_df = pd.DataFrame()
        canonical_df["employee_id"] = source_df[mapping["employee_id"]].astype(str).str.strip()
        canonical_df["employee_name"] = source_df[mapping["employee_name"]].astype(str).str.strip()
        canonical_df["punch_datetime"] = pd.to_datetime(
            source_df[mapping["punch_datetime"]], errors="coerce"
        )
        canonical_df["event_type"] = (
            source_df[mapping["event_type"]].astype(str).str.strip()
            if "event_type" in mapping
            else None
        )
        canonical_df["late_flag"] = False
        canonical_df["absent_flag"] = False

        canonical_df = canonical_df.dropna(subset=["punch_datetime"])
        canonical_df = canonical_df[
            (canonical_df["employee_id"] != "") & (canonical_df["employee_name"] != "")
        ]

    if canonical_df.empty:
        raise ParsingError("No quedaron fichadas validas luego de limpiar datos incompletos.")

    if "work_date" not in canonical_df.columns:
        canonical_df["work_date"] = canonical_df["punch_datetime"].dt.date
    else:
        canonical_df["work_date"] = canonical_df["work_date"].fillna(
            canonical_df["punch_datetime"].dt.date
        )
    canonical_df["work_time"] = canonical_df["punch_datetime"].dt.time
    return canonical_df
