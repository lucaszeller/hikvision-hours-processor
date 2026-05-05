from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import csv
import io
import re

import pandas as pd


class ExceptionConfigError(Exception):
    pass


@dataclass(frozen=True)
class WorkException:
    employee_id: str | None
    exception_date: date
    exception_type: str
    details: str


def _normalize(text: object) -> str:
    return " ".join(str(text).strip().lower().replace("_", " ").split())


def _find_column(columns: list[str], aliases: list[str]) -> str | None:
    normalized_columns = {_normalize(col): col for col in columns}

    for alias in aliases:
        alias_normalized = _normalize(alias)
        if alias_normalized in normalized_columns:
            return normalized_columns[alias_normalized]

    for col in columns:
        col_norm = _normalize(col)
        if any(_normalize(alias) in col_norm for alias in aliases):
            return col

    return None


def _parse_date(value: object) -> date:
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        raise ExceptionConfigError(f"Fecha invalida en excepciones: '{value}'")
    return parsed.date()


def _clean_optional(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"", "nan", "none", "nat", "-"}:
        return ""
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def _build_exceptions_from_dataframe(df: pd.DataFrame) -> list[WorkException]:
    if df.empty:
        return []

    columns = [str(col) for col in df.columns]

    employee_col = _find_column(columns, ["id de persona", "employee id", "legajo", "id"])
    date_col = _find_column(columns, ["fecha", "date", "work date"])
    type_col = _find_column(columns, ["tipo", "type", "excepcion", "motivo"])
    details_col = _find_column(columns, ["detalle", "details", "descripcion", "nota"])

    if date_col is None or type_col is None:
        raise ExceptionConfigError(
            "El archivo de excepciones debe contener columnas de Fecha y Tipo."
        )

    results: list[WorkException] = []

    for _, row in df.iterrows():
        raw_employee = _clean_optional(row[employee_col]) if employee_col else ""
        raw_date = _clean_optional(row[date_col])
        raw_type = _clean_optional(row[type_col])
        raw_details = _clean_optional(row[details_col]) if details_col else ""

        if raw_date == "" and raw_type == "" and raw_employee == "" and raw_details == "":
            continue

        if raw_date == "":
            raise ExceptionConfigError("Fila de excepcion sin Fecha.")
        if raw_type == "":
            raise ExceptionConfigError("Fila de excepcion sin Tipo.")

        results.append(
            WorkException(
                employee_id=raw_employee or None,
                exception_date=_parse_date(raw_date),
                exception_type=raw_type,
                details=raw_details,
            )
        )

    return results


def load_exceptions_file(path: str | Path) -> list[WorkException]:
    file_path = Path(path)
    if not file_path.exists():
        raise ExceptionConfigError(f"No existe el archivo de excepciones: {file_path}")

    suffix = file_path.suffix.lower()
    if suffix == ".csv":
        df = pd.read_csv(file_path)
    elif suffix in {".xlsx", ".xls"}:
        df = pd.read_excel(file_path)
    else:
        raise ExceptionConfigError(
            "Formato de excepciones no soportado. Usa .csv, .xls o .xlsx."
        )

    return _build_exceptions_from_dataframe(df)


def _split_manual_line(line: str) -> list[str]:
    text = line.strip()
    if "|" in text:
        parts = [part.strip() for part in text.split("|", 3)]
    elif ";" in text:
        parts = [part.strip() for part in text.split(";", 3)]
    else:
        reader = csv.reader(io.StringIO(text))
        parts = [part.strip() for part in next(reader)]

    while len(parts) < 4:
        parts.append("")

    return parts[:4]


def parse_manual_exceptions(text: str | None) -> list[WorkException]:
    if text is None:
        return []

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return []

    results: list[WorkException] = []

    for index, line in enumerate(lines, start=1):
        if line.startswith("#"):
            continue

        employee_id, date_text, type_text, details = _split_manual_line(line)

        if index == 1 and _normalize(date_text) in {"fecha", "date"}:
            continue

        if date_text == "" or type_text == "":
            raise ExceptionConfigError(
                f"Linea manual invalida ({index}). Formato: ID|YYYY-MM-DD|TIPO|DETALLE"
            )

        if not re.match(r"^\d{4}-\d{2}-\d{2}$", date_text):
            raise ExceptionConfigError(
                f"Fecha manual invalida en linea {index}: '{date_text}'. Usa YYYY-MM-DD."
            )

        results.append(
            WorkException(
                employee_id=employee_id or None,
                exception_date=_parse_date(date_text),
                exception_type=type_text,
                details=details,
            )
        )

    return results


def merge_exceptions(
    file_exceptions: list[WorkException] | None,
    manual_exceptions: list[WorkException] | None,
) -> list[WorkException]:
    merged = list(file_exceptions or []) + list(manual_exceptions or [])

    # Preserve order while removing exact duplicates.
    deduped: list[WorkException] = []
    seen: set[tuple[str | None, date, str, str]] = set()
    for item in merged:
        key = (item.employee_id, item.exception_date, item.exception_type.strip(), item.details.strip())
        if key in seen:
            continue
        seen.add(key)
        deduped.append(item)

    return deduped


def build_exception_index(exceptions: list[WorkException]) -> dict[tuple[str | None, date], list[WorkException]]:
    index: dict[tuple[str | None, date], list[WorkException]] = {}
    for item in exceptions:
        key = (item.employee_id, item.exception_date)
        index.setdefault(key, []).append(item)
    return index


def find_matching_exceptions(
    index: dict[tuple[str | None, date], list[WorkException]],
    employee_id: str,
    work_date: date,
) -> list[WorkException]:
    employee_key = (employee_id or None, work_date)
    global_key = (None, work_date)

    return list(index.get(employee_key, [])) + list(index.get(global_key, []))
