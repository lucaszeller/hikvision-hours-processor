from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
import csv
import io
import re

import pandas as pd

EXCEPTIONS_COLUMNS = ["ID de persona", "Fecha", "Tipo", "Detalle", "Manual"]


class ExceptionConfigError(Exception):
    pass


@dataclass(frozen=True)
class WorkException:
    employee_id: str | None
    exception_date: date
    exception_type: str
    details: str
    paid_day: bool = False


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


def _read_exceptions_dataframe(file_path: Path) -> pd.DataFrame:
    suffix = file_path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(file_path)
    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(file_path)
    raise ExceptionConfigError(
        "Formato de excepciones no soportado. Usa .csv, .xls o .xlsx."
    )


def _normalize_sheet_name(value: object) -> str:
    return _normalize(str(value))


def _find_sheet_name(file_path: Path, candidates: list[str]) -> str | None:
    if file_path.suffix.lower() not in {".xlsx", ".xls"}:
        return None
    try:
        excel = pd.ExcelFile(file_path)
    except Exception as exc:
        raise ExceptionConfigError(f"No se pudo leer archivo de excepciones '{file_path}': {exc}") from exc

    normalized_map = {_normalize_sheet_name(name): name for name in excel.sheet_names}
    for candidate in candidates:
        key = _normalize_sheet_name(candidate)
        if key in normalized_map:
            return normalized_map[key]
    return None


def _is_absences_template(file_path: Path) -> bool:
    try:
        return _find_sheet_name(file_path, ["Ausencias"]) is not None
    except ExceptionConfigError:
        return False


def _build_exceptions_from_absences_dataframe(df: pd.DataFrame) -> list[WorkException]:
    if df.empty:
        return []

    columns = [str(col) for col in df.columns]
    employee_col = _find_column(columns, ["legajo", "id", "id de persona"])
    type_col = _find_column(columns, ["tipo ausencia", "tipo", "type"])
    from_col = _find_column(columns, ["fecha desde", "desde", "fecha inicio", "inicio"])
    to_col = _find_column(columns, ["fecha hasta", "hasta", "fecha fin", "fin"])
    state_col = _find_column(columns, ["estado", "status"])
    details_col = _find_column(columns, ["observacion", "observación", "detalle", "nota"])

    if employee_col is None or type_col is None or from_col is None:
        raise ExceptionConfigError(
            "La hoja 'Ausencias' debe tener columnas de Legajo, Tipo ausencia y Fecha desde."
        )

    ignored_states = {"cancelado"}
    results: list[WorkException] = []
    for _, row in df.iterrows():
        raw_employee = _clean_optional(row.get(employee_col, ""))
        raw_type = _clean_optional(row.get(type_col, ""))
        raw_from = _clean_optional(row.get(from_col, ""))
        raw_to = _clean_optional(row.get(to_col, "")) if to_col else ""
        raw_state = _clean_optional(row.get(state_col, "")) if state_col else ""
        raw_details = _clean_optional(row.get(details_col, "")) if details_col else ""

        if raw_employee == "" and raw_type == "" and raw_from == "":
            continue
        if raw_employee == "":
            raise ExceptionConfigError("Fila en 'Ausencias' sin Legajo.")
        if raw_type == "":
            raise ExceptionConfigError("Fila en 'Ausencias' sin Tipo ausencia.")
        if raw_from == "":
            raise ExceptionConfigError("Fila en 'Ausencias' sin Fecha desde.")

        state_key = _normalize(raw_state)
        if state_key in ignored_states:
            continue

        paid_day = state_key == "aprobado"

        employee_target: str | None
        # Regla operativa: legajo "1" en template significa "todos los empleados".
        if raw_employee == "1":
            employee_target = None
        else:
            employee_target = raw_employee

        start_date = _parse_date(raw_from)
        end_date = _parse_date(raw_to) if raw_to else start_date
        if end_date < start_date:
            raise ExceptionConfigError(
                f"Rango invalido en 'Ausencias' para legajo {raw_employee}: Fecha hasta menor a Fecha desde."
            )

        exception_type = raw_type.strip().title()
        details = raw_details
        current = start_date
        while current <= end_date:
            results.append(
                WorkException(
                    employee_id=employee_target,
                    exception_date=current,
                    exception_type=exception_type,
                    details=details,
                    paid_day=paid_day,
                )
            )
            current += timedelta(days=1)

    return results


def load_absences_template_file(path: str | Path) -> list[WorkException]:
    file_path = Path(path)
    if not file_path.exists():
        raise ExceptionConfigError(f"No existe el archivo de ausencias: {file_path}")
    if file_path.suffix.lower() not in {".xlsx", ".xls"}:
        return []

    sheet_name = _find_sheet_name(file_path, ["Ausencias"])
    if sheet_name is None:
        return []

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as exc:
        raise ExceptionConfigError(
            f"No se pudo leer la hoja '{sheet_name}' en '{file_path}': {exc}"
        ) from exc

    return _build_exceptions_from_absences_dataframe(df)


def _write_exceptions_dataframe(file_path: Path, df: pd.DataFrame) -> None:
    suffix = file_path.suffix.lower()
    if suffix == ".csv":
        df.to_csv(file_path, index=False)
        return
    if suffix in {".xlsx", ".xls"}:
        df.to_excel(file_path, index=False)
        return
    raise ExceptionConfigError(
        "Formato de excepciones no soportado. Usa .csv, .xls o .xlsx."
    )


def _standardize_exceptions_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty and len(df.columns) == 0:
        return pd.DataFrame(columns=EXCEPTIONS_COLUMNS)

    columns = [str(col) for col in df.columns]
    employee_col = _find_column(columns, ["id de persona", "employee id", "legajo", "id"])
    date_col = _find_column(columns, ["fecha", "date", "work date"])
    type_col = _find_column(columns, ["tipo", "type", "excepcion", "motivo"])
    details_col = _find_column(columns, ["detalle", "details", "descripcion", "nota"])
    manual_col = _find_column(columns, ["manual", "carga manual", "origen manual"])

    if (date_col is None or type_col is None) and not df.empty:
        raise ExceptionConfigError(
            "El archivo de excepciones debe contener columnas de Fecha y Tipo."
        )

    rename_map: dict[str, str] = {}
    if employee_col and employee_col != "ID de persona":
        rename_map[employee_col] = "ID de persona"
    if date_col and date_col != "Fecha":
        rename_map[date_col] = "Fecha"
    if type_col and type_col != "Tipo":
        rename_map[type_col] = "Tipo"
    if details_col and details_col != "Detalle":
        rename_map[details_col] = "Detalle"
    if manual_col and manual_col != "Manual":
        rename_map[manual_col] = "Manual"

    result = df.rename(columns=rename_map).copy()
    for column in EXCEPTIONS_COLUMNS:
        if column not in result.columns:
            result[column] = ""

    return result


def _exception_key(employee_id: str | None, date_value: object, type_text: object, details: object) -> tuple[str, str, str, str]:
    employee_norm = _clean_optional(employee_id)
    date_norm = _parse_date(date_value).isoformat()
    type_norm = _clean_optional(type_text)
    details_norm = _clean_optional(details)
    return (employee_norm, date_norm, type_norm, details_norm)


def ensure_exceptions_file(path: str | Path) -> Path:
    file_path = Path(path)
    if not file_path.exists():
        file_path.parent.mkdir(parents=True, exist_ok=True)
        _write_exceptions_dataframe(file_path, pd.DataFrame(columns=EXCEPTIONS_COLUMNS))
        return file_path

    if _is_absences_template(file_path):
        # Plantilla operacional (Empleados/Ausencias): no normalizar su estructura.
        return file_path

    df = _read_exceptions_dataframe(file_path)
    standardized = _standardize_exceptions_dataframe(df)
    if list(df.columns) != list(standardized.columns):
        _write_exceptions_dataframe(file_path, standardized)
    return file_path


def load_exceptions_file(path: str | Path) -> list[WorkException]:
    file_path = Path(path)
    if not file_path.exists():
        raise ExceptionConfigError(f"No existe el archivo de excepciones: {file_path}")

    if _is_absences_template(file_path):
        return load_absences_template_file(file_path)

    df = _read_exceptions_dataframe(file_path)
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


def append_manual_exceptions_to_file(
    path: str | Path,
    manual_text: str | None,
) -> int:
    manual_exceptions = parse_manual_exceptions(manual_text)
    if not manual_exceptions:
        return 0

    file_path = ensure_exceptions_file(path)
    if _is_absences_template(file_path):
        raise ExceptionConfigError(
            "No se puede anexar carga manual al template de Ausencias. Usa un archivo de excepciones estandar."
        )
    df = _standardize_exceptions_dataframe(_read_exceptions_dataframe(file_path))

    existing_keys: set[tuple[str, str, str, str]] = set()
    for _, row in df.iterrows():
        row_date = _clean_optional(row.get("Fecha", ""))
        row_type = _clean_optional(row.get("Tipo", ""))
        if not row_date or not row_type:
            continue
        try:
            existing_keys.add(
                _exception_key(
                    row.get("ID de persona", ""),
                    row_date,
                    row_type,
                    row.get("Detalle", ""),
                )
            )
        except ExceptionConfigError:
            continue

    new_rows: list[dict[str, str]] = []
    for item in manual_exceptions:
        key = _exception_key(
            item.employee_id or "",
            item.exception_date.isoformat(),
            item.exception_type,
            item.details,
        )
        if key in existing_keys:
            continue
        existing_keys.add(key)
        new_rows.append(
            {
                "ID de persona": item.employee_id or "",
                "Fecha": item.exception_date.isoformat(),
                "Tipo": item.exception_type,
                "Detalle": item.details,
                "Manual": "Si",
            }
        )

    if not new_rows:
        return 0

    df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
    _write_exceptions_dataframe(file_path, df)
    return len(new_rows)


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
