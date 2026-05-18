from __future__ import annotations

import unicodedata
from pathlib import Path
from typing import Any
from datetime import time

import pandas as pd


class ScheduleInfoError(Exception):
    pass


def _normalize(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    return " ".join(text.lower().replace("_", " ").split())


def _clean_id(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"", "nan", "none", "nat", "-"}:
        return ""
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def _is_yes(value: object) -> bool:
    text = _normalize(value)
    return text in {"si", "s", "yes", "y", "true", "1"}


def _is_flexible_attendance(name_value: object, group_value: object) -> bool:
    name_key = _normalize(name_value)
    group_key = _normalize(group_value)
    if name_key == "juarez paulina":
        return True
    return any(token in group_key for token in {"maestranza", "limpieza"})


def _to_time(value: object) -> pd.Timestamp | None:
    if pd.isna(value):
        return None
    if isinstance(value, time):
        return pd.Timestamp.combine(pd.Timestamp.today().date(), value)
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed


def _minutes_between(start: pd.Timestamp | None, end: pd.Timestamp | None) -> int:
    if start is None or end is None:
        return 0
    start_minutes = start.hour * 60 + start.minute
    end_minutes = end.hour * 60 + end.minute
    diff = end_minutes - start_minutes
    return diff if diff > 0 else 0


def _to_minute_of_day(value: pd.Timestamp | None) -> int | None:
    if value is None:
        return None
    return value.hour * 60 + value.minute


def _round_to_30(minutes: int) -> int:
    if minutes <= 0:
        return 0
    quotient, remainder = divmod(minutes, 30)
    if remainder >= 15:
        quotient += 1
    return quotient * 30


def _parse_working_weekdays(value: object) -> set[int] | None:
    text = _normalize(value)
    if text in {"", "nan", "none", "nat", "-"}:
        return None

    day_map = {
        "lunes": 0,
        "martes": 1,
        "miercoles": 2,
        "jueves": 3,
        "viernes": 4,
        "sabado": 5,
        "domingo": 6,
    }

    # Rango habitual: "lunes a viernes"
    range_match = [
        ("lunes a viernes", {0, 1, 2, 3, 4}),
        ("lunes-viernes", {0, 1, 2, 3, 4}),
        ("lunes a sabado", {0, 1, 2, 3, 4, 5}),
        ("lunes-sabado", {0, 1, 2, 3, 4, 5}),
    ]
    for marker, days in range_match:
        if marker in text:
            return set(days)

    normalized = (
        text.replace(",", " ")
        .replace(";", " ")
        .replace("/", " ")
        .replace("-", " ")
        .replace(" y ", " ")
        .replace(" e ", " ")
    )
    tokens = [token.strip() for token in normalized.split() if token.strip()]
    found: set[int] = set()
    for token in tokens:
        if token in day_map:
            found.add(day_map[token])
    return found if found else None


def _best_column(columns: list[str], aliases: list[str]) -> str | None:
    normalized = {_normalize(col): col for col in columns}
    for alias in aliases:
        key = _normalize(alias)
        if key in normalized:
            return normalized[key]

    for col in columns:
        col_key = _normalize(col)
        if any(_normalize(alias) in col_key for alias in aliases):
            return col
    return None


def load_schedule_profiles(info_path: str | Path) -> dict[str, dict[str, Any]]:
    path = Path(info_path)
    if not path.exists():
        return {}

    try:
        with pd.ExcelFile(path) as excel:
            normalized_sheets = {_normalize(name): name for name in excel.sheet_names}
        preferred_sheet = normalized_sheets.get(_normalize("Empleados"))
        if preferred_sheet is not None:
            df = pd.read_excel(path, sheet_name=preferred_sheet)
        else:
            df = pd.read_excel(path)
    except Exception as exc:
        raise ScheduleInfoError(f"No se pudo leer archivo de horarios '{path}': {exc}") from exc

    if df.empty:
        return {}

    columns = [str(c) for c in df.columns]
    col_id = _best_column(columns, ["id", "legajo", "id de persona"])
    col_m_in = _best_column(columns, ["horario ingreso manana", "ingreso manana"])
    col_m_out = _best_column(columns, ["horario salida manana", "salida manana"])
    col_t_in = _best_column(columns, ["horario ingreso tarde", "ingreso tarde"])
    col_t_out = _best_column(columns, ["horario salida tarde", "salida tarde"])
    col_corrido = _best_column(columns, ["horario corrido", "corrido"])
    col_days = _best_column(columns, ["dias", "dias laborales", "dias de trabajo"])
    col_name = _best_column(columns, ["nombre", "name", "empleado"])
    col_group = _best_column(columns, ["grupo", "sector", "area"])

    if col_id is None:
        raise ScheduleInfoError("No se encontro columna de ID en el archivo de horarios.")

    result: dict[str, dict[str, Any]] = {}
    for _, row in df.iterrows():
        employee_id = _clean_id(row[col_id])
        if not employee_id:
            continue

        morning_in = _to_time(row[col_m_in]) if col_m_in else None
        morning_out = _to_time(row[col_m_out]) if col_m_out else None
        afternoon_in = _to_time(row[col_t_in]) if col_t_in else None
        afternoon_out = _to_time(row[col_t_out]) if col_t_out else None
        is_continuous = _is_yes(row[col_corrido]) if col_corrido else False
        is_flexible = _is_flexible_attendance(
            row[col_name] if col_name else "",
            row[col_group] if col_group else "",
        )

        if is_continuous:
            start = morning_in or afternoon_in
            end = afternoon_out or morning_out
            scheduled = _minutes_between(start, end)
        else:
            scheduled = _minutes_between(morning_in, morning_out) + _minutes_between(
                afternoon_in, afternoon_out
            )

        start_minute = _to_minute_of_day(morning_in or afternoon_in)
        working_weekdays = _parse_working_weekdays(row[col_days]) if col_days else None
        if working_weekdays is None:
            working_weekdays = {0, 1, 2, 3, 4}
        if scheduled > 0:
            result[employee_id] = {
                "scheduled_minutes": _round_to_30(scheduled),
                "start_minute": start_minute if start_minute is not None else 0,
                "working_weekdays": working_weekdays,
                "is_continuous": bool(is_continuous),
                "morning_in_minute": _to_minute_of_day(morning_in),
                "morning_out_minute": _to_minute_of_day(morning_out),
                "afternoon_in_minute": _to_minute_of_day(afternoon_in),
                "afternoon_out_minute": _to_minute_of_day(afternoon_out),
                "flexible_attendance": bool(is_flexible),
            }

    return result


def load_scheduled_minutes(info_path: str | Path) -> dict[str, int]:
    profiles = load_schedule_profiles(info_path)
    return {employee_id: values["scheduled_minutes"] for employee_id, values in profiles.items()}


def load_start_minutes(info_path: str | Path) -> dict[str, int]:
    profiles = load_schedule_profiles(info_path)
    return {employee_id: values["start_minute"] for employee_id, values in profiles.items()}
