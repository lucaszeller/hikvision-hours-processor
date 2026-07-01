from __future__ import annotations

from datetime import date, datetime
import re

import pandas as pd

from services.exceptions import WorkException, build_exception_index, find_matching_exceptions

DIARIO_COLUMNS = [
    "ID de persona",
    "Nombre",
    "Fecha",
    "Departamento",
    "Estado",
    "Tramos trabajados",
    "Minutos reales",
    "Minutos redondeados",
    "Minutos extra",
    "Horas extra",
    "Horas totales",
]

MENSUAL_COLUMNS = [
    "ID de persona",
    "Nombre",
    "Dias trabajados",
    "Minutos totales",
    "Minutos extra",
    "Horas extra",
    "Horas totales",
]

INCONSISTENCIAS_COLUMNS = [
    "ID de persona",
    "Nombre",
    "Fecha",
    "Tipo de inconsistencia",
    "Detalle",
]


SATURDAY_START_MINUTE = 7 * 60 + 30
SATURDAY_END_MINUTE = 11 * 60 + 30

# Tipos de excepción que cuentan 9hs (540 min). No se computan hacia el tope semanal de 44hs.
# Se comparan en minúsculas y por contenido parcial.
EIGHT_HOUR_EXCEPTION_KEYWORDS = {
    "vacacion",       # cubre "vacaciones", "vacación", "vacaciones anuales", etc.
    "feriado",
    "falta justificada",
    "faltas justificadas",
    "licencia",
    "accidente",      # cubre "accidente de trabajo", "accidente laboral", etc.
    "estudi",         # cubre "estudiar", "estudio", etc.
    "capacitacion",   # cubre "capacitaciones", "capacitación", etc.
    "viaje",          # cubre "viajes", "viaje de trabajo", etc.
    "ausencia",       # ausencia aprobada (paid_day=True) paga 8hs.
}
EIGHT_HOUR_EXCEPTION_MINUTES = 9 * 60  # 540


def _is_exception_day_type(exception_types: list[str]) -> bool:
    """Retorna True si alguno de los tipos de excepción corresponde a jornada de 8hs."""
    for exc_type in exception_types:
        normalized = exc_type.strip().lower()
        if any(keyword in normalized for keyword in EIGHT_HOUR_EXCEPTION_KEYWORDS):
            return True
    return False


# ---------------------------------------------------------------------------
# Helpers de conversión y formato
# ---------------------------------------------------------------------------

def _minutes_to_hhmm(total_minutes: int) -> str:
    hours, minutes = divmod(max(0, int(total_minutes)), 60)
    return f"{hours:02d}:{minutes:02d}"


def _round_to_30(minutes: int) -> int:
    if minutes <= 0:
        return 0
    quotient, remainder = divmod(minutes, 30)
    if remainder >= 15:
        quotient += 1
    return quotient * 30


def _floor_to_30(minutes: int) -> int:
    if minutes <= 0:
        return 0
    return (int(minutes) // 30) * 30


def _is_blank(value: object) -> bool:
    text = str(value).strip()
    return text == "" or text.lower() in {"nan", "none", "nat", "-"}


def _marker_is_positive(value: object) -> bool:
    text = str(value).strip()
    if text in {"", "-", "0", "0.0", "0:0", "0:00", "nan", "NaN", "None"}:
        return False
    if ":" in text:
        parts = [p.strip() for p in text.split(":") if p.strip() != ""]
        if not parts:
            return False
        try:
            return any(int(p) > 0 for p in parts)
        except ValueError:
            return True
    try:
        return float(text.replace(",", ".")) > 0
    except ValueError:
        return True


def _parse_date(value: object) -> datetime | None:
    if _is_blank(value):
        return None
    parsed = pd.to_datetime(str(value).strip(), errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.to_pydatetime()


def _parse_time(value: object) -> datetime | None:
    if _is_blank(value):
        return None

    text = str(value).strip()

    # Soporta formatos mixtos como "07.31.00", "07:31 hs", etc.
    match = re.search(r"(\d{1,2})[:.](\d{2})(?:[:.](\d{2}))?", text)
    if match:
        hh = int(match.group(1))
        mm = int(match.group(2))
        ss = int(match.group(3) or 0)
        if 0 <= hh <= 23 and 0 <= mm <= 59 and 0 <= ss <= 59:
            return datetime(1900, 1, 1, hh, mm, ss)

    formats = ["%H:%M:%S", "%H:%M"]
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue

    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.to_pydatetime()


def _inconsistency(
    rows: list[dict],
    employee_id: str,
    employee_name: str,
    date_value: object,
    issue_type: str,
    detail: str,
) -> None:
    rows.append(
        {
            "ID de persona": employee_id,
            "Nombre": employee_name,
            "Fecha": date_value,
            "Tipo de inconsistencia": issue_type,
            "Detalle": detail,
        }
    )


def _exception_summary(items: list[WorkException]) -> str:
    parts = []
    for item in items:
        if item.details.strip():
            parts.append(f"{item.exception_type} ({item.details})")
        else:
            parts.append(item.exception_type)
    return " | ".join(parts)


def _extract_expected_span_hhmm(schedule_text: object) -> tuple[str, str] | None:
    text = str(schedule_text or "").strip()
    if text == "":
        return None
    match = re.search(r"(\d{1,2}):(\d{2})(?::\d{2})?\s*[--]\s*(\d{1,2}):(\d{2})(?::\d{2})?", text)
    if not match:
        return None
    h1 = int(match.group(1))
    m1 = int(match.group(2))
    h2 = int(match.group(3))
    m2 = int(match.group(4))
    if not (0 <= h1 <= 23 and 0 <= m1 <= 59 and 0 <= h2 <= 23 and 0 <= m2 <= 59):
        return None
    return f"{h1:02d}:{m1:02d}", f"{h2:02d}:{m2:02d}"


def _expected_bounds_minutes_from_segment(segment: pd.Series) -> tuple[int, int] | None:
    expected_entry = segment.get("expected_entry_dt")
    expected_exit = segment.get("expected_exit_dt")
    if (
        expected_entry is not None
        and expected_exit is not None
        and not pd.isna(expected_entry)
        and not pd.isna(expected_exit)
    ):
        entry_dt = pd.to_datetime(expected_entry, errors="coerce")
        exit_dt = pd.to_datetime(expected_exit, errors="coerce")
        if pd.notna(entry_dt) and pd.notna(exit_dt):
            start_minutes = int(entry_dt.hour) * 60 + int(entry_dt.minute)
            end_minutes = int(exit_dt.hour) * 60 + int(exit_dt.minute)
            if end_minutes > start_minutes:
                return start_minutes, end_minutes

    expected = _extract_expected_span_hhmm(segment.get("schedule"))
    if expected is None:
        return None
    start_h, start_m = [int(v) for v in expected[0].split(":")]
    end_h, end_m = [int(v) for v in expected[1].split(":")]
    start_minutes = start_h * 60 + start_m
    end_minutes = end_h * 60 + end_m
    if end_minutes <= start_minutes:
        return None
    return start_minutes, end_minutes


def _overtime_minutes_by_segment(group: pd.DataFrame) -> int | None:
    had_expected = False
    overtime = 0
    for _, segment in group.iterrows():
        expected_bounds = _expected_bounds_minutes_from_segment(segment)
        if expected_bounds is None:
            continue
        had_expected = True
        expected_start, expected_end = expected_bounds

        real_entry = pd.to_datetime(segment.get("entry_dt"), errors="coerce")
        real_exit = pd.to_datetime(segment.get("exit_dt"), errors="coerce")
        if pd.isna(real_entry) or pd.isna(real_exit):
            continue
        real_start = int(real_entry.hour) * 60 + int(real_entry.minute)
        real_end = int(real_exit.hour) * 60 + int(real_exit.minute)

        early_minutes = max(0, expected_start - real_start)
        late_minutes = max(0, real_end - expected_end)
        overtime += _floor_to_30(early_minutes) + _floor_to_30(late_minutes)
    if had_expected:
        return overtime
    return None


def _format_segment_text(segment: pd.Series) -> str:
    def _pick_dt(primary: object, fallback: object) -> object:
        if primary is None or pd.isna(primary):
            return fallback
        return primary

    def _to_hhmm(value: object) -> str | None:
        if value is None or pd.isna(value):
            return None
        if isinstance(value, pd.Timestamp):
            return value.strftime("%H:%M")
        if isinstance(value, datetime):
            return value.strftime("%H:%M")
        parsed = pd.to_datetime(value, errors="coerce")
        if pd.isna(parsed):
            return None
        return parsed.strftime("%H:%M")

    display_entry = _pick_dt(segment.get("display_entry_dt"), segment.get("entry_dt"))
    display_exit = _pick_dt(segment.get("display_exit_dt"), segment.get("exit_dt"))
    real_start = _to_hhmm(display_entry)
    real_end = _to_hhmm(display_exit)
    if real_start is None or real_end is None:
        return ""
    real_span = f"{real_start}-{real_end}"

    expected_entry = _pick_dt(segment.get("expected_entry_dt"), None)
    expected_exit = _pick_dt(segment.get("expected_exit_dt"), None)
    expected_start = _to_hhmm(expected_entry)
    expected_end = _to_hhmm(expected_exit)
    if expected_start is not None and expected_end is not None:
        expected_span = f"{expected_start}-{expected_end}"
        return f"{real_span} [{expected_span}]"

    expected = _extract_expected_span_hhmm(segment.get("schedule"))
    if expected is not None:
        return f"{real_span} [{expected[0]}-{expected[1]}]"
    return real_span


def _minute_to_datetime(work_date: date, minute_of_day: int | None) -> datetime | None:
    if minute_of_day is None:
        return None
    minute_int = int(minute_of_day)
    if minute_int < 0 or minute_int >= 24 * 60:
        return None
    hour, minute = divmod(minute_int, 60)
    return datetime.combine(work_date, datetime.min.time()).replace(hour=hour, minute=minute, second=0)


def _format_real_expected_range(
    real_start: datetime,
    real_end: datetime,
    expected_start: datetime,
    expected_end: datetime,
) -> str:
    return (
        f"{real_start.strftime('%H:%M')}-{real_end.strftime('%H:%M')} "
        f"[{expected_start.strftime('%H:%M')}-{expected_end.strftime('%H:%M')}]"
    )


def _expected_bounds_for_entry_only(
    profile: dict[str, int | bool | None],
    actual_entry_minutes: int,
) -> tuple[int, int] | None:
    is_continuous = bool(profile.get("is_continuous", False))
    morning_in = profile.get("morning_in_minute")
    morning_out = profile.get("morning_out_minute")
    afternoon_in = profile.get("afternoon_in_minute")
    afternoon_out = profile.get("afternoon_out_minute")

    if is_continuous:
        start = morning_in if morning_in is not None else afternoon_in
        end = afternoon_out if afternoon_out is not None else morning_out
        if start is None or end is None or int(end) <= int(start):
            return None
        return int(start), int(end)

    if morning_in is not None and morning_out is not None and actual_entry_minutes <= int(morning_out):
        if int(morning_out) > int(morning_in):
            return int(morning_in), int(morning_out)
    if afternoon_in is not None and afternoon_out is not None and actual_entry_minutes >= int(afternoon_in):
        if int(afternoon_out) > int(afternoon_in):
            return int(afternoon_in), int(afternoon_out)

    # Fallback: usar el primer tramo valido disponible.
    if morning_in is not None and morning_out is not None and int(morning_out) > int(morning_in):
        return int(morning_in), int(morning_out)
    if afternoon_in is not None and afternoon_out is not None and int(afternoon_out) > int(afternoon_in):
        return int(afternoon_in), int(afternoon_out)
    return None


def _build_split_theoretical_segments(
    employee_id: str,
    employee_name: str,
    work_date: date,
    department: str,
    profile: dict[str, int | bool | None],
    actual_first_entry: datetime | None = None,
    actual_last_exit: datetime | None = None,
) -> list[dict]:
    result: list[dict] = []
    morning_start = _minute_to_datetime(work_date, profile.get("morning_in_minute"))
    morning_end = _minute_to_datetime(work_date, profile.get("morning_out_minute"))
    afternoon_start = _minute_to_datetime(work_date, profile.get("afternoon_in_minute"))
    afternoon_end = _minute_to_datetime(work_date, profile.get("afternoon_out_minute"))
    spans = [(morning_start, morning_end, "Horario Manana")]

    # Sabado: turno unico de 07:30 a 11:30, sin tramo de tarde.
    if work_date.weekday() != 5:
        spans.append((afternoon_start, afternoon_end, "Horario Tarde"))

    for idx, (start_dt, end_dt, schedule_label) in enumerate(spans):
        if start_dt is None or end_dt is None or end_dt <= start_dt:
            continue
        real_minutes = int(round((end_dt - start_dt).total_seconds() / 60))
        rounded_minutes = _floor_to_30(real_minutes)

        display_entry = start_dt
        display_exit = end_dt
        if idx == 0 and actual_first_entry is not None:
            display_entry = actual_first_entry
        if actual_last_exit is not None and (idx == 1 or len(spans) == 1):
            display_exit = actual_last_exit

        result.append(
            {
                "employee_id": employee_id,
                "employee_name": employee_name,
                "department": department,
                "work_date": work_date,
                "entry_dt": start_dt,
                "exit_dt": end_dt,
                "schedule": schedule_label,
                "expected_entry_dt": start_dt,
                "expected_exit_dt": end_dt,
                "display_entry_dt": display_entry,
                "display_exit_dt": display_exit,
                "real_minutes": real_minutes,
                "rounded_minutes": rounded_minutes,
            }
        )
    return result


# ---------------------------------------------------------------------------
# Funciones extraídas del proceso principal
# ---------------------------------------------------------------------------

def _suppress_inconsistencies(
    inconsistencies: list[dict],
    suppress_map: dict[tuple[str, str, date | None], set[str]],
) -> list[dict]:
    """Filtra inconsistencias cuyos tipos estén suprimidos para ese día/empleado."""
    result = []
    for issue in inconsistencies:
        issue_type = str(issue.get("Tipo de inconsistencia", "")).strip()
        issue_date_parsed = _parse_date(issue.get("Fecha"))
        issue_key = (
            str(issue.get("ID de persona", "")).strip(),
            str(issue.get("Nombre", "")).strip(),
            issue_date_parsed.date() if issue_date_parsed is not None else None,
        )
        suppressed_types = suppress_map.get(issue_key)
        if issue_key[2] is None or suppressed_types is None or issue_type not in suppressed_types:
            result.append(issue)
    return result


def _parse_punch_rows(
    df: pd.DataFrame,
    exceptions_index: object,
    scheduled_minutes_by_employee: dict[str, int],
    working_weekdays_by_employee: dict[str, set[int]],
    flexible_attendance_employee_ids: set[str],
    has_schedule_reference: bool,
) -> tuple[list[dict], list[dict], dict, dict, dict[str, str], dict[str, set[str]], set]:
    """
    Itera sobre cada fila del reporte Hikvision: parsea empleado, fecha y horarios;
    detecta inconsistencias; y agrupa tramos en valid_segments y partial_day_punches.

    Retorna:
        valid_segments, inconsistencies, partial_day_punches, day_flags,
        employee_name_by_id, departments_by_id, used_exception_keys
    """
    valid_segments: list[dict] = []
    inconsistencies: list[dict] = []
    partial_day_punches: dict[tuple[str, str, date], dict] = {}
    day_flags: dict[tuple[str, str, date], dict] = {}
    employee_name_by_id: dict[str, str] = {}
    departments_by_id: dict[str, set[str]] = {}
    used_exception_keys: set[tuple[str | None, date, str, str]] = set()

    for _, row in df.iterrows():
        employee_id = str(row.get("employee_id", "")).strip()
        employee_name = str(row.get("employee_name", "")).strip()
        department = str(row.get("department", "")).strip()
        schedule = str(row.get("schedule", "")).strip()
        work_date_raw = row.get("work_date_raw", "")
        entry_raw = row.get("entry_time_raw", "")
        exit_raw = row.get("exit_time_raw", "")

        date_for_report = work_date_raw

        if employee_id != "":
            if employee_name and employee_id not in employee_name_by_id:
                employee_name_by_id[employee_id] = employee_name
            if department:
                departments_by_id.setdefault(employee_id, set()).add(department)

        if employee_id == "":
            _inconsistency(
                inconsistencies, employee_id, employee_name, date_for_report,
                "ID de persona vacio", "La fila no contiene ID de persona.",
            )
        if employee_name == "":
            _inconsistency(
                inconsistencies, employee_id, employee_name, date_for_report,
                "Nombre vacio", "La fila no contiene nombre de empleado.",
            )

        date_parsed = _parse_date(work_date_raw)
        row_exceptions: list[WorkException] = []
        if date_parsed is not None and employee_id != "":
            row_exceptions = find_matching_exceptions(exceptions_index, employee_id, date_parsed.date())
            for item in row_exceptions:
                used_exception_keys.add(
                    (item.employee_id, item.exception_date, item.exception_type.strip(), item.details.strip())
                )

        if date_parsed is None:
            _inconsistency(
                inconsistencies, employee_id, employee_name, date_for_report,
                "Fecha invalida", f"Valor recibido: '{work_date_raw}'.",
            )

        has_entry = not _is_blank(entry_raw)
        has_exit = not _is_blank(exit_raw)
        weekday = date_parsed.weekday() if date_parsed is not None else None
        is_saturday = weekday == 5
        is_sunday = weekday == 6
        is_weekend = is_saturday or is_sunday
        configured_days = working_weekdays_by_employee.get(employee_id, {0, 1, 2, 3, 4})
        is_non_working_day = weekday is not None and weekday not in configured_days
        is_flexible_attendance = str(employee_id).strip() in flexible_attendance_employee_ids

        # Regla: lunes a viernes normales; fines de semana sólo si hay fichada.
        if is_weekend and not (has_entry or has_exit):
            continue
        # Regla por empleado: si no trabaja ese dia y no hay fichada, no marcar ausente.
        if is_non_working_day and not (has_entry or has_exit):
            continue

        late_marker = _marker_is_positive(row.get("late_raw", ""))
        absent_marker = _marker_is_positive(row.get("absent_raw", ""))

        if employee_id != "" and employee_name != "" and date_parsed is not None:
            day_key = (employee_id, employee_name, date_parsed.date())
            day_state = day_flags.setdefault(
                day_key,
                {
                    "late": False,
                    "absent": False,
                    "departments": set(),
                    "exception_types": [],
                    "paid_exception": False,
                },
            )
            if department:
                day_state["departments"].add(department)
            if not is_flexible_attendance:
                day_state["late"] = bool(day_state["late"]) or late_marker
            missing_both_without_exception = (
                not has_entry
                and not has_exit
                and not row_exceptions
                and not is_non_working_day
            )
            if is_flexible_attendance:
                missing_both_without_exception = False
            # El marcador "Ausente" de Hikvision puede venir en 1 aun con fichadas
            # (por configuracion de iVMS/horario). Solo lo tomamos si no hay fichadas.
            absent_marker_without_punches = absent_marker and not (has_entry or has_exit)
            if is_flexible_attendance:
                absent_marker_without_punches = False
            day_state["absent"] = (
                bool(day_state["absent"])
                or absent_marker_without_punches
                or missing_both_without_exception
            )
            for exc in row_exceptions:
                exc_type = str(exc.exception_type).strip()
                if exc_type and exc_type not in day_state["exception_types"]:
                    day_state["exception_types"].append(exc_type)
                if bool(getattr(exc, "paid_day", False)) or str(exc_type).strip().lower() == "presente":
                    day_state["paid_exception"] = True

        if not has_entry and not has_exit:
            if is_flexible_attendance and not row_exceptions:
                continue
            if row_exceptions:
                _inconsistency(
                    inconsistencies, employee_id, employee_name, date_for_report,
                    "Excepcion aplicada",
                    "Se omitio inconsistencia de ausencia sin fichadas por excepcion: "
                    + _exception_summary(row_exceptions),
                )
            else:
                _inconsistency(
                    inconsistencies, employee_id, employee_name, date_for_report,
                    "Ausente", "Faltan Registro de entrada y Registro de salida.",
                )
        elif not has_entry:
            if row_exceptions:
                _inconsistency(
                    inconsistencies, employee_id, employee_name, date_for_report,
                    "Excepcion aplicada",
                    "Se omitio inconsistencia de falta de entrada por excepcion: "
                    + _exception_summary(row_exceptions),
                )
            else:
                _inconsistency(
                    inconsistencies, employee_id, employee_name, date_for_report,
                    "Falta Registro de entrada", "No se encontro hora de entrada.",
                )
        elif not has_exit:
            if row_exceptions:
                _inconsistency(
                    inconsistencies, employee_id, employee_name, date_for_report,
                    "Excepcion aplicada",
                    "Se omitio inconsistencia de falta de salida por excepcion: "
                    + _exception_summary(row_exceptions),
                )
            else:
                _inconsistency(
                    inconsistencies, employee_id, employee_name, date_for_report,
                    "Falta Registro de salida", "No se encontro hora de salida.",
                )

        entry_time = _parse_time(entry_raw) if has_entry else None
        exit_time = _parse_time(exit_raw) if has_exit else None

        if has_entry and entry_time is None:
            _inconsistency(
                inconsistencies, employee_id, employee_name, date_for_report,
                "Hora de entrada invalida", f"Valor recibido: '{entry_raw}'.",
            )
        if has_exit and exit_time is None:
            _inconsistency(
                inconsistencies, employee_id, employee_name, date_for_report,
                "Hora de salida invalida", f"Valor recibido: '{exit_raw}'.",
            )

        if employee_id == "" or employee_name == "" or date_parsed is None:
            continue

        day_key = (employee_id, employee_name, date_parsed.date())
        partial_state = partial_day_punches.setdefault(
            day_key,
            {"entries": [], "exits": [], "departments": set(), "schedules": set()},
        )
        if department:
            partial_state["departments"].add(department)
        if schedule:
            partial_state["schedules"].add(schedule)
        if entry_time is not None:
            partial_state["entries"].append(datetime.combine(date_parsed.date(), entry_time.time()))
        if exit_time is not None:
            partial_state["exits"].append(datetime.combine(date_parsed.date(), exit_time.time()))

        if entry_time is None or exit_time is None:
            continue

        start_dt = datetime.combine(date_parsed.date(), entry_time.time())
        end_dt = datetime.combine(date_parsed.date(), exit_time.time())

        if end_dt <= start_dt:
            if row_exceptions:
                _inconsistency(
                    inconsistencies, employee_id, employee_name, date_parsed.date(),
                    "Excepcion aplicada",
                    "Se omitio inconsistencia de salida menor/igual que entrada por excepcion: "
                    + _exception_summary(row_exceptions),
                )
            else:
                _inconsistency(
                    inconsistencies, employee_id, employee_name, date_parsed.date(),
                    "Salida menor que entrada",
                    f"Entrada {entry_raw} / Salida {exit_raw}.",
                )
            continue

        real_minutes = int(round((end_dt - start_dt).total_seconds() / 60))
        rounded_minutes = _floor_to_30(real_minutes)

        valid_segments.append(
            {
                "employee_id": employee_id,
                "employee_name": employee_name,
                "department": department,
                "work_date": date_parsed.date(),
                "entry_dt": start_dt,
                "exit_dt": end_dt,
                "schedule": schedule,
                "real_minutes": real_minutes,
                "rounded_minutes": rounded_minutes,
            }
        )

    return (
        valid_segments,
        inconsistencies,
        partial_day_punches,
        day_flags,
        employee_name_by_id,
        departments_by_id,
        used_exception_keys,
    )


def _find_split_two_punch_days(
    partial_day_punches: dict,
    split_schedule_by_employee: dict,
) -> tuple[set, set, dict, dict, dict]:
    """
    Detecta días donde el empleado tiene horario partido y solo registró
    2 fichadas (1 entrada + 1 salida).

    Si las 2 fichadas cubren ambos turnos → split_two_punch_days: se reconstruyen
    tramos teóricos para mañana y tarde.

    Si las 2 fichadas cubren solo un turno (solo mañana o solo tarde) →
    partial_one_span_days: se procesan con las horas reales y se marcan como Tardanza.
    El punto de corte es el punto medio entre fin de mañana e inicio de tarde.

    Retorna:
        split_two_punch_days, partial_one_span_days,
        actual_first_entry_minutes_by_day, actual_first_entry_by_day,
        actual_last_exit_by_day
    """
    split_two_punch_days: set[tuple[str, str, date]] = set()
    partial_one_span_days: set[tuple[str, str, date]] = set()
    actual_first_entry_minutes_by_day: dict[tuple[str, str, date], int] = {}
    actual_first_entry_by_day: dict[tuple[str, str, date], datetime] = {}
    actual_last_exit_by_day: dict[tuple[str, str, date], datetime] = {}

    for day_key, data in partial_day_punches.items():
        employee_id, _employee_name, _work_date = day_key
        profile = split_schedule_by_employee.get(str(employee_id).strip(), {})
        if bool(profile.get("is_continuous", False)):
            continue
        entries = list(data.get("entries", []))
        exits = list(data.get("exits", []))
        if not entries or not exits:
            continue
        total_punches = len(entries) + len(exits)
        if total_punches != 2:
            continue
        has_theoretical_both_spans = (
            profile.get("morning_in_minute") is not None
            and profile.get("morning_out_minute") is not None
            and profile.get("afternoon_in_minute") is not None
            and profile.get("afternoon_out_minute") is not None
        )
        if not has_theoretical_both_spans:
            continue

        first_entry = min(entries)
        last_exit = max(exits)
        first_entry_min = first_entry.hour * 60 + first_entry.minute
        last_exit_min = last_exit.hour * 60 + last_exit.minute

        morning_out: int = int(profile["morning_out_minute"])
        afternoon_in: int = int(profile["afternoon_in_minute"])
        midpoint: int = (morning_out + afternoon_in) // 2

        if last_exit_min <= midpoint or first_entry_min >= midpoint:
            # Solo un turno trabajado → Tardanza, horas reales
            partial_one_span_days.add(day_key)
        else:
            # Cubre ambos turnos → reconstruir tramos teóricos
            split_two_punch_days.add(day_key)

        actual_first_entry_minutes_by_day[day_key] = first_entry_min
        actual_first_entry_by_day[day_key] = first_entry
        actual_last_exit_by_day[day_key] = last_exit

    return (
        split_two_punch_days,
        partial_one_span_days,
        actual_first_entry_minutes_by_day,
        actual_first_entry_by_day,
        actual_last_exit_by_day,
    )


def _rebuild_segments(
    valid_segments: list[dict],
    partial_day_punches: dict,
    day_flags: dict,
    split_two_punch_days: set,
    actual_first_entry_minutes_by_day: dict,
    actual_first_entry_by_day: dict,
    actual_last_exit_by_day: dict,
    split_schedule_by_employee: dict,
) -> tuple[list[dict], dict, set, set, dict]:
    """
    Reconstruye tramos para:
    - Días con horario partido y solo 2 fichadas (usa tramos teóricos).
    - Días con solo entrada (infiere salida por horario esperado).
    - Días con entrada y salida en filas separadas (consolida en un solo tramo).

    Retorna:
        valid_segments, day_flags, existing_day_keys,
        inferred_entry_only_days, actual_first_entry_minutes_by_day
    """
    # Elimina tramos crudos de los días que serán reemplazados por tramos teóricos
    preserved_segments: list[dict] = []
    for seg in valid_segments:
        seg_key = (str(seg["employee_id"]), str(seg["employee_name"]), seg["work_date"])
        if seg_key in split_two_punch_days:
            continue
        preserved_segments.append(seg)
    valid_segments = preserved_segments

    existing_day_keys = {
        (str(seg["employee_id"]), str(seg["employee_name"]), seg["work_date"])
        for seg in valid_segments
    }
    inferred_entry_only_days: set[tuple[str, str, date]] = set()

    for day_key, data in partial_day_punches.items():
        employee_id, employee_name, work_date = day_key

        # --- Días con horario partido y solo 2 fichadas: tramos teóricos ---
        if day_key in split_two_punch_days:
            profile = split_schedule_by_employee.get(str(employee_id).strip(), {})
            department = " / ".join(
                sorted([d for d in data.get("departments", set()) if str(d).strip() != ""])
            )
            synthetic_segments = _build_split_theoretical_segments(
                employee_id=employee_id,
                employee_name=employee_name,
                work_date=work_date,
                department=department,
                profile=profile,
                actual_first_entry=actual_first_entry_by_day.get(day_key),
                actual_last_exit=actual_last_exit_by_day.get(day_key),
            )
            if synthetic_segments:
                valid_segments.extend(synthetic_segments)
                existing_day_keys.add(day_key)
                if day_key in day_flags:
                    day_flags[day_key]["absent"] = False
            continue

        entries = list(data.get("entries", []))
        exits = list(data.get("exits", []))

        # --- Días con solo entrada: infiere salida por horario esperado ---
        if day_key not in existing_day_keys and entries and not exits:
            profile = split_schedule_by_employee.get(str(employee_id).strip(), {})
            first_entry = min(entries)
            first_entry_min = first_entry.hour * 60 + first_entry.minute
            bounds = _expected_bounds_for_entry_only(profile, first_entry_min)
            if bounds is not None:
                expected_start_dt = _minute_to_datetime(work_date, bounds[0])
                expected_end_dt = _minute_to_datetime(work_date, bounds[1])
                if expected_start_dt is not None and expected_end_dt is not None and expected_end_dt > first_entry:
                    department = " / ".join(
                        sorted([d for d in data.get("departments", set()) if str(d).strip() != ""])
                    )
                    real_minutes = int(round((expected_end_dt - first_entry).total_seconds() / 60))
                    rounded_minutes = _floor_to_30(real_minutes)
                    valid_segments.append(
                        {
                            "employee_id": employee_id,
                            "employee_name": employee_name,
                            "department": department,
                            "work_date": work_date,
                            "entry_dt": first_entry,
                            "exit_dt": expected_end_dt,
                            "schedule": f"Esperado({expected_start_dt.strftime('%H:%M')}-{expected_end_dt.strftime('%H:%M')})",
                            "expected_entry_dt": expected_start_dt,
                            "expected_exit_dt": expected_end_dt,
                            "display_entry_dt": first_entry,
                            "display_exit_dt": expected_end_dt,
                            "real_minutes": real_minutes,
                            "rounded_minutes": rounded_minutes,
                        }
                    )
                    existing_day_keys.add(day_key)
                    inferred_entry_only_days.add(day_key)
                    actual_first_entry_minutes_by_day[day_key] = first_entry_min
                    if day_key in day_flags:
                        day_flags[day_key]["absent"] = False
                    continue

        # --- Ya tiene tramos completos: no hace falta reconstruir ---
        if day_key in existing_day_keys:
            continue
        if not entries or not exits:
            continue

        # --- Consolidar: primera entrada → última salida ---
        start_dt = min(entries)
        end_dt = max(exits)
        if end_dt <= start_dt:
            continue

        department = " / ".join(
            sorted([d for d in data.get("departments", set()) if str(d).strip() != ""])
        )
        schedule_value = " | ".join(
            sorted([s for s in data.get("schedules", set()) if str(s).strip() != ""])
        )
        real_minutes = int(round((end_dt - start_dt).total_seconds() / 60))
        rounded_minutes = _floor_to_30(real_minutes)

        valid_segments.append(
            {
                "employee_id": employee_id,
                "employee_name": employee_name,
                "department": department,
                "work_date": work_date,
                "entry_dt": start_dt,
                "exit_dt": end_dt,
                "schedule": schedule_value,
                "real_minutes": real_minutes,
                "rounded_minutes": rounded_minutes,
            }
        )

        if day_key in day_flags:
            day_flags[day_key]["absent"] = False

    return (
        valid_segments,
        day_flags,
        existing_day_keys,
        inferred_entry_only_days,
        actual_first_entry_minutes_by_day,
    )


def _inject_exception_days(
    exceptions: list[WorkException],
    used_exception_keys: set,
    known_employee_ids: set[str],
    employee_name_by_id: dict[str, str],
    departments_by_id: dict[str, set[str]],
    day_flags: dict,
) -> list[dict]:
    """
    Asegura que las excepciones (feriados, ausencias) aparezcan en el reporte
    aunque no exista ninguna fila Hikvision para ese empleado y fecha.
    Retorna inconsistencias por excepciones configuradas sin uso.
    """
    extra_inconsistencies: list[dict] = []

    for item in exceptions:
        employee_id = (item.employee_id or "").strip()
        if employee_id == "":
            # Excepcion global: aplica a todos los empleados conocidos.
            target_ids = sorted([emp_id for emp_id in known_employee_ids if emp_id != ""])
        else:
            target_ids = [employee_id]

        for target_id in target_ids:
            employee_name = employee_name_by_id.get(target_id, "")
            day_key = (target_id, employee_name, item.exception_date)
            day_state = day_flags.setdefault(
                day_key,
                {
                    "late": False,
                    "absent": False,
                    "departments": set(),
                    "exception_types": [],
                    "paid_exception": False,
                },
            )
            for dep in departments_by_id.get(target_id, set()):
                day_state["departments"].add(dep)

            exc_type = str(item.exception_type).strip()
            if exc_type and exc_type not in day_state["exception_types"]:
                day_state["exception_types"].append(exc_type)
            if bool(getattr(item, "paid_day", False)) or str(exc_type).strip().lower() == "presente":
                day_state["paid_exception"] = True

        used_exception_keys.add(
            (item.employee_id, item.exception_date, item.exception_type.strip(), item.details.strip())
        )

    for item in exceptions:
        key = (item.employee_id, item.exception_date, item.exception_type.strip(), item.details.strip())
        if key in used_exception_keys:
            continue
        _inconsistency(
            extra_inconsistencies,
            item.employee_id or "",
            "",
            item.exception_date,
            "Excepcion configurada sin uso",
            f"{item.exception_type}. {item.details}".strip(),
        )

    return extra_inconsistencies


def _build_daily_rows(
    segments_df: pd.DataFrame,
    day_flags: dict,
    split_two_punch_days: set,
    actual_first_entry_minutes_by_day: dict,
    actual_first_entry_by_day: dict,
    actual_last_exit_by_day: dict,
    scheduled_minutes_by_employee: dict[str, int],
    scheduled_start_minute_by_employee: dict[str, int],
    working_weekdays_by_employee: dict[str, set[int]],
    split_schedule_by_employee: dict,
    flexible_attendance_employee_ids: set[str],
    has_schedule_reference: bool,
    partial_one_span_days: set | None = None,
) -> tuple[list[dict], list[dict]]:
    """
    Agrupa tramos por empleado/día, calcula horas extra y clasifica el estado
    (Normal, Tarde, Ausente, excepción, etc.).

    Retorna: grouped_rows, extra_inconsistencies
    """
    if partial_one_span_days is None:
        partial_one_span_days = set()
    extra_inconsistencies: list[dict] = []

    day_rows: dict[tuple[str, str, date], dict] = {}
    day_first_entry_minutes: dict[tuple[str, str, date], int] = {}

    for (employee_id, employee_name, work_date), group in segments_df.groupby(
        ["employee_id", "employee_name", "work_date"],
        sort=True,
    ):
        day_key = (str(employee_id), str(employee_name), work_date)
        departments = [d for d in group["department"].astype(str).str.strip().unique() if d]
        department_value = " / ".join(departments)

        segments_text: list[str] = []
        for _, segment in group.iterrows():
            segments_text.append(_format_segment_text(segment))

        real_total = int(group["real_minutes"].sum())
        rounded_total = int(group["rounded_minutes"].sum())

        if work_date.weekday() == 5:
            saturday_first_entry = actual_first_entry_by_day.get(day_key, group["entry_dt"].min())
            saturday_last_exit = actual_last_exit_by_day.get(day_key, group["exit_dt"].max())
            if isinstance(saturday_first_entry, datetime) and isinstance(saturday_last_exit, datetime):
                if saturday_last_exit > saturday_first_entry:
                    real_total = int(round((saturday_last_exit - saturday_first_entry).total_seconds() / 60))
                    rounded_total = _floor_to_30(real_total)
                    segments_text = [
                        (
                            f"{saturday_first_entry.strftime('%H:%M')}-{saturday_last_exit.strftime('%H:%M')} "
                            f"[{SATURDAY_START_MINUTE // 60:02d}:{SATURDAY_START_MINUTE % 60:02d}-{SATURDAY_END_MINUTE // 60:02d}:{SATURDAY_END_MINUTE % 60:02d}]"
                        )
                    ]

        first_entry = group["entry_dt"].min()
        first_entry_minutes = first_entry.hour * 60 + first_entry.minute
        if day_key in actual_first_entry_minutes_by_day:
            first_entry_minutes = int(actual_first_entry_minutes_by_day[day_key])
        day_first_entry_minutes[day_key] = first_entry_minutes

        employee_days = working_weekdays_by_employee.get(str(employee_id).strip(), {0, 1, 2, 3, 4})
        employee_profile = split_schedule_by_employee.get(str(employee_id).strip(), {})
        is_flexible_attendance = str(employee_id).strip() in flexible_attendance_employee_ids

        if is_flexible_attendance:
            overtime_minutes = 0
        elif work_date.weekday() in {5, 6} or work_date.weekday() not in employee_days:
            overtime_minutes = rounded_total
        elif day_key in split_two_punch_days:
            overtime_minutes = 0
        else:
            scheduled_minutes = scheduled_minutes_by_employee.get(str(employee_id).strip())
            if scheduled_minutes is None:
                overtime_minutes = 0
                if has_schedule_reference:
                    _inconsistency(
                        extra_inconsistencies,
                        str(employee_id),
                        str(employee_name),
                        work_date,
                        "Horario no definido",
                        "No hay horario configurado en info.xlsx para calcular horas extra.",
                    )
            else:
                overtime_by_segment = _overtime_minutes_by_segment(group)
                if overtime_by_segment is not None:
                    overtime_minutes = overtime_by_segment
                else:
                    overtime_minutes = _floor_to_30(max(0, real_total - int(scheduled_minutes)))

        # Caso: horario corrido sin horas extra — mostrar rango real vs esperado.
        if (
            bool(employee_profile.get("is_continuous", False))
            and overtime_minutes == 0
            and len(group) == 1
        ):
            expected_start_min = employee_profile.get("morning_in_minute")
            if expected_start_min is None:
                expected_start_min = employee_profile.get("afternoon_in_minute")
            expected_end_min = employee_profile.get("afternoon_out_minute")
            if expected_end_min is None:
                expected_end_min = employee_profile.get("morning_out_minute")

            expected_start_dt = _minute_to_datetime(work_date, expected_start_min)
            expected_end_dt = _minute_to_datetime(work_date, expected_end_min)
            real_start_dt = group.iloc[0]["entry_dt"]
            real_end_dt = group.iloc[0]["exit_dt"]

            if (
                isinstance(real_start_dt, datetime)
                and isinstance(real_end_dt, datetime)
                and expected_start_dt is not None
                and expected_end_dt is not None
                and expected_end_dt > expected_start_dt
            ):
                segments_text = [
                    _format_real_expected_range(
                        real_start_dt, real_end_dt, expected_start_dt, expected_end_dt,
                    )
                ]

        day_rows[day_key] = {
            "ID de persona": employee_id,
            "Nombre": employee_name,
            "Fecha": work_date,
            "Departamento": department_value,
            "Tramos trabajados": " ||| ".join(segments_text),
            "Minutos reales": real_total,
            "Minutos redondeados": rounded_total,
            "Minutos extra": overtime_minutes,
            "Horas extra": _minutes_to_hhmm(overtime_minutes),
            "Horas totales": _minutes_to_hhmm(rounded_total),
        }

    # --- Clasificar estado y construir filas finales ---
    grouped_rows: list[dict] = []
    all_day_keys = sorted(
        set(day_rows.keys()) | set(day_flags.keys()),
        key=lambda k: (k[0], k[2], k[1]),
    )

    for day_key in all_day_keys:
        employee_id, employee_name, work_date = day_key
        is_flexible_attendance = str(employee_id).strip() in flexible_attendance_employee_ids
        state = day_flags.get(
            day_key,
            {
                "late": False,
                "absent": False,
                "departments": set(),
                "exception_types": [],
                "paid_exception": False,
            },
        )
        row_data = day_rows.get(day_key)
        exception_types = [
            str(v).strip() for v in (state.get("exception_types") or []) if str(v).strip() != ""
        ]
        exception_types_lower = [v.lower() for v in exception_types]
        has_presente = "presente" in exception_types_lower
        has_non_presente_exception = any(v != "presente" for v in exception_types_lower)
        suppress_presente_by_saturday_exception = (
            work_date.weekday() == 5 and has_presente and has_non_presente_exception
        )
        apply_presente = has_presente and not suppress_presente_by_saturday_exception
        preferred_exception = exception_types[0] if exception_types else "Normal"
        if suppress_presente_by_saturday_exception:
            preferred_exception = next(
                (v for v in exception_types if v.strip().lower() != "presente"),
                preferred_exception,
            )
        saturday_has_exception = work_date.weekday() == 5 and bool(exception_types)

        if row_data is None:
            if (
                not bool(state.get("absent"))
                and not bool(state.get("late"))
                and not state.get("exception_types")
            ):
                continue
            departments = sorted([d for d in state["departments"] if str(d).strip() != ""])
            paid_exception = bool(state.get("paid_exception"))
            scheduled_minutes = scheduled_minutes_by_employee.get(str(employee_id).strip(), 0)
            if apply_presente:
                paid_minutes = int(scheduled_minutes) if int(scheduled_minutes) > 0 else 0
            elif paid_exception and int(scheduled_minutes) > 0:
                if _is_exception_day_type(exception_types):
                    paid_minutes = EIGHT_HOUR_EXCEPTION_MINUTES
                else:
                    paid_minutes = int(scheduled_minutes)
            else:
                paid_minutes = 0
            row_data = {
                "ID de persona": employee_id,
                "Nombre": employee_name,
                "Fecha": work_date,
                "Departamento": " / ".join(departments),
                "Tramos trabajados": "",
                "Minutos reales": paid_minutes,
                "Minutos redondeados": paid_minutes,
                "Minutos extra": 0,
                "Horas extra": "00:00",
                "Horas totales": _minutes_to_hhmm(paid_minutes),
            }
        elif apply_presente:
            # Regla: si hay PRESENTE en date.xlsx, ignora fichadas reales del dia.
            scheduled_minutes = int(scheduled_minutes_by_employee.get(str(employee_id).strip(), 0) or 0)
            row_data = {
                "ID de persona": employee_id,
                "Nombre": employee_name,
                "Fecha": work_date,
                "Departamento": row_data.get("Departamento", ""),
                "Tramos trabajados": "",
                "Minutos reales": scheduled_minutes,
                "Minutos redondeados": scheduled_minutes,
                "Minutos extra": 0,
                "Horas extra": "00:00",
                "Horas totales": _minutes_to_hhmm(scheduled_minutes),
            }
        elif suppress_presente_by_saturday_exception:
            # Sabado con excepcion real en date.xlsx: ignora fichadas y procesa por excepcion.
            departments = sorted([d for d in state["departments"] if str(d).strip() != ""])
            if not departments and row_data is not None:
                departments = [str(row_data.get("Departamento", "")).strip()]
            paid_exception = bool(state.get("paid_exception"))
            scheduled_minutes = int(scheduled_minutes_by_employee.get(str(employee_id).strip(), 0) or 0)
            paid_minutes = scheduled_minutes if paid_exception and scheduled_minutes > 0 else 0
            row_data = {
                "ID de persona": employee_id,
                "Nombre": employee_name,
                "Fecha": work_date,
                "Departamento": " / ".join([d for d in departments if d]),
                "Tramos trabajados": "",
                "Minutos reales": paid_minutes,
                "Minutos redondeados": paid_minutes,
                "Minutos extra": 0,
                "Horas extra": "00:00",
                "Horas totales": _minutes_to_hhmm(paid_minutes),
            }

        if saturday_has_exception:
            departments = sorted([d for d in state["departments"] if str(d).strip() != ""])
            if not departments and row_data is not None:
                departments = [str(row_data.get("Departamento", "")).strip()]
            row_data = {
                "ID de persona": employee_id,
                "Nombre": employee_name,
                "Fecha": work_date,
                "Departamento": " / ".join([d for d in departments if d]),
                "Tramos trabajados": "",
                "Minutos reales": 0,
                "Minutos redondeados": 0,
                "Minutos extra": 0,
                "Horas extra": "00:00",
                "Horas totales": "00:00",
                "Estado": preferred_exception.title(),
            }
            grouped_rows.append(row_data)
            continue

        worked_minutes = int(row_data.get("Minutos redondeados", 0))
        if worked_minutes > 0:
            if bool(state.get("paid_exception")) and exception_types and str(
                row_data.get("Tramos trabajados", "")
            ).strip() == "":
                row_data["Estado"] = preferred_exception.title()
                grouped_rows.append(row_data)
                continue
            if work_date.weekday() == 6:
                row_data["Estado"] = "Domingo"
                grouped_rows.append(row_data)
                continue
            if is_flexible_attendance:
                row_data["Estado"] = "Normal"
                grouped_rows.append(row_data)
                continue
            # Horario partido con solo un turno trabajado → siempre Tardanza.
            if day_key in partial_one_span_days:
                row_data["Estado"] = "Tarde"
                grouped_rows.append(row_data)
                continue
            # Con fichadas validas no debe clasificarse Ausente.
            first_entry_minutes = day_first_entry_minutes.get(day_key)
            if work_date.weekday() == 5 and first_entry_minutes is not None:
                # Regla sabado: horario fijo 07:30 para todos.
                status = "Tarde" if first_entry_minutes > SATURDAY_START_MINUTE else "Normal"
            else:
                scheduled_start = scheduled_start_minute_by_employee.get(str(employee_id).strip())
                if scheduled_start is not None and first_entry_minutes is not None:
                    status = "Tarde" if first_entry_minutes > int(scheduled_start) else "Normal"
                else:
                    status = "Tarde" if bool(state.get("late")) else "Normal"
        else:
            if exception_types:
                status = preferred_exception.title()
            elif bool(state.get("absent")) and not is_flexible_attendance:
                status = "Ausente"
            elif bool(state.get("late")) and not is_flexible_attendance:
                status = "Tarde"
            else:
                status = "Normal"

        row_data["Estado"] = status
        grouped_rows.append(row_data)

    return grouped_rows, extra_inconsistencies


# ---------------------------------------------------------------------------
# Función principal (orquestador)
# ---------------------------------------------------------------------------

def process_punches(
    df: pd.DataFrame,
    exceptions: list[WorkException] | None = None,
    scheduled_minutes_by_employee: dict[str, int] | None = None,
    scheduled_start_minute_by_employee: dict[str, int] | None = None,
    working_weekdays_by_employee: dict[str, set[int]] | None = None,
    split_schedule_by_employee: dict[str, dict[str, int | bool | None]] | None = None,
    flexible_attendance_employee_ids: set[str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    # --- Valores por defecto ---
    exceptions = exceptions or []
    scheduled_minutes_by_employee = scheduled_minutes_by_employee or {}
    scheduled_start_minute_by_employee = scheduled_start_minute_by_employee or {}
    working_weekdays_by_employee = working_weekdays_by_employee or {}
    split_schedule_by_employee = split_schedule_by_employee or {}
    flexible_attendance_employee_ids = {
        str(v).strip() for v in (flexible_attendance_employee_ids or set()) if str(v).strip() != ""
    }
    has_schedule_reference = bool(scheduled_minutes_by_employee)
    exceptions_index = build_exception_index(exceptions)

    # --- Paso 1: parsear todas las filas del reporte ---
    (
        valid_segments,
        inconsistencies,
        partial_day_punches,
        day_flags,
        employee_name_by_id,
        departments_by_id,
        used_exception_keys,
    ) = _parse_punch_rows(
        df,
        exceptions_index,
        scheduled_minutes_by_employee,
        working_weekdays_by_employee,
        flexible_attendance_employee_ids,
        has_schedule_reference,
    )

    # --- Paso 2: identificar días con horario partido y solo 2 fichadas ---
    (
        split_two_punch_days,
        partial_one_span_days,
        actual_first_entry_minutes_by_day,
        actual_first_entry_by_day,
        actual_last_exit_by_day,
    ) = _find_split_two_punch_days(partial_day_punches, split_schedule_by_employee)

    # --- Paso 3: suprimir inconsistencias provocadas por días de horario partido ---
    all_two_punch_days = split_two_punch_days | partial_one_span_days
    if all_two_punch_days:
        suppress_map = {
            day_key: {"Falta Registro de entrada", "Falta Registro de salida"}
            for day_key in all_two_punch_days
        }
        inconsistencies = _suppress_inconsistencies(inconsistencies, suppress_map)

    # --- Paso 4: reconstruir tramos (horario partido, entrada sola, consolidación) ---
    (
        valid_segments,
        day_flags,
        _existing_day_keys,
        inferred_entry_only_days,
        actual_first_entry_minutes_by_day,
    ) = _rebuild_segments(
        valid_segments,
        partial_day_punches,
        day_flags,
        split_two_punch_days,
        actual_first_entry_minutes_by_day,
        actual_first_entry_by_day,
        actual_last_exit_by_day,
        split_schedule_by_employee,
    )

    # --- Paso 5: suprimir inconsistencias de días con solo entrada inferida ---
    if inferred_entry_only_days:
        suppress_map_entry_only = {
            day_key: {"Falta Registro de salida"}
            for day_key in inferred_entry_only_days
        }
        inconsistencies = _suppress_inconsistencies(inconsistencies, suppress_map_entry_only)

    # --- Paso 6: inyectar días de excepción sin fila Hikvision (feriados, ausencias) ---
    known_employee_ids: set[str] = set(employee_name_by_id.keys()) | {
        str(emp_id).strip() for emp_id in scheduled_minutes_by_employee.keys()
    }
    exception_inconsistencies = _inject_exception_days(
        exceptions,
        used_exception_keys,
        known_employee_ids,
        employee_name_by_id,
        departments_by_id,
        day_flags,
    )
    inconsistencies.extend(exception_inconsistencies)

    # --- Paso 7: construir DataFrames diario y mensual ---
    segments_df = pd.DataFrame(valid_segments)

    if segments_df.empty and not day_flags:
        diario_df = pd.DataFrame(columns=DIARIO_COLUMNS)
        mensual_df = pd.DataFrame(columns=MENSUAL_COLUMNS)
    else:
        if segments_df.empty:
            segments_df = pd.DataFrame(
                columns=[
                    "employee_id", "employee_name", "department", "work_date",
                    "entry_dt", "exit_dt", "schedule", "real_minutes", "rounded_minutes",
                ]
            )

        segments_df = segments_df.sort_values(
            ["employee_id", "work_date", "entry_dt", "schedule"], kind="stable"
        ).reset_index(drop=True)

        grouped_rows, row_inconsistencies = _build_daily_rows(
            segments_df,
            day_flags,
            split_two_punch_days,
            actual_first_entry_minutes_by_day,
            actual_first_entry_by_day,
            actual_last_exit_by_day,
            scheduled_minutes_by_employee,
            scheduled_start_minute_by_employee,
            working_weekdays_by_employee,
            split_schedule_by_employee,
            flexible_attendance_employee_ids,
            has_schedule_reference,
            partial_one_span_days=partial_one_span_days,
        )
        inconsistencies.extend(row_inconsistencies)

        diario_df = pd.DataFrame(grouped_rows, columns=DIARIO_COLUMNS)
        diario_df = diario_df.sort_values(
            ["ID de persona", "Fecha", "Nombre"], kind="stable"
        ).reset_index(drop=True)

        mensual_df = (
            diario_df.groupby(["ID de persona", "Nombre"], as_index=False)
            .agg(
                {
                    "Fecha": "nunique",
                    "Minutos redondeados": "sum",
                    "Minutos extra": "sum",
                }
            )
            .rename(
                columns={
                    "Fecha": "Dias trabajados",
                    "Minutos redondeados": "Minutos totales",
                }
            )
        )
        mensual_df["Horas extra"] = mensual_df["Minutos extra"].map(_minutes_to_hhmm)
        mensual_df["Horas totales"] = mensual_df["Minutos totales"].map(_minutes_to_hhmm)
        mensual_df = mensual_df[MENSUAL_COLUMNS].sort_values(
            ["ID de persona", "Nombre"], kind="stable"
        ).reset_index(drop=True)

    inconsistencias_df = pd.DataFrame(inconsistencies, columns=INCONSISTENCIAS_COLUMNS)
    if not inconsistencias_df.empty:
        inconsistencias_df = inconsistencias_df.sort_values(
            ["ID de persona", "Fecha", "Nombre", "Tipo de inconsistencia"], kind="stable"
        ).reset_index(drop=True)

    return diario_df, mensual_df, inconsistencias_df
