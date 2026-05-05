from __future__ import annotations

from datetime import date, datetime

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


def process_punches(
    df: pd.DataFrame,
    exceptions: list[WorkException] | None = None,
    scheduled_minutes_by_employee: dict[str, int] | None = None,
    scheduled_start_minute_by_employee: dict[str, int] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    valid_segments: list[dict] = []
    inconsistencies: list[dict] = []
    day_flags: dict[tuple[str, str, date], dict[str, object]] = {}

    exceptions = exceptions or []
    scheduled_minutes_by_employee = scheduled_minutes_by_employee or {}
    scheduled_start_minute_by_employee = scheduled_start_minute_by_employee or {}
    has_schedule_reference = bool(scheduled_minutes_by_employee)
    exceptions_index = build_exception_index(exceptions)
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

        if employee_id == "":
            _inconsistency(
                inconsistencies,
                employee_id,
                employee_name,
                date_for_report,
                "ID de persona vacio",
                "La fila no contiene ID de persona.",
            )
        if employee_name == "":
            _inconsistency(
                inconsistencies,
                employee_id,
                employee_name,
                date_for_report,
                "Nombre vacio",
                "La fila no contiene nombre de empleado.",
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
                inconsistencies,
                employee_id,
                employee_name,
                date_for_report,
                "Fecha invalida",
                f"Valor recibido: '{work_date_raw}'.",
            )

        has_entry = not _is_blank(entry_raw)
        has_exit = not _is_blank(exit_raw)
        weekday = date_parsed.weekday() if date_parsed is not None else None
        is_saturday = weekday == 5
        is_sunday = weekday == 6
        is_weekend = is_saturday or is_sunday

        # Regla: lunes a viernes normales; fines de semana sólo si hay fichada.
        if is_weekend and not (has_entry or has_exit):
            continue

        late_marker = _marker_is_positive(row.get("late_raw", ""))
        absent_marker = _marker_is_positive(row.get("absent_raw", ""))

        if employee_id != "" and employee_name != "" and date_parsed is not None:
            day_key = (employee_id, employee_name, date_parsed.date())
            day_state = day_flags.setdefault(
                day_key,
                {"late": False, "absent": False, "departments": set(), "exception_types": []},
            )
            if department:
                day_state["departments"].add(department)
            day_state["late"] = bool(day_state["late"]) or late_marker
            missing_both_without_exception = not has_entry and not has_exit and not row_exceptions
            day_state["absent"] = bool(day_state["absent"]) or absent_marker or missing_both_without_exception
            for exc in row_exceptions:
                exc_type = str(exc.exception_type).strip()
                if exc_type and exc_type not in day_state["exception_types"]:
                    day_state["exception_types"].append(exc_type)

        if not has_entry and not has_exit:
            if row_exceptions:
                _inconsistency(
                    inconsistencies,
                    employee_id,
                    employee_name,
                    date_for_report,
                    "Excepcion aplicada",
                    "Se omitio inconsistencia de ausencia sin fichadas por excepcion: "
                    + _exception_summary(row_exceptions),
                )
            else:
                _inconsistency(
                    inconsistencies,
                    employee_id,
                    employee_name,
                    date_for_report,
                    "Ausente",
                    "Faltan Registro de entrada y Registro de salida.",
                )
        elif not has_entry:
            if row_exceptions:
                _inconsistency(
                    inconsistencies,
                    employee_id,
                    employee_name,
                    date_for_report,
                    "Excepcion aplicada",
                    "Se omitio inconsistencia de falta de entrada por excepcion: "
                    + _exception_summary(row_exceptions),
                )
            else:
                _inconsistency(
                    inconsistencies,
                    employee_id,
                    employee_name,
                    date_for_report,
                    "Falta Registro de entrada",
                    "No se encontro hora de entrada.",
                )
        elif not has_exit:
            if row_exceptions:
                _inconsistency(
                    inconsistencies,
                    employee_id,
                    employee_name,
                    date_for_report,
                    "Excepcion aplicada",
                    "Se omitio inconsistencia de falta de salida por excepcion: "
                    + _exception_summary(row_exceptions),
                )
            else:
                _inconsistency(
                    inconsistencies,
                    employee_id,
                    employee_name,
                    date_for_report,
                    "Falta Registro de salida",
                    "No se encontro hora de salida.",
                )

        entry_time = _parse_time(entry_raw) if has_entry else None
        exit_time = _parse_time(exit_raw) if has_exit else None

        if has_entry and entry_time is None:
            _inconsistency(
                inconsistencies,
                employee_id,
                employee_name,
                date_for_report,
                "Hora de entrada invalida",
                f"Valor recibido: '{entry_raw}'.",
            )
        if has_exit and exit_time is None:
            _inconsistency(
                inconsistencies,
                employee_id,
                employee_name,
                date_for_report,
                "Hora de salida invalida",
                f"Valor recibido: '{exit_raw}'.",
            )

        if (
            employee_id == ""
            or employee_name == ""
            or date_parsed is None
            or entry_time is None
            or exit_time is None
        ):
            continue

        start_dt = datetime.combine(date_parsed.date(), entry_time.time())
        end_dt = datetime.combine(date_parsed.date(), exit_time.time())

        if end_dt <= start_dt:
            if row_exceptions:
                _inconsistency(
                    inconsistencies,
                    employee_id,
                    employee_name,
                    date_parsed.date(),
                    "Excepcion aplicada",
                    "Se omitio inconsistencia de salida menor/igual que entrada por excepcion: "
                    + _exception_summary(row_exceptions),
                )
            else:
                _inconsistency(
                    inconsistencies,
                    employee_id,
                    employee_name,
                    date_parsed.date(),
                    "Salida menor que entrada",
                    f"Entrada {entry_raw} / Salida {exit_raw}.",
                )
            continue

        real_minutes = int(round((end_dt - start_dt).total_seconds() / 60))
        rounded_minutes = _round_to_30(real_minutes)

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

    for item in exceptions:
        key = (item.employee_id, item.exception_date, item.exception_type.strip(), item.details.strip())
        if key in used_exception_keys:
            continue
        _inconsistency(
            inconsistencies,
            item.employee_id or "",
            "",
            item.exception_date,
            "Excepcion configurada sin uso",
            f"{item.exception_type}. {item.details}".strip(),
        )

    segments_df = pd.DataFrame(valid_segments)

    if segments_df.empty and not day_flags:
        diario_df = pd.DataFrame(columns=DIARIO_COLUMNS)
        mensual_df = pd.DataFrame(columns=MENSUAL_COLUMNS)
    else:
        if segments_df.empty:
            segments_df = pd.DataFrame(
                columns=[
                    "employee_id",
                    "employee_name",
                    "department",
                    "work_date",
                    "entry_dt",
                    "exit_dt",
                    "schedule",
                    "real_minutes",
                    "rounded_minutes",
                ]
            )

        segments_df = segments_df.sort_values(
            ["employee_id", "work_date", "entry_dt", "schedule"],
            kind="stable",
        ).reset_index(drop=True)

        day_rows: dict[tuple[str, str, date], dict] = {}
        day_first_entry_minutes: dict[tuple[str, str, date], int] = {}
        for (employee_id, employee_name, work_date), group in segments_df.groupby(
            ["employee_id", "employee_name", "work_date"],
            sort=True,
        ):
            departments = [d for d in group["department"].astype(str).str.strip().unique() if d]
            department_value = " / ".join(departments)

            segments_text: list[str] = []
            for _, segment in group.iterrows():
                segment_line = (
                    f"{segment['entry_dt'].strftime('%H:%M')} - {segment['exit_dt'].strftime('%H:%M')}"
                )
                if segment["schedule"]:
                    segment_line += f" ({segment['schedule']})"
                segments_text.append(segment_line)

            real_total = int(group["real_minutes"].sum())
            rounded_total = int(group["rounded_minutes"].sum())
            first_entry = group["entry_dt"].min()
            first_entry_minutes = first_entry.hour * 60 + first_entry.minute
            day_first_entry_minutes[(str(employee_id), str(employee_name), work_date)] = first_entry_minutes
            if work_date.weekday() in {5, 6}:
                overtime_minutes = rounded_total
            else:
                scheduled_minutes = scheduled_minutes_by_employee.get(str(employee_id).strip())
                if scheduled_minutes is None:
                    overtime_minutes = 0
                    if has_schedule_reference:
                        _inconsistency(
                            inconsistencies,
                            str(employee_id),
                            str(employee_name),
                            work_date,
                            "Horario no definido",
                            "No hay horario configurado en info.xlsx para calcular horas extra.",
                        )
                else:
                    overtime_minutes = max(0, rounded_total - int(scheduled_minutes))

            day_rows[(str(employee_id), str(employee_name), work_date)] = {
                "ID de persona": employee_id,
                "Nombre": employee_name,
                "Fecha": work_date,
                "Departamento": department_value,
                "Tramos trabajados": " | ".join(segments_text),
                "Minutos reales": real_total,
                "Minutos redondeados": rounded_total,
                "Minutos extra": overtime_minutes,
                "Horas extra": _minutes_to_hhmm(overtime_minutes),
                "Horas totales": _minutes_to_hhmm(rounded_total),
            }

        grouped_rows: list[dict] = []
        all_day_keys = sorted(
            set(day_rows.keys()) | set(day_flags.keys()),
            key=lambda k: (k[0], k[2], k[1]),
        )
        for day_key in all_day_keys:
            employee_id, employee_name, work_date = day_key
            state = day_flags.get(
                day_key,
                {"late": False, "absent": False, "departments": set(), "exception_types": []},
            )
            row_data = day_rows.get(day_key)

            if row_data is None:
                if (
                    not bool(state.get("absent"))
                    and not bool(state.get("late"))
                    and not state.get("exception_types")
                ):
                    continue
                departments = sorted([d for d in state["departments"] if str(d).strip() != ""])
                row_data = {
                    "ID de persona": employee_id,
                    "Nombre": employee_name,
                    "Fecha": work_date,
                    "Departamento": " / ".join(departments),
                    "Tramos trabajados": "",
                    "Minutos reales": 0,
                    "Minutos redondeados": 0,
                    "Minutos extra": 0,
                    "Horas extra": "00:00",
                    "Horas totales": "00:00",
                }

            worked_minutes = int(row_data.get("Minutos redondeados", 0))
            if worked_minutes > 0:
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
                exception_types = state.get("exception_types") or []
                if exception_types:
                    status = str(exception_types[0]).strip().title()
                elif bool(state.get("absent")):
                    status = "Ausente"
                elif bool(state.get("late")):
                    status = "Tarde"
                else:
                    status = "Normal"

            row_data["Estado"] = status
            grouped_rows.append(row_data)

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
            ["ID de persona", "Fecha", "Nombre", "Tipo de inconsistencia"],
            kind="stable",
        ).reset_index(drop=True)

    return diario_df, mensual_df, inconsistencias_df
