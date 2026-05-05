from __future__ import annotations

from datetime import date, datetime

import pandas as pd

from services.exceptions import WorkException, build_exception_index, find_matching_exceptions

DIARIO_COLUMNS = [
    "ID de persona",
    "Nombre",
    "Fecha",
    "Departamento",
    "Tramos trabajados",
    "Minutos reales",
    "Minutos redondeados",
    "Horas totales",
]

MENSUAL_COLUMNS = [
    "ID de persona",
    "Nombre",
    "Dias trabajados",
    "Minutos totales",
    "Horas totales",
]

INCONSISTENCIAS_COLUMNS = [
    "ID de persona",
    "Nombre",
    "Fecha",
    "Tipo de inconsistencia",
    "Detalle",
]


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
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    valid_segments: list[dict] = []
    inconsistencies: list[dict] = []

    exceptions = exceptions or []
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
                    "Ambos vacios",
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

    if segments_df.empty:
        diario_df = pd.DataFrame(columns=DIARIO_COLUMNS)
        mensual_df = pd.DataFrame(columns=MENSUAL_COLUMNS)
    else:
        segments_df = segments_df.sort_values(
            ["employee_id", "work_date", "entry_dt", "schedule"],
            kind="stable",
        ).reset_index(drop=True)

        grouped_rows: list[dict] = []
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

            grouped_rows.append(
                {
                    "ID de persona": employee_id,
                    "Nombre": employee_name,
                    "Fecha": work_date,
                    "Departamento": department_value,
                    "Tramos trabajados": " | ".join(segments_text),
                    "Minutos reales": real_total,
                    "Minutos redondeados": rounded_total,
                    "Horas totales": _minutes_to_hhmm(rounded_total),
                }
            )

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
                }
            )
            .rename(
                columns={
                    "Fecha": "Dias trabajados",
                    "Minutos redondeados": "Minutos totales",
                }
            )
        )
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
