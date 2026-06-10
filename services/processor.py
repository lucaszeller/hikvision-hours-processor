from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Callable

import pandas as pd

from services.calculator import process_punches
from services.exceptions import (
    ExceptionConfigError,
    load_absences_template_file,
    load_exceptions_file,
    merge_exceptions,
    parse_manual_exceptions,
)
from services.exporter import export_report
from services.parser import load_hikvision_excel
from services.schedule_info import ScheduleInfoError, load_schedule_profiles


class ValidationError(Exception):
    pass


def _filter_to_report_month(source_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.Period | None]:
    if source_df.empty or "work_date_raw" not in source_df.columns:
        return source_df, None

    parsed = pd.to_datetime(source_df["work_date_raw"], errors="coerce")
    if parsed.dropna().empty:
        return source_df, None

    periods = parsed.dt.to_period("M")
    counts = periods.value_counts()
    if counts.empty:
        return source_df, None

    max_count = int(counts.max())
    candidates = [period for period, count in counts.items() if int(count) == max_count]
    selected_period = max(candidates)

    mask = periods == selected_period
    filtered = source_df.loc[mask.fillna(False)].copy()
    if filtered.empty:
        return source_df, None

    return filtered.reset_index(drop=True), selected_period


def _filter_exceptions_to_report_month(exceptions: list, period: pd.Period | None) -> list:
    if period is None:
        return list(exceptions)
    result = []
    for item in exceptions:
        item_period = pd.Timestamp(item.exception_date).to_period("M")
        if item_period == period:
            result.append(item)
    return result


def _hhmm_to_minutes(value: str) -> int:
    text = str(value).strip()
    if ":" not in text:
        raise ValidationError(f"Formato de horas invalido: '{value}'")

    hours_text, minutes_text = text.split(":", 1)
    try:
        hours = int(hours_text)
        minutes = int(minutes_text)
    except ValueError as exc:
        raise ValidationError(f"Formato de horas invalido: '{value}'") from exc

    if hours < 0 or minutes < 0 or minutes >= 60:
        raise ValidationError(f"Formato de horas invalido: '{value}'")

    return hours * 60 + minutes


def _validate_daily_consistency(daily_df: pd.DataFrame) -> None:
    if daily_df.empty:
        return

    required = {
        "ID de persona",
        "Nombre",
        "Fecha",
        "Estado",
        "Minutos reales",
        "Minutos redondeados",
        "Minutos extra",
        "Horas extra",
        "Horas totales",
    }
    missing = sorted(required - set(daily_df.columns))
    if missing:
        raise ValidationError(
            "Faltan columnas en hoja Diario para validacion: " + ", ".join(missing)
        )

    normalized = daily_df.copy()
    normalized["ID _key"] = normalized["ID de persona"].astype(str).str.strip()
    normalized["Fecha _key"] = pd.to_datetime(normalized["Fecha"], errors="coerce")

    invalid_dates = normalized[normalized["Fecha _key"].isna()]
    if not invalid_dates.empty:
        row_number = int(invalid_dates.index[0]) + 2
        raise ValidationError(
            f"Inconsistencia en Diario fila {row_number}: Fecha invalida '{invalid_dates.iloc[0]['Fecha']}'."
        )

    duplicate_mask = normalized.duplicated(subset=["ID _key", "Fecha _key"], keep=False)
    if duplicate_mask.any():
        duplicate_row = normalized.loc[duplicate_mask].iloc[0]
        row_number = int(duplicate_row.name) + 2
        raise ValidationError(
            "Inconsistencia en Diario: filas duplicadas por empleado/fecha. "
            f"Primera fila detectada: {row_number} "
            f"(ID={duplicate_row['ID de persona']}, Fecha={pd.Timestamp(duplicate_row['Fecha _key']).date()})."
        )

    for idx, row in daily_df.iterrows():
        employee_id = str(row["ID de persona"]).strip()
        employee_name = str(row["Nombre"]).strip()
        status = str(row["Estado"]).strip()
        segments = str(row["Tramos trabajados"]).strip()

        if employee_id == "":
            raise ValidationError(f"Inconsistencia en Diario fila {idx + 2}: ID de persona vacio.")
        if employee_name == "":
            raise ValidationError(f"Inconsistencia en Diario fila {idx + 2}: Nombre vacio.")

        real_minutes = int(row["Minutos reales"])
        rounded_minutes = int(row["Minutos redondeados"])
        extra_minutes = int(row["Minutos extra"])

        if real_minutes < 0 or rounded_minutes < 0 or extra_minutes < 0:
            raise ValidationError(
                "Inconsistencia en Diario fila "
                f"{idx + 2}: minutos negativos detectados "
                f"(reales={real_minutes}, redondeados={rounded_minutes}, extra={extra_minutes})."
            )

        if rounded_minutes % 30 != 0:
            raise ValidationError(
                f"Inconsistencia en Diario fila {idx + 2}: Minutos redondeados={rounded_minutes} "
                "no es multiplo de 30."
            )

        if extra_minutes > rounded_minutes:
            raise ValidationError(
                f"Inconsistencia en Diario fila {idx + 2}: Minutos extra={extra_minutes} "
                f"no puede superar Minutos redondeados={rounded_minutes}."
            )

        if rounded_minutes == 0 and segments != "":
            raise ValidationError(
                "Inconsistencia en Diario fila "
                f"{idx + 2}: hay tramos trabajados pero Minutos redondeados=0."
            )

        status_key = status.lower()
        if status_key in {"ausente", "ausencia"} and (rounded_minutes > 0 or extra_minutes > 0):
            raise ValidationError(
                f"Inconsistencia en Diario fila {idx + 2}: Estado='{status}' no puede tener minutos trabajados."
            )
        if status_key == "tarde" and rounded_minutes <= 0:
            raise ValidationError(
                f"Inconsistencia en Diario fila {idx + 2}: Estado='Tarde' requiere minutos trabajados."
            )

        hhmm_minutes = _hhmm_to_minutes(row["Horas totales"])
        if rounded_minutes != hhmm_minutes:
            raise ValidationError(
                "Inconsistencia en Diario fila "
                f"{idx + 2}: Horas totales='{row['Horas totales']}' no coincide "
                f"con Minutos redondeados={rounded_minutes}."
            )
        extra_hhmm = _hhmm_to_minutes(row["Horas extra"])
        if extra_minutes != extra_hhmm:
            raise ValidationError(
                "Inconsistencia en Diario fila "
                f"{idx + 2}: Horas extra='{row['Horas extra']}' no coincide "
                f"con Minutos extra={extra_minutes}."
            )


def _validate_monthly_consistency(daily_df: pd.DataFrame, monthly_df: pd.DataFrame) -> None:
    if daily_df.empty and monthly_df.empty:
        return

    if daily_df.empty and not monthly_df.empty:
        raise ValidationError("Mensual tiene datos pero Diario esta vacio.")

    if not daily_df.empty and monthly_df.empty:
        raise ValidationError("Diario tiene datos pero Mensual esta vacio.")

    required_daily = {"ID de persona", "Nombre", "Minutos redondeados"}
    required_monthly = {
        "ID de persona",
        "Nombre",
        "Minutos totales",
        "Horas totales",
        "Dias trabajados",
        "Minutos extra",
        "Horas extra",
    }
    missing_daily = sorted(required_daily - set(daily_df.columns))
    missing_monthly = sorted(required_monthly - set(monthly_df.columns))

    if missing_daily:
        raise ValidationError(
            "Faltan columnas en Diario para validacion mensual: " + ", ".join(missing_daily)
        )
    if missing_monthly:
        raise ValidationError(
            "Faltan columnas en Mensual para validacion: " + ", ".join(missing_monthly)
        )

    monthly_normalized = monthly_df.copy()
    monthly_normalized["ID _key"] = monthly_normalized["ID de persona"].astype(str).str.strip()
    monthly_duplicates = monthly_normalized.duplicated(subset=["ID _key", "Nombre"], keep=False)
    if monthly_duplicates.any():
        duplicate_row = monthly_normalized.loc[monthly_duplicates].iloc[0]
        raise ValidationError(
            "Inconsistencia en Mensual: filas duplicadas por empleado. "
            f"ID={duplicate_row['ID de persona']} Nombre={duplicate_row['Nombre']}."
        )

    expected_monthly = (
        daily_df.groupby(["ID de persona", "Nombre"], as_index=False)
        .agg({"Fecha": "nunique", "Minutos redondeados": "sum", "Minutos extra": "sum"})
        .rename(
            columns={
                "Fecha": "Dias esperados",
                "Minutos redondeados": "Minutos esperados",
                "Minutos extra": "Minutos extra esperados",
            }
        )
    )

    merged = monthly_df.merge(
        expected_monthly,
        on=["ID de persona", "Nombre"],
        how="outer",
        indicator=True,
    )

    if not merged[merged["_merge"] != "both"].empty:
        raise ValidationError("Mensual no coincide con Diario (empleados faltantes o extra).")

    for _, row in merged.iterrows():
        expected_minutes = int(row["Minutos esperados"])
        monthly_minutes = int(row["Minutos totales"])
        monthly_extra = int(row["Minutos extra"])
        expected_extra = int(row["Minutos extra esperados"])
        monthly_days = int(row["Dias trabajados"])
        expected_days = int(row["Dias esperados"])

        if monthly_minutes < 0 or monthly_extra < 0:
            raise ValidationError(
                "Minutos negativos en Mensual para "
                f"ID {row['ID de persona']} - {row['Nombre']}."
            )
        if monthly_days < 0:
            raise ValidationError(
                "Dias trabajados negativos en Mensual para "
                f"ID {row['ID de persona']} - {row['Nombre']}."
            )
        if monthly_extra > monthly_minutes:
            raise ValidationError(
                "Horas extra inconsistentes en Mensual para "
                f"ID {row['ID de persona']} - {row['Nombre']}: "
                f"minutos extra={monthly_extra}, minutos totales={monthly_minutes}."
            )
        if monthly_minutes % 30 != 0 or monthly_extra % 30 != 0:
            raise ValidationError(
                "Minutos no redondeados en Mensual para "
                f"ID {row['ID de persona']} - {row['Nombre']}."
            )

        if expected_minutes != monthly_minutes:
            raise ValidationError(
                "Inconsistencia mensual para "
                f"ID {row['ID de persona']} - {row['Nombre']}: "
                f"mensual={monthly_minutes}, esperado={expected_minutes}."
            )
        if expected_extra != monthly_extra:
            raise ValidationError(
                "Inconsistencia de horas extra mensual para "
                f"ID {row['ID de persona']} - {row['Nombre']}: "
                f"mensual={monthly_extra}, esperado={expected_extra}."
            )

        if expected_days != monthly_days:
            raise ValidationError(
                "Dias trabajados inconsistentes para "
                f"ID {row['ID de persona']} - {row['Nombre']}: "
                f"mensual={monthly_days}, esperado={expected_days}."
            )

        hhmm_minutes = _hhmm_to_minutes(row["Horas totales"])
        if hhmm_minutes != monthly_minutes:
            raise ValidationError(
                "Formato de horas mensual inconsistente para "
                f"ID {row['ID de persona']} - {row['Nombre']}: "
                f"Horas totales='{row['Horas totales']}' y minutos={monthly_minutes}."
            )
        extra_minutes = int(row["Minutos extra"])
        extra_hhmm = _hhmm_to_minutes(row["Horas extra"])
        if extra_minutes != extra_hhmm:
            raise ValidationError(
                "Formato de horas extra mensual inconsistente para "
                f"ID {row['ID de persona']} - {row['Nombre']}: "
                f"Horas extra='{row['Horas extra']}' y minutos extra={extra_minutes}."
            )


def _validate_results(daily_df: pd.DataFrame, monthly_df: pd.DataFrame) -> None:
    _validate_daily_consistency(daily_df)
    _validate_monthly_consistency(daily_df, monthly_df)


def _validate_no_inconsistencies(inconsistencies_df: pd.DataFrame) -> None:
    if inconsistencies_df.empty:
        return

    counts = (
        inconsistencies_df["Tipo de inconsistencia"].value_counts().sort_index().to_dict()
    )
    summary = ", ".join(f"{issue}: {count}" for issue, count in counts.items())

    examples = []
    for _, row in inconsistencies_df.head(5).iterrows():
        examples.append(
            f"{row['ID de persona']} | {row['Nombre']} | {row['Fecha']} | "
            f"{row['Tipo de inconsistencia']} | {row['Detalle']}"
        )

    raise ValidationError(
        "Se detectaron inconsistencias de fichadas. "
        "Modo estricto activo: no se genera reporte.\n"
        f"Resumen: {summary}\n"
        "Primeros casos:\n"
        + "\n".join(examples)
    )


class ProcessorService:
    def process_file(
        self,
        input_path: str | Path,
        output_dir: str | Path | None = None,
        progress_callback: Callable[[int, str], None] | None = None,
        strict_mode: bool = False,
        exceptions_file: str | Path | None = None,
        manual_exceptions_text: str | None = None,
    ) -> Path:
        source = Path(input_path)
        target_dir = Path(output_dir) if output_dir is not None else source.parent

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = target_dir / f"reporte_horas_{stamp}.xlsx"

        if progress_callback:
            progress_callback(10, "Leyendo reporte Hikvision...")
        source_df = load_hikvision_excel(source)
        source_df, selected_period = _filter_to_report_month(source_df)
        if progress_callback and selected_period is not None:
            progress_callback(
                14,
                f"Filtrando reporte al mes {selected_period.strftime('%Y-%m')} (mes predominante).",
            )

        file_exceptions: list = []
        manual_exceptions: list = []
        template_absences: list = []

        try:
            if exceptions_file is not None:
                if progress_callback:
                    progress_callback(25, "Cargando excepciones desde archivo...")
                file_exceptions = load_exceptions_file(exceptions_file)

            if manual_exceptions_text and manual_exceptions_text.strip():
                if progress_callback:
                    progress_callback(35, "Cargando excepciones manuales...")
                manual_exceptions = parse_manual_exceptions(manual_exceptions_text)

            template_candidates = [Path("date.xlsx"), source.parent / "date.xlsx"]
            for candidate in template_candidates:
                if candidate.exists():
                    if progress_callback:
                        progress_callback(40, f"Cargando ausencias desde {candidate.name}...")
                    template_absences = load_absences_template_file(candidate)
                    break
        except ExceptionConfigError as exc:
            raise ValidationError(str(exc)) from exc

        merged_exceptions = merge_exceptions(
            merge_exceptions(file_exceptions, manual_exceptions),
            template_absences,
        )
        merged_exceptions = _filter_exceptions_to_report_month(merged_exceptions, selected_period)

        info_path_candidates = [
            Path("date.xlsx"),
            source.parent / "date.xlsx",
            Path("info.xlsx"),
            source.parent / "info.xlsx",
        ]
        schedule_minutes: dict[str, int] = {}
        start_minutes: dict[str, int] = {}
        working_weekdays_by_employee: dict[str, set[int]] = {}
        split_schedule_by_employee: dict[str, dict[str, int | bool | None]] = {}
        flexible_attendance_employee_ids: set[str] = set()
        for candidate in info_path_candidates:
            if candidate.exists():
                try:
                    if progress_callback:
                        progress_callback(45, f"Cargando horarios desde {candidate.name}...")
                    profiles = load_schedule_profiles(candidate)
                    schedule_minutes = {
                        employee_id: values["scheduled_minutes"] for employee_id, values in profiles.items()
                    }
                    start_minutes = {
                        employee_id: values["start_minute"] for employee_id, values in profiles.items()
                    }
                    working_weekdays_by_employee = {
                        employee_id: set(values.get("working_weekdays", {0, 1, 2, 3, 4}))
                        for employee_id, values in profiles.items()
                    }
                    split_schedule_by_employee = {
                        employee_id: {
                            "is_continuous": bool(values.get("is_continuous", False)),
                            "morning_in_minute": values.get("morning_in_minute"),
                            "morning_out_minute": values.get("morning_out_minute"),
                            "afternoon_in_minute": values.get("afternoon_in_minute"),
                            "afternoon_out_minute": values.get("afternoon_out_minute"),
                        }
                        for employee_id, values in profiles.items()
                    }
                    flexible_attendance_employee_ids = {
                        str(employee_id).strip()
                        for employee_id, values in profiles.items()
                        if bool(values.get("flexible_attendance", False))
                    }
                except ScheduleInfoError as exc:
                    raise ValidationError(str(exc)) from exc
                break

        if progress_callback:
            progress_callback(55, "Calculando horas y detectando inconsistencias...")
        daily_df, monthly_df, inconsistencies_df = process_punches(
            source_df,
            exceptions=merged_exceptions,
            scheduled_minutes_by_employee=schedule_minutes,
            scheduled_start_minute_by_employee=start_minutes,
            working_weekdays_by_employee=working_weekdays_by_employee,
            split_schedule_by_employee=split_schedule_by_employee,
            flexible_attendance_employee_ids=flexible_attendance_employee_ids,
        )

        if progress_callback:
            progress_callback(72, "Validando consistencia de resultados...")
        try:
            _validate_results(daily_df, monthly_df)
        except ValidationError as exc:
            if strict_mode:
                raise
            validation_issue = pd.DataFrame(
                [
                    {
                        "ID de persona": "",
                        "Nombre": "",
                        "Fecha": "",
                        "Tipo de inconsistencia": "Validacion de reporte",
                        "Detalle": str(exc),
                    }
                ]
            )
            inconsistencies_df = pd.concat(
                [inconsistencies_df, validation_issue],
                ignore_index=True,
            )
            if progress_callback:
                progress_callback(
                    76,
                    "Se detectaron inconsistencias de validacion; se continua y se incluyen en el reporte.",
                )

        if strict_mode:
            _validate_no_inconsistencies(inconsistencies_df)
        elif progress_callback and not inconsistencies_df.empty:
            counts = (
                inconsistencies_df["Tipo de inconsistencia"]
                .value_counts()
                .sort_index()
                .to_dict()
            )
            summary = ", ".join(f"{issue}: {count}" for issue, count in counts.items())
            progress_callback(78, f"Inconsistencias detectadas (incluidas en reporte): {summary}")

        if progress_callback:
            progress_callback(85, "Generando Excel final...")
        report_path = export_report(output_path, daily_df, monthly_df, inconsistencies_df)

        if progress_callback:
            progress_callback(100, "Proceso completado.")
        return report_path
