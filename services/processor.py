from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Callable

import pandas as pd

from services.calculator import process_punches
from services.exceptions import (
    ExceptionConfigError,
    load_exceptions_file,
    merge_exceptions,
    parse_manual_exceptions,
)
from services.exporter import export_report
from services.parser import load_hikvision_excel
from services.schedule_info import ScheduleInfoError, load_schedule_profiles


class ValidationError(Exception):
    pass


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

    for idx, row in daily_df.iterrows():
        rounded_minutes = int(row["Minutos redondeados"])
        hhmm_minutes = _hhmm_to_minutes(row["Horas totales"])
        if rounded_minutes != hhmm_minutes:
            raise ValidationError(
                "Inconsistencia en Diario fila "
                f"{idx + 2}: Horas totales='{row['Horas totales']}' no coincide "
                f"con Minutos redondeados={rounded_minutes}."
            )
        extra_minutes = int(row["Minutos extra"])
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

    expected_monthly = (
        daily_df.groupby(["ID de persona", "Nombre"], as_index=False)
        .agg({"Fecha": "nunique", "Minutos redondeados": "sum"})
        .rename(
            columns={
                "Fecha": "Dias esperados",
                "Minutos redondeados": "Minutos esperados",
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
        if expected_minutes != monthly_minutes:
            raise ValidationError(
                "Inconsistencia mensual para "
                f"ID {row['ID de persona']} - {row['Nombre']}: "
                f"mensual={monthly_minutes}, esperado={expected_minutes}."
            )

        expected_days = int(row["Dias esperados"])
        monthly_days = int(row["Dias trabajados"])
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

        file_exceptions = []
        manual_exceptions = []

        try:
            if exceptions_file is not None:
                if progress_callback:
                    progress_callback(25, "Cargando excepciones desde archivo...")
                file_exceptions = load_exceptions_file(exceptions_file)

            if manual_exceptions_text and manual_exceptions_text.strip():
                if progress_callback:
                    progress_callback(35, "Cargando excepciones manuales...")
                manual_exceptions = parse_manual_exceptions(manual_exceptions_text)
        except ExceptionConfigError as exc:
            raise ValidationError(str(exc)) from exc

        merged_exceptions = merge_exceptions(file_exceptions, manual_exceptions)

        info_path_candidates = [Path("info.xlsx"), source.parent / "info.xlsx"]
        schedule_minutes: dict[str, int] = {}
        start_minutes: dict[str, int] = {}
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
        )

        if progress_callback:
            progress_callback(72, "Validando consistencia de resultados...")
        _validate_results(daily_df, monthly_df)

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
