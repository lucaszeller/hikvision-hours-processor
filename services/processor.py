from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Callable

import pandas as pd

from services.calculator import process_punches
from services.exporter import export_report
from services.parser import load_hikvision_excel


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
        "Minutos reales",
        "Minutos redondeados",
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
    ) -> Path:
        source = Path(input_path)
        target_dir = Path(output_dir) if output_dir is not None else source.parent

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = target_dir / f"reporte_horas_{stamp}.xlsx"

        if progress_callback:
            progress_callback(10, "Leyendo reporte Hikvision...")
        source_df = load_hikvision_excel(source)

        if progress_callback:
            progress_callback(55, "Calculando horas y detectando inconsistencias...")
        daily_df, monthly_df, inconsistencies_df = process_punches(source_df)

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
