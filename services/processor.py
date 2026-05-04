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
    if minutes < 0 or minutes >= 60 or hours < 0:
        raise ValidationError(f"Formato de horas invalido: '{value}'")
    return hours * 60 + minutes


def _validate_daily_consistency(daily_df: pd.DataFrame) -> None:
    if daily_df.empty:
        return

    required = {"Legajo", "Nombre", "Fecha", "Horas totales", "Minutos totales"}
    missing = sorted(required - set(daily_df.columns))
    if missing:
        raise ValidationError(
            "Faltan columnas en horas diarias para validacion: " + ", ".join(missing)
        )

    for idx, row in daily_df.iterrows():
        reported_minutes = int(row["Minutos totales"])
        parsed_minutes = _hhmm_to_minutes(row["Horas totales"])
        if reported_minutes != parsed_minutes:
            raise ValidationError(
                "Inconsistencia diaria detectada en fila "
                f"{idx + 2}: Horas totales='{row['Horas totales']}' "
                f"no coincide con Minutos totales={reported_minutes}."
            )


def _validate_monthly_consistency(daily_df: pd.DataFrame, monthly_df: pd.DataFrame) -> None:
    if daily_df.empty and monthly_df.empty:
        return

    if daily_df.empty and not monthly_df.empty:
        raise ValidationError("Resumen mensual tiene datos pero horas diarias esta vacio.")

    if not daily_df.empty and monthly_df.empty:
        raise ValidationError("Horas diarias tiene datos pero resumen mensual esta vacio.")

    required_daily = {"Legajo", "Nombre", "Minutos totales"}
    required_monthly = {"Legajo", "Nombre", "Minutos totales", "Horas mensuales"}
    missing_daily = sorted(required_daily - set(daily_df.columns))
    missing_monthly = sorted(required_monthly - set(monthly_df.columns))

    if missing_daily:
        raise ValidationError(
            "Faltan columnas en horas diarias para validacion mensual: "
            + ", ".join(missing_daily)
        )
    if missing_monthly:
        raise ValidationError(
            "Faltan columnas en resumen mensual para validacion: "
            + ", ".join(missing_monthly)
        )

    expected_monthly = (
        daily_df.groupby(["Legajo", "Nombre"], as_index=False)["Minutos totales"]
        .sum()
        .rename(columns={"Minutos totales": "Minutos esperados"})
    )

    merged = monthly_df.merge(
        expected_monthly,
        on=["Legajo", "Nombre"],
        how="outer",
        indicator=True,
    )

    missing_rows = merged[merged["_merge"] != "both"]
    if not missing_rows.empty:
        raise ValidationError(
            "Resumen mensual no coincide con horas diarias (empleados faltantes o extra)."
        )

    for idx, row in merged.iterrows():
        monthly_minutes = int(row["Minutos totales"])
        expected_minutes = int(row["Minutos esperados"])
        if monthly_minutes != expected_minutes:
            raise ValidationError(
                "Inconsistencia mensual detectada para "
                f"Legajo {row['Legajo']} - {row['Nombre']}: "
                f"mensual={monthly_minutes}, esperado={expected_minutes}."
            )

        parsed_hhmm = _hhmm_to_minutes(row["Horas mensuales"])
        if parsed_hhmm != monthly_minutes:
            raise ValidationError(
                "Formato mensual inconsistente para "
                f"Legajo {row['Legajo']} - {row['Nombre']}: "
                f"Horas mensuales='{row['Horas mensuales']}' no coincide con minutos={monthly_minutes}."
            )


def _validate_results(daily_df: pd.DataFrame, monthly_df: pd.DataFrame) -> None:
    _validate_daily_consistency(daily_df)
    _validate_monthly_consistency(daily_df, monthly_df)


def _validate_no_inconsistencies(inconsistencies_df: pd.DataFrame) -> None:
    if inconsistencies_df.empty:
        return

    issue_counts = (
        inconsistencies_df["Tipo de inconsistencia"]
        .value_counts()
        .sort_index()
        .to_dict()
    )
    summary = ", ".join(f"{issue}: {count}" for issue, count in issue_counts.items())

    examples = []
    for _, row in inconsistencies_df.head(5).iterrows():
        examples.append(
            f"{row['Legajo']} | {row['Nombre']} | {row['Fecha']} | "
            f"{row['Tipo de inconsistencia']} | {row['Detalle']}"
        )

    details = "\n".join(examples)
    raise ValidationError(
        "Se detectaron inconsistencias de fichadas. "
        "Modo estricto activo: no se genera reporte con datos dudosos.\n"
        f"Resumen: {summary}\n"
        f"Primeros casos:\n{details}"
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
        if output_dir is None:
            output_dir = source.parent

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = Path(output_dir) / f"reporte_horas_{stamp}.xlsx"

        if progress_callback:
            progress_callback(10, "Leyendo fichadas del archivo...")
        punches_df = load_hikvision_excel(source)

        if progress_callback:
            progress_callback(55, "Calculando horas e inconsistencias...")
        daily_df, monthly_df, inconsistencies_df = process_punches(punches_df)

        if progress_callback:
            progress_callback(72, "Validando consistencia de resultados...")
        _validate_results(daily_df, monthly_df)
        if strict_mode:
            _validate_no_inconsistencies(inconsistencies_df)
        elif progress_callback and not inconsistencies_df.empty:
            issue_counts = (
                inconsistencies_df["Tipo de inconsistencia"]
                .value_counts()
                .sort_index()
                .to_dict()
            )
            summary = ", ".join(f"{issue}: {count}" for issue, count in issue_counts.items())
            progress_callback(78, f"Se detectaron inconsistencias (se incluyen en reporte): {summary}")

        if progress_callback:
            progress_callback(85, "Generando reporte final...")
        report_path = export_report(output_path, daily_df, monthly_df, inconsistencies_df)

        if progress_callback:
            progress_callback(100, "Proceso completado.")
        return report_path
