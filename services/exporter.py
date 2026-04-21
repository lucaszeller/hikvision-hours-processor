from __future__ import annotations

from pathlib import Path

import pandas as pd


class ExportError(Exception):
    pass


def export_report(output_path: str | Path, daily_df: pd.DataFrame, monthly_df: pd.DataFrame, inconsistencies_df: pd.DataFrame) -> Path:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    try:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            daily_df.to_excel(writer, sheet_name="Horas diarias", index=False)
            monthly_df.to_excel(writer, sheet_name="Resumen mensual", index=False)
            inconsistencies_df.to_excel(writer, sheet_name="Inconsistencias", index=False)
    except Exception as exc:
        raise ExportError(f"No se pudo exportar el Excel: {exc}") from exc

    return output
