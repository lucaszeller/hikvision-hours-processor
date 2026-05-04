from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


class ExportError(Exception):
    pass


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
ALT_ROW_FILL = PatternFill("solid", fgColor="F5F8FC")
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)

SHEET_LAYOUTS = {
    "Horas diarias": {
        "widths": {
            "Legajo": 12,
            "Nombre": 28,
            "Fecha": 12,
            "Cantidad de fichadas": 20,
            "Marcaciones (I/S)": 44,
            "Tramos trabajados": 56,
            "Horas totales": 14,
            "Minutos totales": 16,
        },
        "left_cols": {"Nombre", "Marcaciones (I/S)", "Tramos trabajados"},
        "wrap_cols": {"Marcaciones (I/S)", "Tramos trabajados"},
        "date_cols": {"Fecha"},
        "priority_sort": ["Fecha", "Nombre", "Legajo"],
    },
    "Resumen mensual": {
        "widths": {
            "Legajo": 12,
            "Nombre": 28,
            "Minutos totales": 16,
            "Horas mensuales": 14,
        },
        "left_cols": {"Nombre"},
        "wrap_cols": set(),
        "date_cols": set(),
        "priority_sort": ["Nombre", "Legajo"],
    },
    "Inconsistencias": {
        "widths": {
            "Legajo": 12,
            "Nombre": 28,
            "Fecha": 12,
            "Tipo de inconsistencia": 24,
            "Detalle": 64,
        },
        "left_cols": {"Nombre", "Tipo de inconsistencia", "Detalle"},
        "wrap_cols": {"Detalle"},
        "date_cols": {"Fecha"},
        "priority_sort": ["Fecha", "Nombre", "Legajo"],
    },
}


def _sort_for_report(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    available = [col for col in columns if col in df.columns]
    if not available or df.empty:
        return df
    return df.sort_values(available, kind="stable").reset_index(drop=True)


def _apply_sheet_format(worksheet, sheet_name: str) -> None:
    config = SHEET_LAYOUTS[sheet_name]
    headers = [cell.value for cell in worksheet[1]]
    header_to_idx = {str(value): idx + 1 for idx, value in enumerate(headers) if value is not None}

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for cell in worksheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER

    for header, width in config["widths"].items():
        col_idx = header_to_idx.get(header)
        if col_idx:
            worksheet.column_dimensions[get_column_letter(col_idx)].width = width

    for row_idx in range(2, worksheet.max_row + 1):
        if row_idx % 2 == 0:
            for col_idx in range(1, worksheet.max_column + 1):
                worksheet.cell(row=row_idx, column=col_idx).fill = ALT_ROW_FILL

        for header, col_idx in header_to_idx.items():
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.border = THIN_BORDER

            horizontal = "left" if header in config["left_cols"] else "center"
            wrap = header in config["wrap_cols"]
            cell.alignment = Alignment(horizontal=horizontal, vertical="top", wrap_text=wrap)

            if header in config["date_cols"] and cell.value not in ("", None):
                cell.number_format = "DD/MM/YYYY"

def export_report(
    output_path: str | Path,
    daily_df: pd.DataFrame,
    monthly_df: pd.DataFrame,
    inconsistencies_df: pd.DataFrame,
) -> Path:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    temp_output = output.with_suffix(f"{output.suffix}.tmp")

    daily_export = _sort_for_report(daily_df.copy(), SHEET_LAYOUTS["Horas diarias"]["priority_sort"])
    monthly_export = _sort_for_report(monthly_df.copy(), SHEET_LAYOUTS["Resumen mensual"]["priority_sort"])
    inconsistencies_export = _sort_for_report(
        inconsistencies_df.copy(), SHEET_LAYOUTS["Inconsistencias"]["priority_sort"]
    )

    try:
        with pd.ExcelWriter(temp_output, engine="openpyxl") as writer:
            daily_export.to_excel(writer, sheet_name="Horas diarias", index=False)
            monthly_export.to_excel(writer, sheet_name="Resumen mensual", index=False)
            inconsistencies_export.to_excel(writer, sheet_name="Inconsistencias", index=False)

            workbook = writer.book
            for sheet_name in ("Horas diarias", "Resumen mensual", "Inconsistencias"):
                worksheet = workbook[sheet_name]
                _apply_sheet_format(worksheet, sheet_name)

        temp_output.replace(output)
    except Exception as exc:
        if temp_output.exists():
            temp_output.unlink(missing_ok=True)
        raise ExportError(f"No se pudo exportar el Excel: {exc}") from exc

    return output
