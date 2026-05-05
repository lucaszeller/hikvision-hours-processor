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
LATE_FILL = PatternFill("solid", fgColor="FFF59D")
ABSENT_FILL = PatternFill("solid", fgColor="EF9A9A")
ABSENT_FONT = Font(color="7F1D1D", bold=True)
EXCEPTION_FILL = PatternFill("solid", fgColor="A9DF8F")
EXCEPTION_FONT = Font(color="14532D", bold=True)
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)

SHEET_LAYOUTS = {
    "Diario": {
        "widths": {
            "ID de persona": 14,
            "Nombre": 28,
            "Fecha": 12,
            "Departamento": 22,
            "Estado": 12,
            "Tramos trabajados": 64,
            "Minutos reales": 16,
            "Minutos redondeados": 20,
            "Minutos extra": 16,
            "Horas extra": 14,
            "Horas totales": 14,
        },
        "left_cols": {"Nombre", "Departamento", "Tramos trabajados"},
        "wrap_cols": {"Tramos trabajados"},
        "date_cols": {"Fecha"},
        "priority_sort": ["ID de persona", "Fecha", "Nombre"],
    },
    "Mensual": {
        "widths": {
            "ID de persona": 14,
            "Nombre": 28,
            "Dias trabajados": 16,
            "Minutos totales": 16,
            "Minutos extra": 16,
            "Horas extra": 14,
            "Horas totales": 14,
        },
        "left_cols": {"Nombre"},
        "wrap_cols": set(),
        "date_cols": set(),
        "priority_sort": ["ID de persona", "Nombre"],
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

            if sheet_name == "Diario" and header == "Estado":
                status = str(cell.value or "").strip().lower()
                if status == "tarde":
                    cell.fill = LATE_FILL
                elif status == "ausente":
                    cell.fill = ABSENT_FILL
                    cell.font = ABSENT_FONT
                elif status not in {"", "normal"}:
                    cell.fill = EXCEPTION_FILL
                    cell.font = EXCEPTION_FONT


def export_report(
    output_path: str | Path,
    daily_df: pd.DataFrame,
    monthly_df: pd.DataFrame,
    inconsistencies_df: pd.DataFrame,
) -> Path:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    temp_output = output.with_suffix(f"{output.suffix}.tmp")

    diario_export = _sort_for_report(daily_df.copy(), SHEET_LAYOUTS["Diario"]["priority_sort"])
    mensual_export = _sort_for_report(monthly_df.copy(), SHEET_LAYOUTS["Mensual"]["priority_sort"])
    try:
        with pd.ExcelWriter(temp_output, engine="openpyxl") as writer:
            diario_export.to_excel(writer, sheet_name="Diario", index=False)
            mensual_export.to_excel(writer, sheet_name="Mensual", index=False)

            workbook = writer.book
            for sheet_name in ("Diario", "Mensual"):
                _apply_sheet_format(workbook[sheet_name], sheet_name)

        temp_output.replace(output)
    except Exception as exc:
        if temp_output.exists():
            temp_output.unlink(missing_ok=True)
        raise ExportError(f"No se pudo exportar el Excel: {exc}") from exc

    return output
