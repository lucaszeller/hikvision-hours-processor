from __future__ import annotations

from pathlib import Path
import unicodedata

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


class ExportError(Exception):
    pass


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
ALT_ROW_FILL = PatternFill("solid", fgColor="F5F8FC")
STATUS_STYLES = {
    # Base
    "normal": {"fill": None, "font": None},
    "domingo": {"fill": PatternFill("solid", fgColor="D9D9D9"), "font": None},  # gris claro
    "tarde": {"fill": PatternFill("solid", fgColor="2E7D32"), "font": Font(color="FFFFFF", bold=True)},  # verde oscuro
    "tardanza": {"fill": PatternFill("solid", fgColor="2E7D32"), "font": Font(color="FFFFFF", bold=True)},
    "ausente": {"fill": PatternFill("solid", fgColor="EF9A9A"), "font": Font(color="7F1D1D", bold=True)},  # rojo
    "ausencia": {"fill": PatternFill("solid", fgColor="EF9A9A"), "font": Font(color="7F1D1D", bold=True)},
    # Excepciones solicitadas
    "vacaciones": {"fill": PatternFill("solid", fgColor="FFF59D"), "font": None},  # amarillo
    "estudiar": {"fill": PatternFill("solid", fgColor="B3E5FC"), "font": None},  # celeste
    "capacitacion": {"fill": PatternFill("solid", fgColor="D1C4E9"), "font": None},  # morado claro
    "sancion sin goce de sueldo": {
        "fill": PatternFill("solid", fgColor="90CAF9"),  # azul
        "font": None,
    },
    "no trabajado": {"fill": PatternFill("solid", fgColor="FFF9C4"), "font": None},  # amarillo claro
    "licencia": {"fill": PatternFill("solid", fgColor="FFCC80"), "font": None},  # naranja
    "feriado": {"fill": PatternFill("solid", fgColor="A9DF8F"), "font": Font(color="14532D", bold=True)},  # verde manzana
    "accidente de trabajo": {
        "fill": PatternFill("solid", fgColor="8E24AA"),  # violeta fuerte
        "font": Font(color="FFFFFF", bold=True),
    },
}
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
    "resumen-estudio": {
        "widths": {
            "Dia de semana": 18,
            "Nombre del empleado": 34,
            "Horas normales": 18,
            "Horas extra": 16,
        },
        "left_cols": {"Dia de semana", "Nombre del empleado"},
        "wrap_cols": set(),
        "date_cols": set(),
        "priority_sort": ["Dia de semana", "Nombre del empleado"],
    },
}


def _sort_for_report(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    available = [col for col in columns if col in df.columns]
    if not available or df.empty:
        return df
    return df.sort_values(available, kind="stable").reset_index(drop=True)


def _minutes_to_hhmm(total_minutes: int) -> str:
    hours, minutes = divmod(max(0, int(total_minutes)), 60)
    return f"{hours:02d}:{minutes:02d}"


def _normalize_status(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode("ascii")
    return " ".join(text.strip().lower().split())


def _status_style(status_text: object) -> dict[str, object]:
    key = _normalize_status(status_text)
    if key in STATUS_STYLES:
        return STATUS_STYLES[key]
    if key and key not in {"normal"}:
        # Cualquier excepcion no mapeada explicitamente mantiene verde manzana por defecto.
        return STATUS_STYLES["feriado"]
    return STATUS_STYLES["normal"]


def _weekday_name_es(value: object) -> str:
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return ""
    names = [
        "Lunes",
        "Martes",
        "Miercoles",
        "Jueves",
        "Viernes",
        "Sabado",
        "Domingo",
    ]
    return names[int(parsed.weekday())]


def _build_study_summary(daily_df: pd.DataFrame) -> pd.DataFrame:
    columns = ["Dia de semana", "Nombre del empleado", "Horas normales", "Horas extra"]
    if daily_df.empty:
        return pd.DataFrame(columns=columns)

    needed = {"Fecha", "Nombre", "Minutos redondeados", "Minutos extra"}
    if not needed.issubset(set(daily_df.columns)):
        return pd.DataFrame(columns=columns)

    working = daily_df.copy()
    working["_FechaOrden"] = pd.to_datetime(working["Fecha"], errors="coerce")
    working = working.sort_values(["_FechaOrden", "Nombre"], kind="stable").reset_index(drop=True)
    total_minutes = pd.to_numeric(working["Minutos redondeados"], errors="coerce").fillna(0).astype(int)
    extra_minutes = pd.to_numeric(working["Minutos extra"], errors="coerce").fillna(0).astype(int)
    normal_minutes = (total_minutes - extra_minutes).clip(lower=0)

    summary = pd.DataFrame(
        {
            "Dia de semana": working["Fecha"].map(_weekday_name_es),
            "Nombre del empleado": working["Nombre"].astype(str),
            "Horas normales": normal_minutes.map(_minutes_to_hhmm),
            "Horas extra": extra_minutes.map(_minutes_to_hhmm),
        }
    )
    return summary


def _apply_sheet_format(worksheet, sheet_name: str) -> None:
    config = SHEET_LAYOUTS[sheet_name]
    headers = [cell.value for cell in worksheet[1]]
    header_to_idx = {str(value): idx + 1 for idx, value in enumerate(headers) if value is not None}
    status_col_idx = header_to_idx.get("Estado")

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
        row_fill = None
        row_font = None
        row_status = ""
        if sheet_name == "Diario" and status_col_idx is not None:
            row_status = str(worksheet.cell(row=row_idx, column=status_col_idx).value or "").strip().lower()
            style = _status_style(row_status)
            row_fill = style.get("fill")
            row_font = style.get("font")

        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            if row_fill is not None:
                cell.fill = row_fill
            elif row_idx % 2 == 0:
                cell.fill = ALT_ROW_FILL

        for header, col_idx in header_to_idx.items():
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.border = THIN_BORDER

            horizontal = "left" if header in config["left_cols"] else "center"
            wrap = header in config["wrap_cols"]
            cell.alignment = Alignment(horizontal=horizontal, vertical="top", wrap_text=wrap)
            if row_font is not None:
                cell.font = row_font

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

    diario_export = _sort_for_report(daily_df.copy(), SHEET_LAYOUTS["Diario"]["priority_sort"])
    mensual_export = _sort_for_report(monthly_df.copy(), SHEET_LAYOUTS["Mensual"]["priority_sort"])
    resumen_estudio_export = _build_study_summary(diario_export)
    try:
        with pd.ExcelWriter(temp_output, engine="openpyxl") as writer:
            diario_export.to_excel(writer, sheet_name="Diario", index=False)
            mensual_export.to_excel(writer, sheet_name="Mensual", index=False)
            resumen_estudio_export.to_excel(writer, sheet_name="resumen-estudio", index=False)

            workbook = writer.book
            for sheet_name in ("Diario", "Mensual", "resumen-estudio"):
                _apply_sheet_format(workbook[sheet_name], sheet_name)

        temp_output.replace(output)
    except Exception as exc:
        if temp_output.exists():
            temp_output.unlink(missing_ok=True)
        raise ExportError(f"No se pudo exportar el Excel: {exc}") from exc

    return output
