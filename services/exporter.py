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
    "presente": {"fill": None, "font": None},
    "domingo": {"fill": PatternFill("solid", fgColor="D9D9D9"), "font": None},  # gris
    "tarde": {"fill": PatternFill("solid", fgColor="2E7D32"), "font": Font(color="FFFFFF", bold=True)},  # verde oscuro
    "tardanza": {"fill": PatternFill("solid", fgColor="2E7D32"), "font": Font(color="FFFFFF", bold=True)},
    "ausente": {"fill": PatternFill("solid", fgColor="EF9A9A"), "font": Font(color="7F1D1D", bold=True)},  # rojo
    "ausencia": {"fill": PatternFill("solid", fgColor="EF9A9A"), "font": Font(color="7F1D1D", bold=True)},
    # Excepciones solicitadas
    "vacaciones": {"fill": PatternFill("solid", fgColor="FFF59D"), "font": None},  # amarillo
    "estudiar": {"fill": PatternFill("solid", fgColor="B3E5FC"), "font": None},  # celeste
    "capacitacion": {"fill": PatternFill("solid", fgColor="F8BBD0"), "font": None},  # rosado
    "capacitaciones": {"fill": PatternFill("solid", fgColor="F8BBD0"), "font": None},  # rosado
    "suspencion": {"fill": PatternFill("solid", fgColor="90CAF9"), "font": None},  # azul
    "suspension": {"fill": PatternFill("solid", fgColor="90CAF9"), "font": None},  # azul
    "sancion sin goce de sueldo": {"fill": PatternFill("solid", fgColor="90CAF9"), "font": None},  # azul
    "no trabajado": {"fill": PatternFill("solid", fgColor="D7CCC8"), "font": None},  # marron claro
    "licencia": {"fill": PatternFill("solid", fgColor="FFCC80"), "font": None},  # naranja
    "enfermedad": {"fill": PatternFill("solid", fgColor="FFCC80"), "font": None},  # naranja
    "feriado": {"fill": PatternFill("solid", fgColor="A9DF8F"), "font": Font(color="14532D", bold=True)},  # verde manzana
    "accidente de trabajo": {
        "fill": PatternFill("solid", fgColor="8E24AA"),  # violeta
        "font": Font(color="FFFFFF", bold=True),
    },
    "art": {"fill": PatternFill("solid", fgColor="8E24AA"), "font": Font(color="FFFFFF", bold=True)},  # alias
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
    "Liquidar": {
        "widths": {
            "Fecha": 12,
            "Dia": 14,
            "Dia #": 8,
        },
        "left_cols": {"Dia"},
        "wrap_cols": set(),
        "date_cols": {"Fecha"},
        "priority_sort": ["Fecha"],
    },
}

INCIDENCIAS_SHEET = "Incidencias"
INCIDENCIAS_SUMMARY_COLUMNS = ["ID de persona", "Nombre", "Total tardanzas", "Total ausencias"]
INCIDENCIAS_DETAIL_COLUMNS = ["ID de persona", "Nombre", "Fecha", "Estado", "Fichada"]


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


def _load_employee_catalog(base_dir: Path) -> tuple[list[tuple[str, str, str]], bool]:
    candidates = [base_dir / "date.xlsx", base_dir / "info.xlsx"]
    result: list[tuple[str, str, str]] = []
    by_id: dict[str, str] = {}
    loaded_from_file = False

    def _normalize(text: object) -> str:
        clean = unicodedata.normalize("NFKD", str(text)).encode("ascii", "ignore").decode("ascii")
        return " ".join(clean.strip().lower().replace("_", " ").split())

    def _clean_id(value: object) -> str:
        if pd.isna(value):
            return ""
        text = str(value).strip()
        if text.lower() in {"", "nan", "none", "nat", "-"}:
            return ""
        if text.endswith(".0") and text[:-2].isdigit():
            return text[:-2]
        return text

    for path in candidates:
        if not path.exists():
            continue
        try:
            with pd.ExcelFile(path) as excel:
                sheet_map = {_normalize(name): name for name in excel.sheet_names}
            sheet_name = sheet_map.get(_normalize("Empleados"))
            if sheet_name is not None:
                df = pd.read_excel(path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(path)
        except Exception:
            continue

        if df.empty:
            continue
        columns = [str(c) for c in df.columns]
        col_map = {_normalize(c): c for c in columns}
        id_col = None
        name_col = None
        for key, original in col_map.items():
            if id_col is None and ("id" == key or "legajo" in key or "id de persona" in key):
                id_col = original
            if name_col is None and ("nombre" == key or "empleado" in key or "name" == key):
                name_col = original
        if id_col is None:
            continue

        for _, row in df.iterrows():
            emp_id = _clean_id(row.get(id_col, ""))
            if not emp_id:
                continue
            emp_name = str(row.get(name_col, "")).strip() if name_col else ""
            if emp_name.lower() in {"", "nan", "none", "nat", "-"}:
                emp_name = f"ID {emp_id}"
            if emp_id not in by_id:
                by_id[emp_id] = emp_name
        loaded_from_file = True
        break

    for emp_id, emp_name in by_id.items():
        result.append((emp_id, emp_name, f"{emp_id} {emp_name}"))
    result.sort(key=lambda item: (item[1].lower(), item[0]))
    return result, loaded_from_file


def _build_incidencias_tables(
    daily_df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    summary_empty = pd.DataFrame(columns=INCIDENCIAS_SUMMARY_COLUMNS)
    detail_empty = pd.DataFrame(columns=INCIDENCIAS_DETAIL_COLUMNS)
    if daily_df.empty:
        return summary_empty, detail_empty

    needed = {"ID de persona", "Nombre", "Fecha", "Estado", "Tramos trabajados"}
    if not needed.issubset(set(daily_df.columns)):
        return summary_empty, detail_empty

    working = daily_df.copy()
    working["ID de persona"] = working["ID de persona"].astype(str).str.strip()
    working["Nombre"] = working["Nombre"].astype(str).str.strip()
    working["_FechaOrden"] = pd.to_datetime(working["Fecha"], errors="coerce")
    working["_estado_norm"] = working["Estado"].map(_normalize_status)
    working["_es_tarde"] = working["_estado_norm"].isin({"tarde", "tardanza"})
    working["_es_ausente"] = working["_estado_norm"].isin({"ausente", "ausencia"})

    employees = (
        working[["ID de persona", "Nombre"]]
        .drop_duplicates()
        .sort_values(["Nombre", "ID de persona"], kind="stable")
        .reset_index(drop=True)
    )
    tardy_counts = (
        working.groupby(["ID de persona", "Nombre"], as_index=False)["_es_tarde"]
        .sum()
        .rename(columns={"_es_tarde": "Total tardanzas"})
    )
    absent_counts = (
        working.groupby(["ID de persona", "Nombre"], as_index=False)["_es_ausente"]
        .sum()
        .rename(columns={"_es_ausente": "Total ausencias"})
    )

    summary = employees.merge(tardy_counts, on=["ID de persona", "Nombre"], how="left").merge(
        absent_counts,
        on=["ID de persona", "Nombre"],
        how="left",
    )
    summary["Total tardanzas"] = summary["Total tardanzas"].fillna(0).astype(int)
    summary["Total ausencias"] = summary["Total ausencias"].fillna(0).astype(int)
    summary = summary[INCIDENCIAS_SUMMARY_COLUMNS]

    detail = working[(working["_es_tarde"]) | (working["_es_ausente"])].copy()
    if detail.empty:
        return summary, detail_empty

    detail = detail.sort_values(
        ["ID de persona", "_FechaOrden", "Nombre"],
        kind="stable",
    )
    detail_df = pd.DataFrame(
        {
            "ID de persona": detail["ID de persona"],
            "Nombre": detail["Nombre"],
            "Fecha": detail["Fecha"],
            "Estado": detail["Estado"].astype(str),
            "Fichada": detail["Tramos trabajados"].fillna("").astype(str),
        }
    )
    return summary, detail_df[INCIDENCIAS_DETAIL_COLUMNS]


def _build_liquidar_sheet(
    daily_df: pd.DataFrame,
    base_dir: Path,
) -> tuple[pd.DataFrame, dict[tuple[int, int], str]]:
    cols_base = ["Fecha", "Dia", "Dia #"]
    if daily_df.empty:
        return pd.DataFrame(columns=cols_base), {}

    needed = {"Fecha", "ID de persona", "Nombre", "Minutos redondeados", "Minutos extra", "Estado"}
    if not needed.issubset(set(daily_df.columns)):
        return pd.DataFrame(columns=cols_base), {}

    working = daily_df.copy()
    working["_fecha"] = pd.to_datetime(working["Fecha"], errors="coerce")
    working = working[working["_fecha"].notna()].copy()
    if working.empty:
        return pd.DataFrame(columns=cols_base), {}

    def _clean_id(value: object) -> str:
        text = str(value).strip()
        if text.endswith(".0") and text[:-2].isdigit():
            return text[:-2]
        return text

    catalog, loaded_from_file = _load_employee_catalog(base_dir)
    from_daily: dict[str, str] = {}
    for _, row in working.iterrows():
        emp_id = _clean_id(row["ID de persona"])
        emp_name = str(row["Nombre"]).strip()
        if emp_id and emp_id not in from_daily:
            from_daily[emp_id] = emp_name

    catalog_by_id = {emp_id: (emp_name, label) for emp_id, emp_name, label in catalog}
    # Si no hay plantilla de empleados disponible, usa fallback desde el Diario.
    # Si hay date.xlsx/info.xlsx cargado, respeta solo esos empleados.
    if not loaded_from_file:
        for emp_id, emp_name in from_daily.items():
            if emp_id not in catalog_by_id:
                catalog_by_id[emp_id] = (emp_name, f"{emp_id} {emp_name}")

    employees = [(emp_id, name, label) for emp_id, (name, label) in catalog_by_id.items()]
    employees.sort(key=lambda item: (item[1].lower(), item[0]))
    employee_labels = [item[2] for item in employees]

    day_names = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado", "Domingo"]
    by_key: dict[tuple[pd.Timestamp, str], dict] = {}
    for _, row in working.iterrows():
        day = row["_fecha"].normalize()
        emp_id = _clean_id(row["ID de persona"])
        by_key[(day, emp_id)] = row

    min_day = working["_fecha"].min().normalize()
    max_day = working["_fecha"].max().normalize()
    worked_saturdays = {
        day.normalize()
        for day in working.loc[working["_fecha"].dt.weekday == 5, "_fecha"]
    }
    all_days = [
        day
        for day in pd.date_range(min_day, max_day, freq="D")
        if int(day.weekday()) < 5 or day.normalize() in worked_saturdays
    ]

    rows: list[dict[str, object]] = []
    cell_status_map: dict[tuple[int, int], str] = {}
    totals_minutes_by_employee: dict[str, dict[str, int]] = {
        emp_id: {
            "q1_common": 0,
            "q1_extra": 0,
            "q2_common": 0,
            "q2_extra": 0,
        }
        for emp_id, _, _ in employees
    }
    excel_row = 2  # header is row 1

    def _append_day_row(day: pd.Timestamp) -> None:
        nonlocal excel_row
        row_data: dict[str, object] = {
            "Fecha": day.date(),
            "Dia": day_names[int(day.weekday())],
            "Dia #": int(day.day),
        }
        for emp_index, (emp_id, _, label) in enumerate(employees, start=0):
            rec = by_key.get((day, emp_id))
            if rec is None:
                row_data[label] = ""
                continue
            total_minutes = int(pd.to_numeric(rec.get("Minutos redondeados", 0), errors="coerce") or 0)
            extra_minutes = int(pd.to_numeric(rec.get("Minutos extra", 0), errors="coerce") or 0)
            normal_minutes = max(0, total_minutes - extra_minutes)
            row_data[label] = round(normal_minutes / 60, 2)
            half = "q1" if int(day.day) <= 15 else "q2"
            totals_minutes_by_employee[emp_id][f"{half}_common"] += normal_minutes
            totals_minutes_by_employee[emp_id][f"{half}_extra"] += max(0, extra_minutes)
            status = str(rec.get("Estado", "")).strip()
            # Employee cols start after A,B,C
            cell_status_map[(excel_row, 4 + emp_index)] = status
        rows.append(row_data)
        excel_row += 1

    def _hours(minutes: int) -> float:
        return round(max(0, int(minutes)) / 60, 2)

    def _summary_row(label_text: str, metric_key: str) -> dict[str, object]:
        row_data: dict[str, object] = {"Fecha": "", "Dia": label_text, "Dia #": ""}
        for emp_id, _, emp_label in employees:
            row_data[emp_label] = _hours(totals_minutes_by_employee[emp_id][metric_key])
        return row_data

    q1_days = [day for day in all_days if int(day.day) <= 15]
    q2_days = [day for day in all_days if int(day.day) > 15]

    for day in q1_days:
        _append_day_row(day)

    if q1_days:
        rows.append(_summary_row("1ra Quincena - Horas Comunes", "q1_common"))
        rows.append(_summary_row("1ra Quincena - Horas Extras", "q1_extra"))
        excel_row += 2

    for day in q2_days:
        _append_day_row(day)

    if q2_days:
        rows.append(_summary_row("2da Quincena - Horas Comunes", "q2_common"))
        rows.append(_summary_row("2da Quincena - Horas Extras", "q2_extra"))
        excel_row += 2

    total_common_row: dict[str, object] = {"Fecha": "", "Dia": "Total Mes - Horas Comunes", "Dia #": ""}
    total_extra_row: dict[str, object] = {"Fecha": "", "Dia": "Total Mes - Horas Extras", "Dia #": ""}
    for emp_id, _, emp_label in employees:
        total_common_row[emp_label] = _hours(
            totals_minutes_by_employee[emp_id]["q1_common"] + totals_minutes_by_employee[emp_id]["q2_common"]
        )
        total_extra_row[emp_label] = _hours(
            totals_minutes_by_employee[emp_id]["q1_extra"] + totals_minutes_by_employee[emp_id]["q2_extra"]
        )
    rows.extend([total_common_row, total_extra_row])

    liquidar_df = pd.DataFrame(rows, columns=cols_base + employee_labels)
    return liquidar_df, cell_status_map


def _apply_liquidar_format(
    worksheet,
    status_cells: dict[tuple[int, int], str],
) -> None:
    headers = [cell.value for cell in worksheet[1]]
    header_to_idx = {str(value): idx + 1 for idx, value in enumerate(headers) if value is not None}

    worksheet.freeze_panes = "D2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for cell in worksheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER

    for header, width in SHEET_LAYOUTS["Liquidar"]["widths"].items():
        col_idx = header_to_idx.get(header)
        if col_idx:
            worksheet.column_dimensions[get_column_letter(col_idx)].width = width

    for col_idx in range(4, worksheet.max_column + 1):
        worksheet.column_dimensions[get_column_letter(col_idx)].width = 17

    for row_idx in range(2, worksheet.max_row + 1):
        first_col_value = worksheet.cell(row=row_idx, column=1).value
        is_daily_row = pd.notna(pd.to_datetime(first_col_value, errors="coerce"))
        row_label = str(worksheet.cell(row=row_idx, column=2).value or "").strip().lower()
        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.border = THIN_BORDER
            if is_daily_row and row_idx % 2 == 0:
                cell.fill = ALT_ROW_FILL
            if (not is_daily_row) and row_label:
                if "horas extras" in row_label:
                    cell.fill = PatternFill("solid", fgColor="FFD7B5")
                elif "horas comunes" in row_label:
                    cell.fill = PatternFill("solid", fgColor="C7E0FF")
                # Quincenas y total mensual: resaltar letras y valores.
                cell.font = Font(color="0F172A", bold=True)

            if col_idx == 1 and is_daily_row and cell.value not in ("", None):
                cell.number_format = "DD/MM/YYYY"
            if col_idx >= 4:
                cell.number_format = "0.00"

            horizontal = "center"
            if col_idx in {2}:
                horizontal = "left"
            cell.alignment = Alignment(horizontal=horizontal, vertical="center")

    # Color cells by status for liquidation view
    for (row_idx, col_idx), status in status_cells.items():
        style = _status_style(status)
        fill = style.get("fill")
        font = style.get("font")
        cell = worksheet.cell(row=row_idx, column=col_idx)
        if fill is not None:
            cell.fill = fill
        if font is not None:
            cell.font = font

    _append_liquidar_legend(worksheet)


def _append_liquidar_legend(worksheet) -> None:
    legend_items = [
        ("Domingo", "domingo"),
        ("Ausente", "ausente"),
        ("Vacaciones", "vacaciones"),
        ("Estudiar", "estudiar"),
        ("Capacitacion", "capacitacion"),
        ("Suspencion", "suspencion"),
        ("No trabajado", "no trabajado"),
        ("Licencia", "licencia"),
        ("Feriado", "feriado"),
        ("Accidente de trabajo", "accidente de trabajo"),
        ("Tardanza", "tardanza"),
    ]

    start_row = worksheet.max_row + 2
    title_cell = worksheet.cell(row=start_row, column=1)
    title_cell.value = "Leyenda de estados"
    title_cell.fill = HEADER_FILL
    title_cell.font = HEADER_FONT
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    title_cell.border = THIN_BORDER
    worksheet.merge_cells(
        start_row=start_row,
        start_column=1,
        end_row=start_row,
        end_column=4,
    )
    for col in range(1, 5):
        worksheet.cell(row=start_row, column=col).border = THIN_BORDER

    for idx, (label, status_key) in enumerate(legend_items):
        row_idx = start_row + 1 + (idx // 2)
        if idx % 2 == 0:
            box_col, text_col = 1, 2
        else:
            box_col, text_col = 3, 4

        box = worksheet.cell(row=row_idx, column=box_col)
        text = worksheet.cell(row=row_idx, column=text_col)
        style = _status_style(status_key)

        box.value = ""
        box.border = THIN_BORDER
        fill = style.get("fill")
        if fill is not None:
            box.fill = fill

        text.value = label
        text.border = THIN_BORDER
        text.alignment = Alignment(horizontal="left", vertical="center")
        text.font = Font(color="000000", bold=True)


def _apply_incidencias_format(worksheet, summary_rows: int, detail_header_row: int) -> None:
    headers_row_1 = [cell.value for cell in worksheet[1]]
    header_row_1_idx = {str(value): idx + 1 for idx, value in enumerate(headers_row_1) if value is not None}

    worksheet.freeze_panes = "A2"

    for col_idx, width in enumerate((14, 30, 16, 16, 56), start=1):
        worksheet.column_dimensions[get_column_letter(col_idx)].width = width

    # Header resumen (fila 1)
    for cell in worksheet[1]:
        if cell.value is None:
            continue
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER

    # Filas resumen
    for row_idx in range(2, summary_rows + 2):
        for col_idx in range(1, 5):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.border = THIN_BORDER
            if row_idx % 2 == 0:
                cell.fill = ALT_ROW_FILL
            horizontal = "left" if col_idx == 2 else "center"
            cell.alignment = Alignment(horizontal=horizontal, vertical="center")

    # Header detalle
    for col_idx in range(1, 6):
        cell = worksheet.cell(row=detail_header_row, column=col_idx)
        if cell.value is None:
            continue
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER

    # Detalle filas
    for row_idx in range(detail_header_row + 1, worksheet.max_row + 1):
        status_text = worksheet.cell(row=row_idx, column=4).value
        style = _status_style(status_text)
        row_fill = style.get("fill")
        row_font = style.get("font")
        for col_idx in range(1, 6):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.border = THIN_BORDER
            if row_fill is not None:
                cell.fill = row_fill
            elif row_idx % 2 == 0:
                cell.fill = ALT_ROW_FILL
            if row_font is not None:
                cell.font = row_font
            if col_idx == 3 and cell.value not in ("", None):
                cell.number_format = "DD/MM/YYYY"
            if col_idx in {2, 5}:
                horizontal = "left"
            else:
                horizontal = "center"
            cell.alignment = Alignment(
                horizontal=horizontal,
                vertical="top",
                wrap_text=(col_idx == 5),
            )

    # Keep filter on summary header row only
    if header_row_1_idx:
        worksheet.auto_filter.ref = f"A1:D{max(1, summary_rows + 1)}"


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
    incidencias_summary, incidencias_detail = _build_incidencias_tables(diario_export)
    liquidar_export, liquidar_status_cells = _build_liquidar_sheet(
        diario_export,
        output.parent,
    )
    try:
        with pd.ExcelWriter(temp_output, engine="openpyxl") as writer:
            diario_export.to_excel(writer, sheet_name="Diario", index=False)
            mensual_export.to_excel(writer, sheet_name="Mensual", index=False)
            incidencias_summary.to_excel(writer, sheet_name=INCIDENCIAS_SHEET, index=False)
            detail_startrow = len(incidencias_summary) + 3
            incidencias_detail.to_excel(
                writer,
                sheet_name=INCIDENCIAS_SHEET,
                index=False,
                startrow=detail_startrow,
            )
            liquidar_export.to_excel(writer, sheet_name="Liquidar", index=False)

            workbook = writer.book
            for sheet_name in ("Diario", "Mensual"):
                _apply_sheet_format(workbook[sheet_name], sheet_name)
            _apply_incidencias_format(
                workbook[INCIDENCIAS_SHEET],
                summary_rows=len(incidencias_summary),
                detail_header_row=detail_startrow + 1,
            )
            _apply_liquidar_format(workbook["Liquidar"], liquidar_status_cells)

        temp_output.replace(output)
    except Exception as exc:
        if temp_output.exists():
            temp_output.unlink(missing_ok=True)
        raise ExportError(f"No se pudo exportar el Excel: {exc}") from exc

    return output
