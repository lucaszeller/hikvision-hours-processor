from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from services.exporter import export_report


def _rgb(cell) -> str:
    return str(cell.fill.fgColor.rgb or "")


def test_diario_status_colors_fill_entire_row(tmp_path: Path) -> None:
    daily_df = pd.DataFrame(
        {
            "ID de persona": ["20", "30"],
            "Nombre": ["Ana", "Luis"],
            "Fecha": [pd.Timestamp("2026-05-01").date(), pd.Timestamp("2026-05-01").date()],
            "Departamento": ["A", "B"],
            "Estado": ["Tarde", "Ausente"],
            "Tramos trabajados": ["07:31 - 12:00", ""],
            "Minutos reales": [269, 0],
            "Minutos redondeados": [270, 0],
            "Minutos extra": [0, 0],
            "Horas extra": ["00:00", "00:00"],
            "Horas totales": ["04:30", "00:00"],
        }
    )
    monthly_df = pd.DataFrame(
        columns=[
            "ID de persona",
            "Nombre",
            "Dias trabajados",
            "Minutos totales",
            "Minutos extra",
            "Horas extra",
            "Horas totales",
        ]
    )
    inconsistencies_df = pd.DataFrame(
        columns=["ID de persona", "Nombre", "Fecha", "Tipo de inconsistencia", "Detalle"]
    )

    output = tmp_path / "reporte.xlsx"
    export_report(output, daily_df, monthly_df, inconsistencies_df)

    wb = load_workbook(output)
    ws = wb["Diario"]

    # Fila 2 = "Tarde" -> verde oscuro en toda la fila.
    assert _rgb(ws["A2"]).endswith("2E7D32")
    assert _rgb(ws["F2"]).endswith("2E7D32")

    # Fila 3 = "Ausente" -> rojo en toda la fila.
    assert _rgb(ws["A3"]).endswith("EF9A9A")
    assert _rgb(ws["J3"]).endswith("EF9A9A")


def test_diario_domingo_row_is_light_gray(tmp_path: Path) -> None:
    daily_df = pd.DataFrame(
        {
            "ID de persona": ["20"],
            "Nombre": ["Ana"],
            "Fecha": [pd.Timestamp("2026-05-10").date()],
            "Departamento": ["A"],
            "Estado": ["Domingo"],
            "Tramos trabajados": ["08:00 - 12:00"],
            "Minutos reales": [240],
            "Minutos redondeados": [240],
            "Minutos extra": [240],
            "Horas extra": ["04:00"],
            "Horas totales": ["04:00"],
        }
    )
    monthly_df = pd.DataFrame(
        columns=[
            "ID de persona",
            "Nombre",
            "Dias trabajados",
            "Minutos totales",
            "Minutos extra",
            "Horas extra",
            "Horas totales",
        ]
    )
    inconsistencies_df = pd.DataFrame(
        columns=["ID de persona", "Nombre", "Fecha", "Tipo de inconsistencia", "Detalle"]
    )

    output = tmp_path / "reporte_domingo.xlsx"
    export_report(output, daily_df, monthly_df, inconsistencies_df)

    wb = load_workbook(output)
    ws = wb["Diario"]
    assert _rgb(ws["A2"]).endswith("D9D9D9")
    assert _rgb(ws["G2"]).endswith("D9D9D9")


def test_diario_presente_row_has_no_status_color_override(tmp_path: Path) -> None:
    daily_df = pd.DataFrame(
        {
            "ID de persona": ["20"],
            "Nombre": ["Ana"],
            "Fecha": [pd.Timestamp("2026-05-20").date()],
            "Departamento": ["A"],
            "Estado": ["Presente"],
            "Tramos trabajados": [""],
            "Minutos reales": [540],
            "Minutos redondeados": [540],
            "Minutos extra": [0],
            "Horas extra": ["00:00"],
            "Horas totales": ["09:00"],
        }
    )
    monthly_df = pd.DataFrame(
        columns=[
            "ID de persona",
            "Nombre",
            "Dias trabajados",
            "Minutos totales",
            "Minutos extra",
            "Horas extra",
            "Horas totales",
        ]
    )
    inconsistencies_df = pd.DataFrame(
        columns=["ID de persona", "Nombre", "Fecha", "Tipo de inconsistencia", "Detalle"]
    )

    output = tmp_path / "reporte_presente.xlsx"
    export_report(output, daily_df, monthly_df, inconsistencies_df)

    wb = load_workbook(output)
    ws = wb["Diario"]
    # Fila par con estado sin color especifico: conserva alternado, no color de estado.
    assert _rgb(ws["A2"]).endswith("F5F8FC")


def test_export_does_not_create_resumen_estudio_sheet(tmp_path: Path) -> None:
    daily_df = pd.DataFrame(
        {
            "ID de persona": ["20", "30"],
            "Nombre": ["Ana", "Luis"],
            "Fecha": [pd.Timestamp("2026-05-05").date(), pd.Timestamp("2026-05-06").date()],
            "Departamento": ["A", "B"],
            "Estado": ["Normal", "Tarde"],
            "Tramos trabajados": ["07:30 - 12:00", "07:31 - 12:00"],
            "Minutos reales": [270, 269],
            "Minutos redondeados": [270, 270],
            "Minutos extra": [30, 0],
            "Horas extra": ["00:30", "00:00"],
            "Horas totales": ["04:30", "04:30"],
        }
    )
    monthly_df = pd.DataFrame(
        columns=[
            "ID de persona",
            "Nombre",
            "Dias trabajados",
            "Minutos totales",
            "Minutos extra",
            "Horas extra",
            "Horas totales",
        ]
    )
    inconsistencies_df = pd.DataFrame(
        columns=["ID de persona", "Nombre", "Fecha", "Tipo de inconsistencia", "Detalle"]
    )

    output = tmp_path / "reporte_resumen.xlsx"
    export_report(output, daily_df, monthly_df, inconsistencies_df)

    wb = load_workbook(output)
    assert "resumen-estudio" not in wb.sheetnames


def test_export_creates_liquidar_sheet_with_status_colors(tmp_path: Path) -> None:
    daily_df = pd.DataFrame(
        {
            "ID de persona": ["20", "30", "20", "30"],
            "Nombre": ["Ana", "Luis", "Ana", "Luis"],
            "Fecha": [
                pd.Timestamp("2026-05-05").date(),
                pd.Timestamp("2026-05-05").date(),
                pd.Timestamp("2026-05-06").date(),
                pd.Timestamp("2026-05-06").date(),
            ],
            "Departamento": ["A", "B", "A", "B"],
            "Estado": ["Vacaciones", "Feriado", "Enfermedad", "Tardanza"],
            "Tramos trabajados": ["", "", "", "07:31 - 12:00"],
            "Minutos reales": [540, 540, 0, 270],
            "Minutos redondeados": [540, 540, 0, 270],
            "Minutos extra": [0, 0, 0, 0],
            "Horas extra": ["00:00", "00:00", "00:00", "00:00"],
            "Horas totales": ["09:00", "09:00", "00:00", "04:30"],
        }
    )
    monthly_df = pd.DataFrame(
        columns=[
            "ID de persona",
            "Nombre",
            "Dias trabajados",
            "Minutos totales",
            "Minutos extra",
            "Horas extra",
            "Horas totales",
        ]
    )
    inconsistencies_df = pd.DataFrame(
        columns=["ID de persona", "Nombre", "Fecha", "Tipo de inconsistencia", "Detalle"]
    )

    output = tmp_path / "reporte_liquidar.xlsx"
    export_report(output, daily_df, monthly_df, inconsistencies_df)

    wb = load_workbook(output)
    assert "Liquidar" in wb.sheetnames
    ws = wb["Liquidar"]

    # Columnas base
    assert ws["A1"].value == "Fecha"
    assert ws["B1"].value == "Dia"
    assert ws["C1"].value == "Dia #"

    # Empleados esperados
    assert ws["D1"].value == "20 Ana"
    assert ws["E1"].value == "30 Luis"

    # 2026-05-05 row
    assert str(ws["A2"].value)[:10] == "2026-05-05"
    assert ws["D2"].value == 9
    assert ws["E2"].value == 9
    # Vacaciones amarillo / Feriado verde manzana
    assert _rgb(ws["D2"]).endswith("FFF59D")
    assert _rgb(ws["E2"]).endswith("A9DF8F")

    # 2026-05-06 row
    assert str(ws["A3"].value)[:10] == "2026-05-06"
    assert ws["D3"].value == 0
    assert ws["E3"].value == 4.5
    # Enfermedad naranja, tardanza verde oscuro.
    assert _rgb(ws["D3"]).endswith("FFCC80")
    assert _rgb(ws["E3"]).endswith("2E7D32")


def test_liquidar_uses_only_normal_hours_excluding_overtime(tmp_path: Path) -> None:
    daily_df = pd.DataFrame(
        {
            "ID de persona": ["20"],
            "Nombre": ["Ana"],
            "Fecha": [pd.Timestamp("2026-05-05").date()],
            "Departamento": ["A"],
            "Estado": ["Tardanza"],
            "Tramos trabajados": ["07:30 - 17:00"],
            "Minutos reales": [570],
            "Minutos redondeados": [570],   # 9.5h totales
            "Minutos extra": [90],          # 1.5h extra
            "Horas extra": ["01:30"],
            "Horas totales": ["09:30"],
        }
    )
    monthly_df = pd.DataFrame(
        columns=[
            "ID de persona",
            "Nombre",
            "Dias trabajados",
            "Minutos totales",
            "Minutos extra",
            "Horas extra",
            "Horas totales",
        ]
    )
    inconsistencies_df = pd.DataFrame(
        columns=["ID de persona", "Nombre", "Fecha", "Tipo de inconsistencia", "Detalle"]
    )

    output = tmp_path / "reporte_liquidar_normales.xlsx"
    export_report(output, daily_df, monthly_df, inconsistencies_df)

    wb = load_workbook(output)
    ws = wb["Liquidar"]

    # 9.5h - 1.5h = 8.0h normales
    assert ws["D2"].value == 8


def test_liquidar_uses_only_employees_from_date_xlsx_when_available(tmp_path: Path) -> None:
    empleados_df = pd.DataFrame(
        {
            "Id": ["20"],
            "Nombre": ["Ana"],
        }
    )
    with pd.ExcelWriter(tmp_path / "date.xlsx", engine="openpyxl") as writer:
        empleados_df.to_excel(writer, sheet_name="Empleados", index=False)

    daily_df = pd.DataFrame(
        {
            "ID de persona": ["20", "30"],
            "Nombre": ["Ana", "Luis"],
            "Fecha": [pd.Timestamp("2026-05-05").date(), pd.Timestamp("2026-05-05").date()],
            "Departamento": ["A", "B"],
            "Estado": ["Normal", "Normal"],
            "Tramos trabajados": ["07:30 - 12:00", "08:00 - 12:00"],
            "Minutos reales": [270, 240],
            "Minutos redondeados": [270, 240],
            "Minutos extra": [0, 0],
            "Horas extra": ["00:00", "00:00"],
            "Horas totales": ["04:30", "04:00"],
        }
    )
    monthly_df = pd.DataFrame(
        columns=[
            "ID de persona",
            "Nombre",
            "Dias trabajados",
            "Minutos totales",
            "Minutos extra",
            "Horas extra",
            "Horas totales",
        ]
    )
    inconsistencies_df = pd.DataFrame(
        columns=["ID de persona", "Nombre", "Fecha", "Tipo de inconsistencia", "Detalle"]
    )

    output = tmp_path / "reporte_liquidar_catalogo.xlsx"
    export_report(output, daily_df, monthly_df, inconsistencies_df)

    wb = load_workbook(output)
    ws = wb["Liquidar"]

    # Solo debe incluir legajo 20 (de date.xlsx), no el 30.
    headers = [ws.cell(row=1, column=col).value for col in range(1, ws.max_column + 1)]
    assert "20 Ana" in headers
    assert "30 Luis" not in headers


def test_liquidar_adds_quincena_common_and_extra_totals(tmp_path: Path) -> None:
    daily_df = pd.DataFrame(
        {
            "ID de persona": ["20", "20", "20"],
            "Nombre": ["Ana", "Ana", "Ana"],
                "Fecha": [
                    pd.Timestamp("2026-05-05").date(),  # 1ra quincena
                    pd.Timestamp("2026-05-18").date(),  # 2da quincena
                    pd.Timestamp("2026-05-20").date(),  # 2da quincena
                ],
            "Departamento": ["A", "A", "A"],
            "Estado": ["Normal", "Tardanza", "Normal"],
            "Tramos trabajados": ["07:30 - 12:00", "07:30 - 17:30", "07:30 - 12:00"],
            "Minutos reales": [270, 600, 270],
            "Minutos redondeados": [270, 600, 270],
            "Minutos extra": [30, 120, 0],
            "Horas extra": ["00:30", "02:00", "00:00"],
            "Horas totales": ["04:30", "10:00", "04:30"],
        }
    )
    monthly_df = pd.DataFrame(
        columns=[
            "ID de persona",
            "Nombre",
            "Dias trabajados",
            "Minutos totales",
            "Minutos extra",
            "Horas extra",
            "Horas totales",
        ]
    )
    inconsistencies_df = pd.DataFrame(
        columns=["ID de persona", "Nombre", "Fecha", "Tipo de inconsistencia", "Detalle"]
    )

    output = tmp_path / "reporte_liquidar_quincena.xlsx"
    export_report(output, daily_df, monthly_df, inconsistencies_df)

    wb = load_workbook(output)
    ws = wb["Liquidar"]

    labels = {}
    for row in range(2, ws.max_row + 1):
        labels[str(ws.cell(row=row, column=2).value)] = row

    # 1ra quincena: comunes=4.0h, extras=0.5h
    r = labels["1ra Quincena - Horas Comunes"]
    assert ws.cell(row=r, column=4).value == 4.0
    r = labels["1ra Quincena - Horas Extras"]
    assert ws.cell(row=r, column=4).value == 0.5

    # 2da quincena: comunes=12.5h, extras=2.0h
    r = labels["2da Quincena - Horas Comunes"]
    assert ws.cell(row=r, column=4).value == 12.5
    r = labels["2da Quincena - Horas Extras"]
    assert ws.cell(row=r, column=4).value == 2.0

    # Totales mes: comunes=16.5h, extras=2.5h
    r = labels["Total Mes - Horas Comunes"]
    assert ws.cell(row=r, column=4).value == 16.5
    r = labels["Total Mes - Horas Extras"]
    assert ws.cell(row=r, column=4).value == 2.5

    # Orden esperado: cierra 1ra quincena, luego siguen dias de 2da quincena.
    row_q1_end = labels["1ra Quincena - Horas Extras"]
    row_day_18 = None
    row_day_20 = None
    for row in range(2, ws.max_row + 1):
        date_value = ws.cell(row=row, column=1).value
        text = str(date_value)
        if "2026-05-18" in text:
            row_day_18 = row
        if "2026-05-20" in text:
            row_day_20 = row
    assert row_day_18 is not None and row_day_18 > row_q1_end
    assert row_day_20 is not None and row_day_20 > row_q1_end


def test_liquidar_adds_status_legend_at_bottom(tmp_path: Path) -> None:
    daily_df = pd.DataFrame(
        {
            "ID de persona": ["20", "20"],
            "Nombre": ["Ana", "Ana"],
            "Fecha": [pd.Timestamp("2026-05-05").date(), pd.Timestamp("2026-05-06").date()],
            "Departamento": ["A", "A"],
            "Estado": ["Feriado", "Tardanza"],
            "Tramos trabajados": ["", "07:31-12:00 [07:30-12:00]"],
            "Minutos reales": [0, 269],
            "Minutos redondeados": [0, 240],
            "Minutos extra": [0, 0],
            "Horas extra": ["00:00", "00:00"],
            "Horas totales": ["00:00", "04:00"],
        }
    )
    monthly_df = pd.DataFrame(
        columns=[
            "ID de persona",
            "Nombre",
            "Dias trabajados",
            "Minutos totales",
            "Minutos extra",
            "Horas extra",
            "Horas totales",
        ]
    )
    inconsistencies_df = pd.DataFrame(
        columns=["ID de persona", "Nombre", "Fecha", "Tipo de inconsistencia", "Detalle"]
    )

    output = tmp_path / "reporte_liquidar_legend.xlsx"
    export_report(output, daily_df, monthly_df, inconsistencies_df)

    wb = load_workbook(output)
    ws = wb["Liquidar"]

    title_row = None
    for row in range(1, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == "Leyenda de estados":
            title_row = row
            break
    assert title_row is not None

    found_feriado = False
    for row in range(title_row + 1, ws.max_row + 1):
        for text_col in (2, 4):
            if ws.cell(row=row, column=text_col).value == "Feriado":
                box_col = text_col - 1
                assert _rgb(ws.cell(row=row, column=box_col)).endswith("A9DF8F")
                found_feriado = True
    assert found_feriado


def test_export_creates_incidencias_sheet_with_summary_and_detail(tmp_path: Path) -> None:
    daily_df = pd.DataFrame(
        {
            "ID de persona": ["20", "20", "30", "30"],
            "Nombre": ["Ana", "Ana", "Luis", "Luis"],
            "Fecha": [
                pd.Timestamp("2026-05-05").date(),
                pd.Timestamp("2026-05-06").date(),
                pd.Timestamp("2026-05-05").date(),
                pd.Timestamp("2026-05-06").date(),
            ],
            "Departamento": ["A", "A", "B", "B"],
            "Estado": ["Tarde", "Ausente", "Normal", "Tardanza"],
            "Tramos trabajados": ["07:31 - 12:00", "", "08:00 - 12:00", "08:05 - 12:00"],
            "Minutos reales": [269, 0, 240, 235],
            "Minutos redondeados": [270, 0, 240, 240],
            "Minutos extra": [0, 0, 0, 0],
            "Horas extra": ["00:00", "00:00", "00:00", "00:00"],
            "Horas totales": ["04:30", "00:00", "04:00", "04:00"],
        }
    )
    monthly_df = pd.DataFrame(
        columns=[
            "ID de persona",
            "Nombre",
            "Dias trabajados",
            "Minutos totales",
            "Minutos extra",
            "Horas extra",
            "Horas totales",
        ]
    )
    inconsistencies_df = pd.DataFrame(
        columns=["ID de persona", "Nombre", "Fecha", "Tipo de inconsistencia", "Detalle"]
    )

    output = tmp_path / "reporte_incidencias.xlsx"
    export_report(output, daily_df, monthly_df, inconsistencies_df)

    wb = load_workbook(output)
    assert "Incidencias" in wb.sheetnames
    ws = wb["Incidencias"]

    # Resumen (arriba)
    assert ws["A1"].value == "ID de persona"
    assert ws["B1"].value == "Nombre"
    assert ws["C1"].value == "Total tardanzas"
    assert ws["D1"].value == "Total ausencias"

    # Ana: 1 tarde, 1 ausente. Luis: 1 tardanza, 0 ausencias.
    summary = {}
    row = 2
    while ws.cell(row=row, column=1).value not in ("", None):
        emp_id = str(ws.cell(row=row, column=1).value)
        summary[emp_id] = {
            "tardanzas": ws.cell(row=row, column=3).value,
            "ausencias": ws.cell(row=row, column=4).value,
        }
        row += 1
    assert summary["20"]["tardanzas"] == 1
    assert summary["20"]["ausencias"] == 1
    assert summary["30"]["tardanzas"] == 1
    assert summary["30"]["ausencias"] == 0

    # Detalle debajo: fecha/estado/fichada de tarde+ausente.
    detail_header_row = row + 2
    assert ws.cell(row=detail_header_row, column=1).value == "ID de persona"
    assert ws.cell(row=detail_header_row, column=2).value == "Nombre"
    assert ws.cell(row=detail_header_row, column=3).value == "Fecha"
    assert ws.cell(row=detail_header_row, column=4).value == "Estado"
    assert ws.cell(row=detail_header_row, column=5).value == "Fichada"

    details = []
    current = detail_header_row + 1
    while ws.cell(row=current, column=1).value not in ("", None):
        fichada_value = ws.cell(row=current, column=5).value
        details.append(
            (
                str(ws.cell(row=current, column=1).value),
                str(ws.cell(row=current, column=4).value),
                "" if fichada_value in ("", None) else str(fichada_value),
            )
        )
        current += 1

    assert ("20", "Tarde", "07:31 - 12:00") in details
    assert ("20", "Ausente", "") in details
    assert ("30", "Tardanza", "08:05 - 12:00") in details
