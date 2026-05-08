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


def test_export_creates_resumen_estudio_sheet(tmp_path: Path) -> None:
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
    assert "resumen-estudio" in wb.sheetnames
    ws = wb["resumen-estudio"]
    headers = [ws["A1"].value, ws["B1"].value, ws["C1"].value, ws["D1"].value]
    assert headers == ["Dia de semana", "Nombre del empleado", "Horas normales", "Horas extra"]
    assert ws["A2"].value == "Martes"
    assert ws["B2"].value == "Ana"
    assert ws["C2"].value == "04:00"
    assert ws["D2"].value == "00:30"


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
