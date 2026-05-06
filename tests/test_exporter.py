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
