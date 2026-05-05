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

    # Fila 2 = "Tarde" -> amarillo en toda la fila.
    assert _rgb(ws["A2"]).endswith("FFF59D")
    assert _rgb(ws["F2"]).endswith("FFF59D")

    # Fila 3 = "Ausente" -> rojo en toda la fila.
    assert _rgb(ws["A3"]).endswith("EF9A9A")
    assert _rgb(ws["J3"]).endswith("EF9A9A")
