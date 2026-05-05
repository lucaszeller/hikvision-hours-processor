from __future__ import annotations

from pathlib import Path

import pandas as pd

from services.exceptions import load_exceptions_file, parse_manual_exceptions


def test_parse_manual_exceptions_pipe_format() -> None:
    text = "20|2026-05-01|Feriado|Dia del trabajador\n|2026-05-25|Feriado|Patrio"

    results = parse_manual_exceptions(text)

    assert len(results) == 2
    assert results[0].employee_id == "20"
    assert results[0].exception_type == "Feriado"
    assert results[1].employee_id is None


def test_load_exceptions_csv(tmp_path: Path) -> None:
    csv_path = tmp_path / "exceptions.csv"
    pd.DataFrame(
        {
            "ID de persona": ["20", ""],
            "Fecha": ["2026-05-01", "2026-05-25"],
            "Tipo": ["Vacaciones", "Feriado"],
            "Detalle": ["Semana 1", "Patrio"],
        }
    ).to_csv(csv_path, index=False)

    results = load_exceptions_file(csv_path)

    assert len(results) == 2
    assert results[0].employee_id == "20"
    assert results[1].employee_id is None
    assert results[1].exception_type == "Feriado"
