from __future__ import annotations

from pathlib import Path

import pandas as pd

from services.exceptions import (
    append_manual_exceptions_to_file,
    ensure_exceptions_file,
    load_exceptions_file,
    parse_manual_exceptions,
)


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


def test_ensure_exceptions_file_adds_manual_column(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "feriados_nacionales_argentina_2026.xlsx"
    pd.DataFrame(
        {
            "ID de persona": [""],
            "Fecha": ["2026-05-25"],
            "Tipo": ["Feriado"],
            "Detalle": ["Patrio"],
        }
    ).to_excel(xlsx_path, index=False)

    ensure_exceptions_file(xlsx_path)
    df = pd.read_excel(xlsx_path)

    assert "Manual" in df.columns


def test_append_manual_exceptions_to_file_marks_manual_and_dedupes(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "feriados_nacionales_argentina_2026.xlsx"
    ensure_exceptions_file(xlsx_path)

    manual_text = "20|2026-05-01|Feriado|Dia del trabajador\n|2026-05-25|Feriado|Patrio"

    added_first = append_manual_exceptions_to_file(xlsx_path, manual_text)
    added_second = append_manual_exceptions_to_file(xlsx_path, manual_text)
    df = pd.read_excel(xlsx_path)

    assert added_first == 2
    assert added_second == 0
    assert len(df) == 2
    assert set(df["Manual"].astype(str)) == {"Si"}
