from __future__ import annotations

from pathlib import Path

import pandas as pd

from services.exceptions import (
    append_manual_exceptions_to_file,
    ensure_exceptions_file,
    load_absences_template_file,
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


def test_load_absences_template_expands_date_range_and_skips_cancelled(tmp_path: Path) -> None:
    template_path = tmp_path / "date.xlsx"
    empleados_df = pd.DataFrame({"Id": ["20"], "Nombre": ["Ana"]})
    ausencias_df = pd.DataFrame(
        {
            "Legajo": ["20", "20"],
            "Tipo ausencia": ["VACACIONES", "ENFERMEDAD"],
            "Fecha desde": ["2026-05-01", "2026-05-10"],
            "Fecha hasta": ["2026-05-03", "2026-05-10"],
            "Estado": ["APROBADO", "CANCELADO"],
            "Observación": ["Invierno", "No aplica"],
        }
    )
    with pd.ExcelWriter(template_path, engine="openpyxl") as writer:
        empleados_df.to_excel(writer, sheet_name="Empleados", index=False)
        ausencias_df.to_excel(writer, sheet_name="Ausencias", index=False)

    results = load_absences_template_file(template_path)

    assert len(results) == 3
    dates = [str(item.exception_date) for item in results]
    assert dates == ["2026-05-01", "2026-05-02", "2026-05-03"]
    assert all(item.exception_type == "Vacaciones" for item in results)
    assert all(item.paid_day is True for item in results)


def test_load_exceptions_file_supports_date_template(tmp_path: Path) -> None:
    template_path = tmp_path / "date.xlsx"
    empleados_df = pd.DataFrame({"Id": ["20"], "Nombre": ["Ana"]})
    ausencias_df = pd.DataFrame(
        {
            "Legajo": ["20"],
            "Tipo ausencia": ["ART"],
            "Fecha desde": ["2026-05-20"],
            "Fecha hasta": ["2026-05-20"],
            "Estado": ["CARGADO"],
            "Observación": ["Accidente leve"],
        }
    )
    with pd.ExcelWriter(template_path, engine="openpyxl") as writer:
        empleados_df.to_excel(writer, sheet_name="Empleados", index=False)
        ausencias_df.to_excel(writer, sheet_name="Ausencias", index=False)

    results = load_exceptions_file(template_path)

    assert len(results) == 1
    assert results[0].employee_id == "20"
    assert str(results[0].exception_date) == "2026-05-20"
    assert results[0].exception_type == "Art"


def test_absences_template_sets_paid_day_only_when_aprobado(tmp_path: Path) -> None:
    template_path = tmp_path / "date.xlsx"
    empleados_df = pd.DataFrame({"Id": ["20", "30"], "Nombre": ["Ana", "Luis"]})
    ausencias_df = pd.DataFrame(
        {
            "Legajo": ["20", "30", "20"],
            "Tipo ausencia": ["VACACIONES", "ENFERMEDAD", "ART"],
            "Fecha desde": ["2026-06-01", "2026-06-02", "2026-06-03"],
            "Fecha hasta": ["2026-06-01", "2026-06-02", "2026-06-03"],
            "Estado": ["APROBADO", "PENDIENTE", "RECHAZADO"],
            "Observación": ["ok", "espera", "no"],
        }
    )
    with pd.ExcelWriter(template_path, engine="openpyxl") as writer:
        empleados_df.to_excel(writer, sheet_name="Empleados", index=False)
        ausencias_df.to_excel(writer, sheet_name="Ausencias", index=False)

    results = load_absences_template_file(template_path)
    by_type = {item.exception_type: item for item in results}
    assert by_type["Vacaciones"].paid_day is True
    assert by_type["Enfermedad"].paid_day is False
    assert by_type["Art"].paid_day is False
