from __future__ import annotations

import pandas as pd

from services.calculator import process_punches
from services.exceptions import WorkException


def _base_input() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "employee_id": ["20", "20"],
            "employee_name": ["De Carli Gonzalo", "De Carli Gonzalo"],
            "department": ["Produccion", "Produccion"],
            "schedule": ["Manana(07:30:00-12:00:00)", "Tarde(13:00:00-17:30:00)"],
            "work_date_raw": ["2026-04-01", "2026-04-01"],
            "entry_time_raw": ["07:30:56", "13:00:05"],
            "exit_time_raw": ["12:02:12", "17:32:45"],
        }
    )


def test_calculate_daily_and_monthly_hours_with_rounding() -> None:
    df = _base_input()

    daily, monthly, inconsistencies = process_punches(
        df,
        scheduled_minutes_by_employee={"20": 480},
    )

    assert len(daily) == 1
    assert daily.iloc[0]["Estado"] == "Normal"
    assert int(daily.iloc[0]["Minutos reales"]) == 544
    assert int(daily.iloc[0]["Minutos redondeados"]) == 540
    assert int(daily.iloc[0]["Minutos extra"]) == 0
    assert daily.iloc[0]["Horas extra"] == "00:00"
    assert daily.iloc[0]["Horas totales"] == "09:00"
    assert "07:30-12:02 [07:30-12:00]" in daily.iloc[0]["Tramos trabajados"]

    assert len(monthly) == 1
    assert int(monthly.iloc[0]["Dias trabajados"]) == 1
    assert int(monthly.iloc[0]["Minutos totales"]) == 540
    assert int(monthly.iloc[0]["Minutos extra"]) == 0
    assert monthly.iloc[0]["Horas extra"] == "00:00"
    assert monthly.iloc[0]["Horas totales"] == "09:00"

    assert inconsistencies.empty


def test_detects_required_inconsistencies_without_stopping_processing() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "", "30", "40"],
            "employee_name": ["Ana", "Bruno", "", "Carla"],
            "department": ["A", "B", "C", "D"],
            "schedule": ["", "", "", ""],
            "work_date_raw": ["2026-04-01", "fecha-mala", "2026-04-01", "2026-04-01"],
            "entry_time_raw": ["08:00", "08:00", "", "14:00"],
            "exit_time_raw": ["12:00", "17:00", "", "13:00"],
        }
    )

    daily, monthly, inconsistencies = process_punches(df)

    assert len(daily) == 1
    assert len(monthly) == 1
    assert not inconsistencies.empty

    issue_types = set(inconsistencies["Tipo de inconsistencia"])
    assert "ID de persona vacio" in issue_types
    assert "Nombre vacio" in issue_types
    assert "Fecha invalida" in issue_types
    assert "Ausente" in issue_types
    assert "Salida menor que entrada" in issue_types


def test_exception_replaces_missing_mark_inconsistency() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20"],
            "employee_name": ["Ana"],
            "department": ["A"],
            "schedule": [""],
            "work_date_raw": ["2026-05-01"],
            "entry_time_raw": [""],
            "exit_time_raw": [""],
        }
    )

    exceptions = [
        WorkException(
            employee_id="20",
            exception_date=pd.Timestamp("2026-05-01").date(),
            exception_type="Feriado",
            details="Dia del trabajador",
        )
    ]

    daily, monthly, inconsistencies = process_punches(df, exceptions=exceptions)

    assert len(daily) == 1
    assert len(monthly) == 1
    assert daily.iloc[0]["Estado"] == "Feriado"
    assert daily.iloc[0]["Horas totales"] == "00:00"
    assert daily.iloc[0]["Horas extra"] == "00:00"
    assert len(inconsistencies) == 1
    assert inconsistencies.iloc[0]["Tipo de inconsistencia"] == "Excepcion aplicada"
    assert "Feriado" in inconsistencies.iloc[0]["Detalle"]


def test_exception_from_template_is_rendered_even_without_source_row_for_date() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20"],
            "employee_name": ["Ana"],
            "department": ["A"],
            "schedule": [""],
            "work_date_raw": ["2026-05-01"],
            "entry_time_raw": ["07:30"],
            "exit_time_raw": ["12:00"],
        }
    )

    exceptions = [
        WorkException(
            employee_id="20",
            exception_date=pd.Timestamp("2026-05-03").date(),
            exception_type="Enfermedad",
            details="Licencia medica",
        )
    ]

    daily, monthly, inconsistencies = process_punches(
        df,
        exceptions=exceptions,
        scheduled_minutes_by_employee={"20": 240},
        scheduled_start_minute_by_employee={"20": 450},
    )

    # Debe existir el dia de enfermedad aunque no haya fila/fichada ese dia.
    by_date = {str(row["Fecha"]): row for _, row in daily.iterrows()}
    assert "2026-05-03" in by_date
    assert by_date["2026-05-03"]["Estado"] == "Enfermedad"
    assert by_date["2026-05-03"]["Horas totales"] == "00:00"

    # Mensual suma ambos dias (trabajado + excepcion con 0 horas).
    assert len(monthly) == 1
    assert int(monthly.iloc[0]["Dias trabajados"]) == 2
    # No debe quedar "Excepcion configurada sin uso".
    issue_types = set(inconsistencies["Tipo de inconsistencia"]) if not inconsistencies.empty else set()
    assert "Excepcion configurada sin uso" not in issue_types


def test_approved_template_exception_pays_day_pending_rejected_do_not() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20"],
            "employee_name": ["Ana"],
            "department": ["A"],
            "schedule": [""],
            "work_date_raw": ["2026-06-01"],
            "entry_time_raw": ["07:30"],
            "exit_time_raw": ["12:00"],
        }
    )

    exceptions = [
        WorkException(
            employee_id="20",
            exception_date=pd.Timestamp("2026-06-02").date(),
            exception_type="Vacaciones",
            details="Aprobado",
            paid_day=True,
        ),
        WorkException(
            employee_id="20",
            exception_date=pd.Timestamp("2026-06-03").date(),
            exception_type="Enfermedad",
            details="Pendiente",
            paid_day=False,
        ),
        WorkException(
            employee_id="20",
            exception_date=pd.Timestamp("2026-06-04").date(),
            exception_type="Art",
            details="Rechazado",
            paid_day=False,
        ),
    ]

    daily, monthly, _ = process_punches(
        df,
        exceptions=exceptions,
        scheduled_minutes_by_employee={"20": 540},
        scheduled_start_minute_by_employee={"20": 450},
    )

    by_date = {str(row["Fecha"]): row for _, row in daily.iterrows()}

    # Aprobado: dia pago con horas normales del horario (09:00).
    assert by_date["2026-06-02"]["Estado"] == "Vacaciones"
    assert int(by_date["2026-06-02"]["Minutos redondeados"]) == 540
    assert by_date["2026-06-02"]["Horas totales"] == "09:00"
    assert int(by_date["2026-06-02"]["Minutos extra"]) == 0

    # Pendiente/Rechazado: se muestra el estado pero no paga minutos/horas.
    assert by_date["2026-06-03"]["Estado"] == "Enfermedad"
    assert int(by_date["2026-06-03"]["Minutos redondeados"]) == 0
    assert by_date["2026-06-03"]["Horas totales"] == "00:00"

    assert by_date["2026-06-04"]["Estado"] == "Art"
    assert int(by_date["2026-06-04"]["Minutos redondeados"]) == 0
    assert by_date["2026-06-04"]["Horas totales"] == "00:00"

    assert len(monthly) == 1
    # 06-01 trabajado 04:30 (270) + 06-02 pago 09:00 (540) = 810
    assert int(monthly.iloc[0]["Minutos totales"]) == 810


def test_global_exception_paid_day_applies_to_all_known_employees() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "30"],
            "employee_name": ["Ana", "Luis"],
            "department": ["A", "B"],
            "schedule": ["", ""],
            "work_date_raw": ["2026-07-08", "2026-07-08"],
            "entry_time_raw": ["07:30", "08:00"],
            "exit_time_raw": ["12:00", "12:00"],
        }
    )

    exceptions = [
        WorkException(
            employee_id=None,  # legajo 1 en template
            exception_date=pd.Timestamp("2026-07-09").date(),
            exception_type="Feriado",
            details="Global",
            paid_day=True,
        )
    ]

    daily, _, _ = process_punches(
        df,
        exceptions=exceptions,
        scheduled_minutes_by_employee={"20": 540, "30": 480},
        scheduled_start_minute_by_employee={"20": 450, "30": 480},
    )

    rows_feriado = daily[daily["Fecha"].astype(str) == "2026-07-09"]
    assert len(rows_feriado) == 2

    by_id = {str(row["ID de persona"]): row for _, row in rows_feriado.iterrows()}
    assert by_id["20"]["Estado"] == "Feriado"
    assert int(by_id["20"]["Minutos redondeados"]) == 540
    assert by_id["20"]["Horas totales"] == "09:00"
    assert by_id["30"]["Estado"] == "Feriado"
    assert int(by_id["30"]["Minutos redondeados"]) == 480
    assert by_id["30"]["Horas totales"] == "08:00"


def test_non_working_employee_days_do_not_mark_absent_and_count_punch_as_overtime() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["118", "118", "118"],
            "employee_name": ["Zeller Lucas Ezequiel"] * 3,
            "department": ["A"] * 3,
            "schedule": [""] * 3,
            # lunes (trabaja), martes (no trabaja), jueves (no trabaja con fichada)
            "work_date_raw": ["2026-06-01", "2026-06-02", "2026-06-04"],
            "entry_time_raw": ["07:30", "", "07:30"],
            "exit_time_raw": ["12:00", "", "12:00"],
            "late_raw": ["", "", ""],
            "absent_raw": ["", "", ""],
        }
    )

    daily, _, _ = process_punches(
        df,
        scheduled_minutes_by_employee={"118": 270},
        scheduled_start_minute_by_employee={"118": 450},
        working_weekdays_by_employee={"118": {0, 2, 4}},  # lun/mie/vie
    )

    by_date = {str(row["Fecha"]): row for _, row in daily.iterrows()}
    assert "2026-06-02" not in by_date  # martes sin fichada: no ausente
    assert by_date["2026-06-01"]["Estado"] == "Normal"
    # jueves con fichada en dia no laborable: se toma como extra completa
    assert by_date["2026-06-04"]["Estado"] == "Normal"
    assert int(by_date["2026-06-04"]["Minutos redondeados"]) == 270
    assert int(by_date["2026-06-04"]["Minutos extra"]) == 270


def test_marks_late_and_absent_status_for_daily_cells() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "20"],
            "employee_name": ["Ana", "Ana"],
            "department": ["A", "A"],
            "schedule": ["", ""],
            "work_date_raw": ["2026-05-01", "2026-05-04"],
            "entry_time_raw": ["08:10", ""],
            "exit_time_raw": ["12:00", ""],
            "late_raw": ["00:10", ""],
            "absent_raw": ["", "01:00"],
        }
    )

    daily, _, _ = process_punches(df, scheduled_minutes_by_employee={"20": 240})
    status_by_date = {str(row["Fecha"]): row["Estado"] for _, row in daily.iterrows()}

    assert status_by_date["2026-05-01"] == "Tarde"
    assert status_by_date["2026-05-04"] == "Ausente"


def test_present_day_never_classified_as_ausente_even_with_absent_marker() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20"],
            "employee_name": ["Ana"],
            "department": ["A"],
            "schedule": [""],
            "work_date_raw": ["2026-05-05"],
            "entry_time_raw": ["08:28"],
            "exit_time_raw": ["12:00"],
            "late_raw": ["00:58"],
            "absent_raw": ["04:00"],
        }
    )

    daily, _, _ = process_punches(df, scheduled_minutes_by_employee={"20": 240})

    assert len(daily) == 1
    assert daily.iloc[0]["Estado"] == "Tarde"


def test_late_is_determined_by_employee_schedule_start_minute() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "20", "20", "20"],
            "employee_name": ["Ana", "Ana", "Ana", "Ana"],
            "department": ["A", "A", "A", "A"],
            "schedule": ["", "", "", ""],
            "work_date_raw": ["2026-05-06", "2026-05-06", "2026-05-07", "2026-05-07"],
            "entry_time_raw": ["07:30", "12:00", "07:31", "12:00"],
            "exit_time_raw": ["12:00", "16:30", "12:00", "16:30"],
            "late_raw": ["", "", "", ""],
            "absent_raw": ["", "", "", ""],
        }
    )

    daily, _, _ = process_punches(
        df,
        scheduled_minutes_by_employee={"20": 540},
        scheduled_start_minute_by_employee={"20": 450},
    )
    status_by_date = {str(row["Fecha"]): row["Estado"] for _, row in daily.iterrows()}

    assert status_by_date["2026-05-06"] == "Normal"  # 07:30 exacto
    assert status_by_date["2026-05-07"] == "Tarde"   # 07:31


def test_weekend_rule_only_counts_weekend_days_with_punches_and_as_overtime() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "20", "20", "20", "20", "20"],
            "employee_name": ["Ana"] * 6,
            "department": ["A"] * 6,
            "schedule": [""] * 6,
            # 2026-05-08 viernes, 2026-05-09 sabado, 2026-05-10 domingo
            "work_date_raw": [
                "2026-05-08",
                "2026-05-08",
                "2026-05-09",
                "2026-05-09",
                "2026-05-10",
                "2026-05-10",
            ],
            "entry_time_raw": ["07:30", "13:30", "08:00", "", "08:00", ""],
            "exit_time_raw": ["12:00", "18:00", "12:00", "", "12:00", ""],
            "late_raw": ["", "", "", "", "", ""],
            "absent_raw": ["", "", "", "", "", ""],
        }
    )

    daily, monthly, _ = process_punches(df, scheduled_minutes_by_employee={"20": 540})

    # Viernes + Sabado con fichadas + Domingo con fichadas.
    assert len(daily) == 3

    friday = daily[daily["Fecha"].astype(str) == "2026-05-08"].iloc[0]
    saturday = daily[daily["Fecha"].astype(str) == "2026-05-09"].iloc[0]
    sunday = daily[daily["Fecha"].astype(str) == "2026-05-10"].iloc[0]

    assert int(friday["Minutos redondeados"]) == 540
    assert int(friday["Minutos extra"]) == 0

    assert int(saturday["Minutos redondeados"]) == 240
    assert int(saturday["Minutos extra"]) == 240
    assert saturday["Horas extra"] == "04:00"
    assert int(sunday["Minutos redondeados"]) == 240
    assert int(sunday["Minutos extra"]) == 240
    assert sunday["Horas extra"] == "04:00"
    assert sunday["Estado"] == "Domingo"

    assert len(monthly) == 1
    assert int(monthly.iloc[0]["Minutos extra"]) == 480


def test_saturday_lateness_uses_fixed_0730_start_for_all() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "20", "30", "30"],
            "employee_name": ["Ana", "Ana", "Luis", "Luis"],
            "department": ["A", "A", "B", "B"],
            "schedule": ["", "", "", ""],
            # 2026-05-09 es sabado
            "work_date_raw": ["2026-05-09", "2026-05-09", "2026-05-09", "2026-05-09"],
            "entry_time_raw": ["07:30", "12:00", "07:31", "12:00"],
            "exit_time_raw": ["11:30", "16:00", "11:30", "16:00"],
            "late_raw": ["", "", "", ""],
            "absent_raw": ["", "", "", ""],
        }
    )

    daily, _, _ = process_punches(
        df,
        scheduled_minutes_by_employee={"20": 540, "30": 540},
        # Aunque el horario de inicio del empleado sea distinto,
        # en sabado debe usar siempre 07:30.
        scheduled_start_minute_by_employee={"20": 450, "30": 600},
    )

    status_by_id = {str(row["ID de persona"]): row["Estado"] for _, row in daily.iterrows()}
    assert status_by_id["20"] == "Normal"  # 07:30 exacto
    assert status_by_id["30"] == "Tarde"   # 07:31


def test_empty_input_returns_empty_dataframes_with_expected_columns() -> None:
    df = pd.DataFrame(
        columns=[
            "employee_id",
            "employee_name",
            "department",
            "schedule",
            "work_date_raw",
            "entry_time_raw",
            "exit_time_raw",
        ]
    )

    daily, monthly, inconsistencies = process_punches(df)

    assert list(daily.columns) == [
        "ID de persona",
        "Nombre",
        "Fecha",
        "Departamento",
        "Estado",
        "Tramos trabajados",
        "Minutos reales",
        "Minutos redondeados",
        "Minutos extra",
        "Horas extra",
        "Horas totales",
    ]
    assert list(monthly.columns) == [
        "ID de persona",
        "Nombre",
        "Dias trabajados",
        "Minutos totales",
        "Minutos extra",
        "Horas extra",
        "Horas totales",
    ]
    assert list(inconsistencies.columns) == [
        "ID de persona",
        "Nombre",
        "Fecha",
        "Tipo de inconsistencia",
        "Detalle",
    ]
    assert daily.empty and monthly.empty and inconsistencies.empty


def test_split_schedule_two_punches_uses_theoretical_segments_without_missing_punch_inconsistencies() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "20"],
            "employee_name": ["Ana", "Ana"],
            "department": ["A", "A"],
            "schedule": ["", ""],
            "work_date_raw": ["2026-05-05", "2026-05-05"],
            "entry_time_raw": ["07:45", ""],
            "exit_time_raw": ["", "17:40"],
            "late_raw": ["", ""],
            "absent_raw": ["", ""],
        }
    )

    split_profiles = {
        "20": {
            "is_continuous": False,  # Horario corrido = NO
            "morning_in_minute": 450,
            "morning_out_minute": 720,
            "afternoon_in_minute": 780,
            "afternoon_out_minute": 1050,
        }
    }

    daily, monthly, inconsistencies = process_punches(
        df,
        scheduled_minutes_by_employee={"20": 540},
        scheduled_start_minute_by_employee={"20": 450},
        split_schedule_by_employee=split_profiles,
    )

    assert len(daily) == 1
    assert int(daily.iloc[0]["Minutos redondeados"]) == 540
    assert int(daily.iloc[0]["Minutos extra"]) == 0
    assert daily.iloc[0]["Estado"] == "Tarde"
    tramos = str(daily.iloc[0]["Tramos trabajados"])
    assert "07:45-12:00 [07:30-12:00]" in tramos
    assert "13:00-17:40 [13:00-17:30]" in tramos
    assert " ||| " in tramos

    assert len(monthly) == 1
    assert int(monthly.iloc[0]["Minutos totales"]) == 540
    assert int(monthly.iloc[0]["Minutos extra"]) == 0

    if not inconsistencies.empty:
        issue_types = set(inconsistencies["Tipo de inconsistencia"])
        assert "Falta Registro de entrada" not in issue_types
        assert "Falta Registro de salida" not in issue_types


def test_split_schedule_single_continuous_row_is_not_counted_as_continuous_span() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20"],
            "employee_name": ["Ana"],
            "department": ["A"],
            "schedule": [""],
            "work_date_raw": ["2026-05-06"],
            "entry_time_raw": ["07:30"],
            "exit_time_raw": ["18:30"],  # sin fichadas intermedias
            "late_raw": [""],
            "absent_raw": [""],
        }
    )

    split_profiles = {
        "20": {
            "is_continuous": False,  # Horario corrido = NO
            "morning_in_minute": 450,
            "morning_out_minute": 720,
            "afternoon_in_minute": 780,
            "afternoon_out_minute": 1050,
        }
    }

    daily, monthly, _ = process_punches(
        df,
        scheduled_minutes_by_employee={"20": 540},
        scheduled_start_minute_by_employee={"20": 450},
        split_schedule_by_employee=split_profiles,
    )

    assert len(daily) == 1
    # Debe tomar tramos teoricos (9h), no 11h continuas.
    assert int(daily.iloc[0]["Minutos redondeados"]) == 540
    assert int(daily.iloc[0]["Minutos extra"]) == 0
    assert daily.iloc[0]["Horas totales"] == "09:00"
    assert daily.iloc[0]["Horas extra"] == "00:00"

    assert len(monthly) == 1
    assert int(monthly.iloc[0]["Minutos totales"]) == 540
    assert int(monthly.iloc[0]["Minutos extra"]) == 0


def test_saturday_split_schedule_uses_single_morning_shift() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "20"],
            "employee_name": ["Ana", "Ana"],
            "department": ["A", "A"],
            "schedule": ["", ""],
            "work_date_raw": ["2026-05-09", "2026-05-09"],
            "entry_time_raw": ["07:41", ""],
            "exit_time_raw": ["", "11:36"],
            "late_raw": ["", ""],
            "absent_raw": ["", ""],
        }
    )

    split_profiles = {
        "20": {
            "is_continuous": False,
            "morning_in_minute": 450,
            "morning_out_minute": 690,
            "afternoon_in_minute": 810,
            "afternoon_out_minute": 1080,
        }
    }

    daily, _, inconsistencies = process_punches(
        df,
        scheduled_minutes_by_employee={"20": 240},
        scheduled_start_minute_by_employee={"20": 450},
        split_schedule_by_employee=split_profiles,
    )

    assert len(daily) == 1
    tramos = str(daily.iloc[0]["Tramos trabajados"])
    assert tramos == "07:41-11:36 [07:30-11:30]"
    assert int(daily.iloc[0]["Minutos redondeados"]) == 210
    assert int(daily.iloc[0]["Minutos extra"]) == 210
    assert inconsistencies.empty


def test_mixed_theoretical_and_normal_segments_do_not_fail_with_nat_in_display_columns() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "20", "30"],
            "employee_name": ["Ana", "Ana", "Luis"],
            "department": ["A", "A", "B"],
            "schedule": ["", "", "Manana(08:00:00-12:00:00)"],
            "work_date_raw": ["2026-05-07", "2026-05-07", "2026-05-07"],
            "entry_time_raw": ["07:40", "", "08:00"],
            "exit_time_raw": ["", "17:35", "12:00"],
            "late_raw": ["", "", ""],
            "absent_raw": ["", "", ""],
        }
    )

    split_profiles = {
        "20": {
            "is_continuous": False,
            "morning_in_minute": 450,
            "morning_out_minute": 720,
            "afternoon_in_minute": 780,
            "afternoon_out_minute": 1050,
        }
    }

    daily, monthly, inconsistencies = process_punches(
        df,
        scheduled_minutes_by_employee={"20": 540, "30": 240},
        scheduled_start_minute_by_employee={"20": 450, "30": 480},
        split_schedule_by_employee=split_profiles,
    )

    assert len(daily) == 2
    by_id = {str(row["ID de persona"]): row for _, row in daily.iterrows()}
    assert "20" in by_id and "30" in by_id
    assert "07:40" in str(by_id["20"]["Tramos trabajados"])
    assert "08:00-12:00 [08:00-12:00]" in str(by_id["30"]["Tramos trabajados"])
    assert len(monthly) == 2
    # No debe explotar por NaTType/strftime en armado de tramos.
    assert inconsistencies is not None


def test_continuous_schedule_without_overtime_shows_real_then_expected_single_range() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["77"],
            "employee_name": ["Viajante"],
            "department": ["Ventas"],
            "schedule": [""],
            "work_date_raw": ["2026-05-11"],
            "entry_time_raw": ["07:32"],
            "exit_time_raw": ["18:31"],
            "late_raw": [""],
            "absent_raw": [""],
        }
    )

    profiles = {
        "77": {
            "is_continuous": True,
            "morning_in_minute": 450,   # 07:30
            "morning_out_minute": None,
            "afternoon_in_minute": None,
            "afternoon_out_minute": 1140,  # 19:00
        }
    }

    daily, monthly, _ = process_punches(
        df,
        scheduled_minutes_by_employee={"77": 690},  # 11:30
        scheduled_start_minute_by_employee={"77": 450},
        split_schedule_by_employee=profiles,
    )

    assert len(daily) == 1
    assert int(daily.iloc[0]["Minutos extra"]) == 0
    assert daily.iloc[0]["Tramos trabajados"] == "07:32-18:31 [07:30-19:00]"
    assert len(monthly) == 1


def test_overtime_is_counted_in_30_min_blocks_per_segment() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "20"],
            "employee_name": ["Ana", "Ana"],
            "department": ["A", "A"],
            "schedule": ["Manana(06:30:00-12:00:00)", "Manana(06:30:00-12:00:00)"],
            "work_date_raw": ["2026-05-12", "2026-05-13"],
            "entry_time_raw": ["06:30", "06:30"],
            "exit_time_raw": ["12:30", "12:26"],
            "late_raw": ["", ""],
            "absent_raw": ["", ""],
        }
    )

    daily, _, _ = process_punches(
        df,
        scheduled_minutes_by_employee={"20": 330},
        scheduled_start_minute_by_employee={"20": 390},
    )

    by_date = {str(row["Fecha"]): row for _, row in daily.iterrows()}
    assert int(by_date["2026-05-12"]["Minutos extra"]) == 30  # +30 min => 0.5h extra
    assert by_date["2026-05-12"]["Horas extra"] == "00:30"
    assert int(by_date["2026-05-13"]["Minutos extra"]) == 0   # +26 min => no extra
    assert by_date["2026-05-13"]["Horas extra"] == "00:00"


def test_overtime_rounding_is_applied_separately_to_entry_and_exit() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "20"],
            "employee_name": ["Ana", "Ana"],
            "department": ["A", "A"],
            "schedule": ["Tarde(13:30:00-18:00:00)", "Tarde(13:30:00-18:00:00)"],
            "work_date_raw": ["2026-05-14", "2026-05-15"],
            "entry_time_raw": ["13:10", "13:25"],
            "exit_time_raw": ["18:20", "19:26"],
            "late_raw": ["", ""],
            "absent_raw": ["", ""],
        }
    )

    daily, _, _ = process_punches(
        df,
        scheduled_minutes_by_employee={"20": 270},
        scheduled_start_minute_by_employee={"20": 810},
    )

    by_date = {str(row["Fecha"]): row for _, row in daily.iterrows()}
    # 13:10-18:20 => 20 min antes + 20 min despues => 0 + 0 = 0 extra
    assert int(by_date["2026-05-14"]["Minutos extra"]) == 0
    # 13:25-19:26 => 5 min antes + 86 min despues => 0 + 60 = 60 extra
    assert int(by_date["2026-05-15"]["Minutos extra"]) == 60
    assert by_date["2026-05-15"]["Horas extra"] == "01:00"


def test_entry_only_with_schedule_is_included_in_daily_using_theoretical_exit() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["93"],
            "employee_name": ["Suarez Elina"],
            "department": ["CM"],
            "schedule": ["Elina(08:00:00-13:00:00)"],
            "work_date_raw": ["2026-05-08"],
            "entry_time_raw": ["06:54"],
            "exit_time_raw": [""],
            "late_raw": [""],
            "absent_raw": [""],
        }
    )

    profiles = {
        "93": {
            "is_continuous": True,
            "morning_in_minute": 480,
            "morning_out_minute": 780,
            "afternoon_in_minute": None,
            "afternoon_out_minute": None,
        }
    }

    daily, _, inconsistencies = process_punches(
        df,
        scheduled_minutes_by_employee={"93": 300},
        scheduled_start_minute_by_employee={"93": 480},
        split_schedule_by_employee=profiles,
    )

    assert len(daily) == 1
    row = daily.iloc[0]
    assert str(row["ID de persona"]) == "93"
    assert "06:54-13:00 [08:00-13:00]" in str(row["Tramos trabajados"])
    assert int(row["Minutos redondeados"]) == 360

    if not inconsistencies.empty:
        types = set(inconsistencies["Tipo de inconsistencia"])
        assert "Falta Registro de salida" not in types


def test_presente_exception_pays_employee_scheduled_hours_and_sets_presente_status() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20"],
            "employee_name": ["Ana"],
            "department": ["A"],
            "schedule": [""],
            "work_date_raw": ["2026-05-20"],
            "entry_time_raw": [""],
            "exit_time_raw": [""],
            "late_raw": [""],
            "absent_raw": [""],
        }
    )

    exceptions = [
        WorkException(
            employee_id="20",
            exception_date=pd.Timestamp("2026-05-20").date(),
            exception_type="Presente",
            details="Carga manual",
            paid_day=False,
        )
    ]

    daily, monthly, _ = process_punches(
        df,
        exceptions=exceptions,
        scheduled_minutes_by_employee={"20": 300},
        scheduled_start_minute_by_employee={"20": 480},
    )

    assert len(daily) == 1
    assert daily.iloc[0]["Estado"] == "Presente"
    assert int(daily.iloc[0]["Minutos redondeados"]) == 300
    assert daily.iloc[0]["Horas totales"] == "05:00"
    assert len(monthly) == 1
    assert int(monthly.iloc[0]["Minutos totales"]) == 300


def test_presente_exception_ignores_existing_punches_and_uses_scheduled_hours() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20"],
            "employee_name": ["Ana"],
            "department": ["A"],
            "schedule": ["Manana(08:00:00-13:00:00)"],
            "work_date_raw": ["2026-05-21"],
            "entry_time_raw": ["07:10"],
            "exit_time_raw": ["18:50"],
            "late_raw": [""],
            "absent_raw": [""],
        }
    )

    exceptions = [
        WorkException(
            employee_id="20",
            exception_date=pd.Timestamp("2026-05-21").date(),
            exception_type="Presente",
            details="Ajuste",
            paid_day=False,
        )
    ]

    daily, monthly, _ = process_punches(
        df,
        exceptions=exceptions,
        scheduled_minutes_by_employee={"20": 300},
        scheduled_start_minute_by_employee={"20": 480},
    )

    assert len(daily) == 1
    row = daily.iloc[0]
    assert row["Estado"] == "Presente"
    assert row["Tramos trabajados"] == ""
    assert int(row["Minutos redondeados"]) == 300
    assert int(row["Minutos extra"]) == 0
    assert row["Horas totales"] == "05:00"
    assert row["Horas extra"] == "00:00"

    assert len(monthly) == 1
    assert int(monthly.iloc[0]["Minutos totales"]) == 300
    assert int(monthly.iloc[0]["Minutos extra"]) == 0


def test_non_presente_exception_has_priority_over_presente_on_saturday() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20"],
            "employee_name": ["Ana"],
            "department": ["A"],
            "schedule": [""],
            "work_date_raw": ["2026-05-23"],  # sabado
            "entry_time_raw": [""],
            "exit_time_raw": [""],
            "late_raw": [""],
            "absent_raw": [""],
        }
    )

    exceptions = [
        WorkException(
            employee_id="20",
            exception_date=pd.Timestamp("2026-05-23").date(),
            exception_type="Presente",
            details="General",
            paid_day=False,
        ),
        WorkException(
            employee_id="20",
            exception_date=pd.Timestamp("2026-05-23").date(),
            exception_type="Accidente de trabajo",
            details="ART",
            paid_day=True,
        ),
    ]

    daily, _, _ = process_punches(
        df,
        exceptions=exceptions,
        scheduled_minutes_by_employee={"20": 240},
        scheduled_start_minute_by_employee={"20": 450},
    )

    assert len(daily) == 1
    row = daily.iloc[0]
    assert row["Estado"] == "Accidente De Trabajo"
    assert int(row["Minutos redondeados"]) == 0
    assert row["Horas totales"] == "00:00"


def test_saturday_non_presente_exception_ignores_punches_when_presente_exists() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "20"],
            "employee_name": ["Ana", "Ana"],
            "department": ["A", "A"],
            "schedule": ["", ""],
            "work_date_raw": ["2026-05-23", "2026-05-23"],  # sabado
            "entry_time_raw": ["07:41", ""],
            "exit_time_raw": ["", "11:36"],
            "late_raw": ["", ""],
            "absent_raw": ["", ""],
        }
    )

    exceptions = [
        WorkException(
            employee_id="20",
            exception_date=pd.Timestamp("2026-05-23").date(),
            exception_type="Presente",
            details="General",
            paid_day=False,
        ),
        WorkException(
            employee_id="20",
            exception_date=pd.Timestamp("2026-05-23").date(),
            exception_type="Accidente de trabajo",
            details="ART",
            paid_day=True,
        ),
    ]

    daily, _, _ = process_punches(
        df,
        exceptions=exceptions,
        scheduled_minutes_by_employee={"20": 240},
        scheduled_start_minute_by_employee={"20": 450},
    )

    assert len(daily) == 1
    row = daily.iloc[0]
    assert row["Estado"] == "Accidente De Trabajo"
    assert row["Tramos trabajados"] == ""
    assert int(row["Minutos redondeados"]) == 0
    assert int(row["Minutos extra"]) == 0


def test_saturday_any_exception_zeroes_worked_hours() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "20"],
            "employee_name": ["Ana", "Ana"],
            "department": ["A", "A"],
            "schedule": ["", ""],
            "work_date_raw": ["2026-05-23", "2026-05-23"],  # sabado
            "entry_time_raw": ["07:41", ""],
            "exit_time_raw": ["", "11:36"],
            "late_raw": ["", ""],
            "absent_raw": ["", ""],
        }
    )

    exceptions = [
        WorkException(
            employee_id="20",
            exception_date=pd.Timestamp("2026-05-23").date(),
            exception_type="Vacaciones",
            details="Carga",
            paid_day=True,
        ),
    ]

    daily, _, _ = process_punches(
        df,
        exceptions=exceptions,
        scheduled_minutes_by_employee={"20": 240},
        scheduled_start_minute_by_employee={"20": 450},
    )

    assert len(daily) == 1
    row = daily.iloc[0]
    assert row["Estado"] == "Vacaciones"
    assert int(row["Minutos redondeados"]) == 0
    assert int(row["Minutos extra"]) == 0
    assert row["Horas totales"] == "00:00"


def test_flexible_attendance_employee_does_not_get_late_or_absent_status() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["107", "107"],
            "employee_name": ["Juarez Paulina", "Juarez Paulina"],
            "department": ["CM", "CM"],
            "schedule": ["Elina(08:00:00-13:00:00)", "Elina(08:00:00-13:00:00)"],
            "work_date_raw": ["2026-05-22", "2026-05-23"],
            "entry_time_raw": ["09:30", ""],
            "exit_time_raw": ["13:00", ""],
            "late_raw": ["01:30", ""],
            "absent_raw": ["", "01:00"],
        }
    )

    daily, _, inconsistencies = process_punches(
        df,
        scheduled_minutes_by_employee={"107": 300},
        scheduled_start_minute_by_employee={"107": 480},
        flexible_attendance_employee_ids={"107"},
    )

    by_date = {str(row["Fecha"]): row["Estado"] for _, row in daily.iterrows()}
    assert by_date["2026-05-22"] == "Normal"
    assert "2026-05-23" not in by_date  # sin fichadas: no generar ausente

    if not inconsistencies.empty:
        issue_types = set(inconsistencies["Tipo de inconsistencia"])
        assert "Ausente" not in issue_types


def test_flexible_attendance_employee_never_gets_overtime_minutes() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["107"],
            "employee_name": ["Juarez Paulina"],
            "department": ["CM"],
            "schedule": ["Paulina(09:00:00-15:00:00)"],
            "work_date_raw": ["2026-05-24"],  # domingo
            "entry_time_raw": ["08:00"],
            "exit_time_raw": ["16:30"],
            "late_raw": [""],
            "absent_raw": [""],
        }
    )

    daily, monthly, _ = process_punches(
        df,
        scheduled_minutes_by_employee={"107": 360},
        scheduled_start_minute_by_employee={"107": 540},
        flexible_attendance_employee_ids={"107"},
    )

    assert len(daily) == 1
    assert int(daily.iloc[0]["Minutos redondeados"]) == 510
    assert int(daily.iloc[0]["Minutos extra"]) == 0
    assert daily.iloc[0]["Horas extra"] == "00:00"

    assert len(monthly) == 1
    assert int(monthly.iloc[0]["Minutos extra"]) == 0
