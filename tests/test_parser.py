from datetime import datetime

import pandas as pd

from services.parser import detect_column_mapping


def test_detect_column_mapping_with_split_date_and_time() -> None:
    columns = ["No.", "User Name", "Date", "Time", "Device Name"]
    mapping = detect_column_mapping(columns)

    assert mapping["employee_id"] == "No."
    assert mapping["employee_name"] == "User Name"
    assert mapping["punch_date"] == "Date"
    assert mapping["punch_time"] == "Time"


def test_detect_column_mapping_with_standard_datetime() -> None:
    columns = ["Legajo", "Nombre", "Fecha y hora"]
    mapping = detect_column_mapping(columns)

    assert mapping["employee_id"] == "Legajo"
    assert mapping["employee_name"] == "Nombre"
    assert mapping["punch_datetime"] == "Fecha y hora"


def test_date_time_join_keeps_valid_datetime() -> None:
    date_part = pd.Series([datetime(2026, 4, 1), datetime(2026, 4, 2)])
    time_part = pd.Series(["08:00:00", "17:30:00"])

    joined = pd.to_datetime(date_part.dt.date.astype(str) + " " + time_part.astype(str), errors="coerce")

    assert joined.isna().sum() == 0
    assert joined.iloc[0].hour == 8
    assert joined.iloc[1].hour == 17
