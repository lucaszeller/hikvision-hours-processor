from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd

from services.calculator import process_punches
from services.exporter import export_report


def generate() -> None:
    base = Path(__file__).resolve().parent
    input_path = base / "sample_input.xlsx"
    output_path = base / "sample_output.xlsx"

    sample_df = pd.DataFrame(
        {
            "User ID": ["100", "100", "100", "100", "101", "101", "101"],
            "Employee Name": ["Ana Pérez", "Ana Pérez", "Ana Pérez", "Ana Pérez", "Luis Gómez", "Luis Gómez", "Luis Gómez"],
            "Date Time": [
                datetime(2026, 3, 1, 8, 0),
                datetime(2026, 3, 1, 12, 0),
                datetime(2026, 3, 1, 13, 0),
                datetime(2026, 3, 1, 17, 0),
                datetime(2026, 3, 1, 9, 0),
                datetime(2026, 3, 1, 12, 30),
                datetime(2026, 3, 1, 12, 31),
            ],
            "Event Type": ["Check In", "Check Out", "Check In", "Check Out", "Check In", "Check Out", "Check In"],
        }
    )

    sample_df.to_excel(input_path, index=False)

    normalized = pd.DataFrame(
        {
            "employee_id": sample_df["User ID"],
            "employee_name": sample_df["Employee Name"],
            "punch_datetime": pd.to_datetime(sample_df["Date Time"]),
            "work_date": pd.to_datetime(sample_df["Date Time"]).dt.date,
        }
    )
    daily_df, monthly_df, inconsistencies_df = process_punches(normalized)
    export_report(output_path, daily_df, monthly_df, inconsistencies_df)


if __name__ == "__main__":
    generate()
