from __future__ import annotations

from datetime import timedelta

import pandas as pd


def _timedelta_to_hhmm(delta: timedelta) -> str:
    total_seconds = int(delta.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes = remainder // 60
    return f"{hours:02d}:{minutes:02d}"


def process_punches(df: pd.DataFrame, duplicate_window_minutes: int = 2) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    rows_daily: list[dict] = []
    rows_monthly: list[dict] = []
    inconsistencies: list[dict] = []

    grouped = df.sort_values("punch_datetime").groupby(["employee_id", "employee_name", "work_date"], sort=True)

    for (employee_id, employee_name, work_date), group in grouped:
        timestamps = list(group["punch_datetime"])

        for i in range(1, len(timestamps)):
            diff = timestamps[i] - timestamps[i - 1]
            if diff <= timedelta(minutes=duplicate_window_minutes):
                inconsistencies.append(
                    {
                        "employee_id": employee_id,
                        "employee_name": employee_name,
                        "work_date": work_date,
                        "issue_type": "Duplicado cercano",
                        "details": f"Fichadas con diferencia de {diff}.",
                    }
                )

        if len(timestamps) % 2 != 0:
            inconsistencies.append(
                {
                    "employee_id": employee_id,
                    "employee_name": employee_name,
                    "work_date": work_date,
                    "issue_type": "Cantidad impar",
                    "details": "Quedó una fichada sin pareja (entrada/salida).",
                }
            )

        day_total = timedelta(0)
        segments: list[str] = []

        pair_count = len(timestamps) // 2
        for idx in range(pair_count):
            start = timestamps[idx * 2]
            end = timestamps[idx * 2 + 1]
            if end <= start:
                inconsistencies.append(
                    {
                        "employee_id": employee_id,
                        "employee_name": employee_name,
                        "work_date": work_date,
                        "issue_type": "Orden inválido",
                        "details": f"Salida {end} no es posterior a entrada {start}.",
                    }
                )
                continue

            stretch = end - start
            day_total += stretch
            segments.append(
                f"{start.strftime('%H:%M')} - {end.strftime('%H:%M')} ({_timedelta_to_hhmm(stretch)})"
            )

        rows_daily.append(
            {
                "employee_id": employee_id,
                "employee_name": employee_name,
                "work_date": work_date,
                "punch_count": len(timestamps),
                "segments": " | ".join(segments),
                "total_hours": _timedelta_to_hhmm(day_total),
                "total_minutes": int(day_total.total_seconds() // 60),
            }
        )

    daily_df = pd.DataFrame(rows_daily).sort_values(["employee_name", "work_date"])

    if not daily_df.empty:
        monthly_df = (
            daily_df.groupby(["employee_id", "employee_name"], as_index=False)["total_minutes"]
            .sum()
            .sort_values(["employee_name"])
        )
        monthly_df["monthly_total_hours"] = monthly_df["total_minutes"].apply(
            lambda mins: _timedelta_to_hhmm(timedelta(minutes=int(mins)))
        )
        rows_monthly = monthly_df.to_dict(orient="records")

    inconsistencies_df = pd.DataFrame(inconsistencies)
    monthly_df = pd.DataFrame(rows_monthly)

    return daily_df, monthly_df, inconsistencies_df
