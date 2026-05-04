from __future__ import annotations

from datetime import timedelta

import pandas as pd

DAILY_COLUMNS = [
    "Legajo",
    "Nombre",
    "Fecha",
    "Cantidad de fichadas",
    "Marcaciones (I/S)",
    "Tramos trabajados",
    "Horas totales",
    "Minutos totales",
]
MONTHLY_COLUMNS = ["Legajo", "Nombre", "Minutos totales", "Horas mensuales"]
INCONSISTENCY_COLUMNS = ["Legajo", "Nombre", "Fecha", "Tipo de inconsistencia", "Detalle"]

=======
 main

def _timedelta_to_hhmm(delta: timedelta) -> str:
    total_seconds = int(delta.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes = remainder // 60
    return f"{hours:02d}:{minutes:02d}"


def _build_marks_text(timestamps: list) -> str:
    marks: list[str] = []
    for idx, ts in enumerate(timestamps, start=1):
        tipo = "Ingreso" if idx % 2 != 0 else "Salida"
        marks.append(f"{tipo} {ts.strftime('%H:%M')}")
    return " | ".join(marks)


 main
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
                        "Legajo": employee_id,
                        "Nombre": employee_name,
                        "Fecha": work_date,
                        "Tipo de inconsistencia": "Duplicado cercano",
                        "Detalle": f"Fichadas con diferencia de {diff}.",

                        "employee_id": employee_id,
                        "employee_name": employee_name,
                        "work_date": work_date,
                        "issue_type": "Duplicado cercano",
                        "details": f"Fichadas con diferencia de {diff}.",
 main
                    }
                )

        if len(timestamps) % 2 != 0:
            inconsistencies.append(
                {
                    "Legajo": employee_id,
                    "Nombre": employee_name,
                    "Fecha": work_date,
                    "Tipo de inconsistencia": "Cantidad impar",
                    "Detalle": "Quedó una fichada sin pareja (ingreso/salida).",

                    "employee_id": employee_id,
                    "employee_name": employee_name,
                    "work_date": work_date,
                    "issue_type": "Cantidad impar",
                    "details": "Quedó una fichada sin pareja (entrada/salida).",
 main
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
                        "Legajo": employee_id,
                        "Nombre": employee_name,
                        "Fecha": work_date,
                        "Tipo de inconsistencia": "Orden inválido",
                        "Detalle": f"Salida {end} no es posterior a ingreso {start}.",

                        "employee_id": employee_id,
                        "employee_name": employee_name,
                        "work_date": work_date,
                        "issue_type": "Orden inválido",
                        "details": f"Salida {end} no es posterior a entrada {start}.",
main
                    }
                )
                continue

            stretch = end - start
            day_total += stretch
            segments.append(
                f"Ingreso {start.strftime('%H:%M')} - Salida {end.strftime('%H:%M')} ({_timedelta_to_hhmm(stretch)})"

                f"{start.strftime('%H:%M')} - {end.strftime('%H:%M')} ({_timedelta_to_hhmm(stretch)})"
 main
            )

        rows_daily.append(
            {
                "Legajo": employee_id,
                "Nombre": employee_name,
                "Fecha": work_date,
                "Cantidad de fichadas": len(timestamps),
                "Marcaciones (I/S)": _build_marks_text(timestamps),
                "Tramos trabajados": " | ".join(segments),
                "Horas totales": _timedelta_to_hhmm(day_total),
                "Minutos totales": int(day_total.total_seconds() // 60),
            }
        )

    daily_df = pd.DataFrame(rows_daily, columns=DAILY_COLUMNS)
    if not daily_df.empty:
        daily_df = daily_df.sort_values(["Nombre", "Fecha"])
        monthly_df = (
            daily_df.groupby(["Legajo", "Nombre"], as_index=False)["Minutos totales"]
            .sum()
            .sort_values(["Nombre"])
        )
        monthly_df["Horas mensuales"] = monthly_df["Minutos totales"].apply(
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
 main
            lambda mins: _timedelta_to_hhmm(timedelta(minutes=int(mins)))
        )
        rows_monthly = monthly_df.to_dict(orient="records")


    inconsistencies_df = pd.DataFrame(inconsistencies, columns=INCONSISTENCY_COLUMNS)
    monthly_df = pd.DataFrame(rows_monthly, columns=MONTHLY_COLUMNS)

    inconsistencies_df = pd.DataFrame(inconsistencies)
    monthly_df = pd.DataFrame(rows_monthly)
 main

    return daily_df, monthly_df, inconsistencies_df
