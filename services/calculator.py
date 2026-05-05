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

INCONSISTENCY_COLUMNS = [
    "Legajo",
    "Nombre",
    "Fecha",
    "Tipo de inconsistencia",
    "Detalle",
]


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


def _remove_near_duplicates(
    timestamps: list,
    duplicate_window_minutes: int,
) -> tuple[list, list[timedelta]]:
    if not timestamps:
        return [], []

    clean = [timestamps[0]]
    duplicate_diffs: list[timedelta] = []

    for ts in timestamps[1:]:
        diff = ts - clean[-1]
        if diff <= timedelta(minutes=duplicate_window_minutes):
            duplicate_diffs.append(diff)
            continue
        clean.append(ts)

    return clean, duplicate_diffs


def process_punches(
    df: pd.DataFrame,
    duplicate_window_minutes: int = 2,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    rows_daily: list[dict] = []
    inconsistencies: list[dict] = []

    grouped = (
        df.sort_values("punch_datetime")
<<<<<<< HEAD
        .groupby(["employee_id", "employee_name", "work_date"], sort=True)
=======
         .groupby(["employee_id", "employee_name", "work_date"], sort=True)
>>>>>>> ef586063230aaba8617f425fce740814f056fd53
    )

    for (employee_id, employee_name, work_date), group in grouped:
        raw_timestamps = sorted([ts for ts in group["punch_datetime"] if pd.notna(ts)])
        timestamps, duplicate_diffs = _remove_near_duplicates(
            raw_timestamps, duplicate_window_minutes
        )

        for diff in duplicate_diffs:
            inconsistencies.append(
                {
                    "Legajo": employee_id,
                    "Nombre": employee_name,
                    "Fecha": work_date,
                    "Tipo de inconsistencia": "Duplicado cercano",
                    "Detalle": f"Fichadas con diferencia de {diff}.",
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
                    }
                )
                continue

            stretch = end - start
            day_total += stretch

            segments.append(
                f"Ingreso {start.strftime('%H:%M')} - "
                f"Salida {end.strftime('%H:%M')} "
                f"({_timedelta_to_hhmm(stretch)})"
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

    daily_df = pd.DataFrame(rows_daily)

    if not daily_df.empty:
        daily_df = daily_df.sort_values(["Nombre", "Fecha"])

        monthly_df = (
            daily_df.groupby(["Legajo", "Nombre"], as_index=False)["Minutos totales"]
            .sum()
            .sort_values(["Nombre"])
        )
        monthly_df["Horas mensuales"] = monthly_df["Minutos totales"].apply(
            lambda mins: _timedelta_to_hhmm(timedelta(minutes=int(mins)))
        )
        monthly_df = monthly_df[MONTHLY_COLUMNS]
        daily_df = daily_df[DAILY_COLUMNS]
    else:
        daily_df = pd.DataFrame(columns=DAILY_COLUMNS)
        monthly_df = pd.DataFrame(columns=MONTHLY_COLUMNS)

    inconsistencies_df = pd.DataFrame(inconsistencies, columns=INCONSISTENCY_COLUMNS)

    return daily_df, monthly_df, inconsistencies_df
