from __future__ import annotations

import unicodedata
from pathlib import Path

import pandas as pd

from domain.column_aliases import COLUMN_CANDIDATES

REQUIRED_COLUMNS = ("employee_id", "employee_name", "punch_datetime")


class ParsingError(Exception):
    pass


def _normalize(value: str) -> str:
    text = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    return " ".join(text.lower().replace("_", " ").split())


def _column_match_score(column: str, aliases: list[str]) -> int:
    normalized = _normalize(column)
    best = 0
    for alias in aliases:
        normalized_alias = _normalize(alias)
        if normalized == normalized_alias:
            return 100
        if normalized_alias in normalized:
            best = max(best, 80)
        if normalized in normalized_alias:
            best = max(best, 60)
    return best


def detect_column_mapping(columns: list[str]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for canonical_name, aliases in COLUMN_CANDIDATES.items():
        ranked = sorted(columns, key=lambda c: _column_match_score(c, aliases), reverse=True)
        if ranked and _column_match_score(ranked[0], aliases) > 0:
            mapping[canonical_name] = ranked[0]

    missing = [col for col in REQUIRED_COLUMNS if col not in mapping]
    if missing:
        raise ParsingError(
            "No se pudieron detectar columnas obligatorias: "
            + ", ".join(missing)
            + ". Revisá encabezados del Excel de Hikvision."
        )

    return mapping


def load_hikvision_excel(file_path: str | Path) -> pd.DataFrame:
    path = Path(file_path)
    if not path.exists():
        raise ParsingError(f"No existe el archivo: {path}")

    df = pd.read_excel(path)
    if df.empty:
        raise ParsingError("El archivo no contiene filas.")

    mapping = detect_column_mapping([str(c) for c in df.columns])

    canonical_df = pd.DataFrame()
    canonical_df["employee_id"] = df[mapping["employee_id"]].astype(str).str.strip()
    canonical_df["employee_name"] = df[mapping["employee_name"]].astype(str).str.strip()
    canonical_df["punch_datetime"] = pd.to_datetime(df[mapping["punch_datetime"]], errors="coerce")
    canonical_df["event_type"] = (
        df[mapping["event_type"]].astype(str).str.strip() if "event_type" in mapping else None
    )

    canonical_df = canonical_df.dropna(subset=["punch_datetime"])
    canonical_df = canonical_df[
        (canonical_df["employee_id"] != "") & (canonical_df["employee_name"] != "")
    ]

    if canonical_df.empty:
        raise ParsingError("No quedaron fichadas válidas luego de limpiar datos incompletos.")

    canonical_df["work_date"] = canonical_df["punch_datetime"].dt.date
    canonical_df["work_time"] = canonical_df["punch_datetime"].dt.time
    return canonical_df
