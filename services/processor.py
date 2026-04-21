from __future__ import annotations

from datetime import datetime
from pathlib import Path

from services.calculator import process_punches
from services.exporter import export_report
from services.parser import load_hikvision_excel


class ProcessorService:
    def process_file(self, input_path: str | Path, output_dir: str | Path | None = None) -> Path:
        source = Path(input_path)
        if output_dir is None:
            output_dir = source.parent

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = Path(output_dir) / f"reporte_horas_{stamp}.xlsx"

        punches_df = load_hikvision_excel(source)
        daily_df, monthly_df, inconsistencies_df = process_punches(punches_df)
        return export_report(output_path, daily_df, monthly_df, inconsistencies_df)
