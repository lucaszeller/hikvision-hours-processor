from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Callable

from services.calculator import process_punches
from services.exporter import export_report
from services.parser import load_hikvision_excel


class ProcessorService:
    def process_file(
        self,
        input_path: str | Path,
        output_dir: str | Path | None = None,
        progress_callback: Callable[[int, str], None] | None = None,
    ) -> Path:
        source = Path(input_path)
        if output_dir is None:
            output_dir = source.parent

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = Path(output_dir) / f"reporte_horas_{stamp}.xlsx"

        if progress_callback:
            progress_callback(10, "Leyendo fichadas del archivo...")
        punches_df = load_hikvision_excel(source)

        if progress_callback:
            progress_callback(55, "Calculando horas e inconsistencias...")
        daily_df, monthly_df, inconsistencies_df = process_punches(punches_df)

        if progress_callback:
            progress_callback(85, "Generando reporte final...")
        report_path = export_report(output_path, daily_df, monthly_df, inconsistencies_df)

        if progress_callback:
            progress_callback(100, "Proceso completado.")
        return report_path
