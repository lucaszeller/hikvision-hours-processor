from __future__ import annotations

import os
import queue
import threading
import time
from datetime import datetime, timedelta
import math
from pathlib import Path
import unicodedata

import customtkinter as ctk
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

from services.exceptions import (
    ExceptionConfigError,
    append_manual_exceptions_to_file,
    ensure_exceptions_file,
)
from services.processor import ProcessorService
from services.parser import load_hikvision_excel

try:
    from tkcalendar import Calendar
except Exception:  # pragma: no cover - optional dependency at runtime
    Calendar = None

DEFAULT_EXCEPTIONS_FILENAME = "feriados_nacionales_argentina_2026.xlsx"

COLOR_BG = ("#EEF3F8", "#0A1220")
COLOR_PANEL = ("#FFFFFF", "#111827")
COLOR_PANEL_SOFT = ("#F8FAFC", "#0F172A")
COLOR_BORDER = ("#D7E1ED", "#243244")
COLOR_TEXT = ("#0F172A", "#E5E7EB")
COLOR_TEXT_MUTED = ("#475569", "#93A4B8")
COLOR_PRIMARY = ("#0E7490", "#22D3EE")
COLOR_PRIMARY_HOVER = ("#155E75", "#06B6D4")
COLOR_SECONDARY_HOVER = ("#E2E8F0", "#1E293B")
COLOR_TAB_ACTIVE = ("#E0F2FE", "#10243E")
COLOR_TAB_TEXT_ACTIVE = ("#0C4A6E", "#67E8F9")
COLOR_STATUS_READY_BG = ("#DCFCE7", "#133124")
COLOR_STATUS_READY_TEXT = ("#166534", "#86EFAC")
COLOR_STATUS_WORKING_BG = ("#FEF3C7", "#3F3208")
COLOR_STATUS_WORKING_TEXT = ("#92400E", "#FDE68A")
COLOR_STATUS_WORKING_BG_ALT = ("#FDE68A", "#4A3A0A")
COLOR_STATUS_ERROR_BG = ("#FEE2E2", "#3D1111")
COLOR_STATUS_ERROR_TEXT = ("#991B1B", "#FCA5A5")


class HikvisionApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Hikvision Hours Processor")
        self.geometry("1220x760")
        self.minsize(1080, 640)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("dark-blue")
        self.configure(fg_color=COLOR_BG)

        self.font_title = ctk.CTkFont(family="Bahnschrift", size=26, weight="bold")
        self.font_section = ctk.CTkFont(family="Bahnschrift", size=20, weight="bold")
        self.font_subtitle = ctk.CTkFont(family="Bahnschrift", size=13)
        self.font_body = ctk.CTkFont(family="Bahnschrift", size=12)
        self.font_button = ctk.CTkFont(family="Bahnschrift", size=13, weight="bold")
        self.font_mono = ctk.CTkFont(family="Consolas", size=12)

        self.processor = ProcessorService()
        self.selected_file: Path | None = None
        self.exceptions_file: Path | None = None
        self.output_file: Path | None = None

        self._worker_thread: threading.Thread | None = None
        self._progress_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._started_at: float | None = None

        self.current_page = "report"

        self.log_box: ctk.CTkTextbox | None = None
        self.exceptions_file_entry: ctk.CTkEntry | None = None
        self.exceptions_manual_box: ctk.CTkTextbox | None = None
        self.quick_employee_menu: ctk.CTkOptionMenu | None = None
        self.quick_date_entry: ctk.CTkEntry | None = None
        self.quick_date_end_entry: ctk.CTkEntry | None = None
        self.quick_date_button: ctk.CTkButton | None = None
        self.quick_type_menu: ctk.CTkOptionMenu | None = None
        self.quick_detail_entry: ctk.CTkEntry | None = None
        self.quick_add_button: ctk.CTkButton | None = None
        self.quick_remove_last_button: ctk.CTkButton | None = None
        self.quick_clear_manual_button: ctk.CTkButton | None = None
        self.exceptions_summary_label: ctk.CTkLabel | None = None
        self.context_file_value: ctk.CTkLabel | None = None
        self.context_exceptions_value: ctk.CTkLabel | None = None
        self.context_manual_value: ctk.CTkLabel | None = None
        self.context_output_value: ctk.CTkLabel | None = None
        self.kpi_employees_value: ctk.CTkLabel | None = None
        self.kpi_days_value: ctk.CTkLabel | None = None
        self.kpi_late_value: ctk.CTkLabel | None = None
        self.kpi_absent_value: ctk.CTkLabel | None = None
        self.kpi_overtime_value: ctk.CTkLabel | None = None
        self._logo_image: tk.PhotoImage | None = None
        self._logo_badge: ctk.CTkFrame | None = None
        self._logo_label: ctk.CTkLabel | None = None
        self._employee_selector_values: list[str] = ["Todos"]
        self._employee_label_to_id: dict[str, str] = {"Todos": ""}
        self._status_pulse_job: str | None = None
        self._status_pulse_on: bool = False

        self._build_ui()
        self._initialize_default_exceptions_file()
        self._refresh_employee_options()

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.navbar = ctk.CTkFrame(self, corner_radius=0, height=66, fg_color=("#DCE8F6", "#0E1A2B"))
        self.navbar.grid(row=0, column=0, sticky="ew")
        self.navbar.grid_columnconfigure(0, weight=0)
        self.navbar.grid_columnconfigure(1, weight=0)
        self.navbar.grid_columnconfigure(2, weight=0)
        self.navbar.grid_columnconfigure(3, weight=1)
        self.navbar.grid_columnconfigure(4, weight=0)
        self.navbar.grid_rowconfigure(0, weight=0)
        self.navbar.grid_rowconfigure(1, weight=0, minsize=3)

        self.nav_title = ctk.CTkLabel(
            self.navbar,
            text="Hikvision Hours",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COLOR_TEXT,
        )
        self.nav_title.grid(row=0, column=0, padx=(14, 18), pady=(10, 8), sticky="w")

        self.report_tab_button = ctk.CTkButton(
            self.navbar,
            text="Reporte",
            width=110,
            height=32,
            corner_radius=6,
            fg_color="transparent",
            hover_color=COLOR_SECONDARY_HOVER,
            text_color=COLOR_TEXT_MUTED,
            font=self.font_button,
            command=lambda: self._show_page("report"),
        )
        self.report_tab_button.grid(row=0, column=1, padx=(0, 6), pady=(12, 8), sticky="w")

        self.exceptions_tab_button = ctk.CTkButton(
            self.navbar,
            text="Excepciones",
            width=120,
            height=32,
            corner_radius=6,
            fg_color="transparent",
            hover_color=COLOR_SECONDARY_HOVER,
            text_color=COLOR_TEXT_MUTED,
            font=self.font_button,
            command=lambda: self._show_page("exceptions"),
        )
        self.exceptions_tab_button.grid(row=0, column=2, padx=(0, 10), pady=(12, 8), sticky="w")

        self.report_indicator = ctk.CTkFrame(
            self.navbar,
            height=3,
            width=110,
            corner_radius=0,
            fg_color=COLOR_PRIMARY,
        )
        self.report_indicator.grid(row=1, column=1, padx=(0, 6), sticky="sw")

        self.exceptions_indicator = ctk.CTkFrame(
            self.navbar,
            height=3,
            width=120,
            corner_radius=0,
            fg_color=COLOR_PRIMARY,
        )
        self.exceptions_indicator.grid(row=1, column=2, padx=(0, 10), sticky="sw")

        self._mount_logo()

        self.content = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.content.grid(row=1, column=0, sticky="nsew")
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)

        self.report_page = self._build_report_page(self.content)
        self.exceptions_page = self._build_exceptions_page(self.content)

        self.report_page.grid(row=0, column=0, sticky="nsew")
        self.exceptions_page.grid(row=0, column=0, sticky="nsew")

        self._reset_kpis()
        self._refresh_exceptions_summary()
        self._refresh_context_panel()
        self._show_page("report")

    def _style_primary_button(self, button: ctk.CTkButton, *, height: int | None = None) -> None:
        kwargs: dict[str, object] = {
            "fg_color": COLOR_PRIMARY,
            "hover_color": COLOR_PRIMARY_HOVER,
            "text_color": ("#FFFFFF", "#06202A"),
            "font": self.font_button,
            "corner_radius": 10,
        }
        if height is not None:
            kwargs["height"] = height
        button.configure(**kwargs)

    def _style_secondary_button(self, button: ctk.CTkButton, *, height: int | None = None) -> None:
        kwargs: dict[str, object] = {
            "fg_color": "transparent",
            "hover_color": COLOR_SECONDARY_HOVER,
            "border_width": 1,
            "border_color": COLOR_BORDER,
            "text_color": COLOR_TEXT,
            "font": self.font_body,
            "corner_radius": 10,
        }
        if height is not None:
            kwargs["height"] = height
        button.configure(**kwargs)

    def _set_context_value(self, label: ctk.CTkLabel | None, value: str) -> None:
        if label is not None:
            label.configure(text=value)

    def _minutes_to_hhmm(self, total_minutes: int) -> str:
        hours, minutes = divmod(max(0, int(total_minutes)), 60)
        return f"{hours:02d}:{minutes:02d}"

    def _set_kpi_value(self, label: ctk.CTkLabel | None, value: str) -> None:
        if label is not None:
            label.configure(text=value)

    def _marker_is_positive(self, value: object) -> bool:
        text = str(value).strip()
        if text in {"", "-", "0", "0.0", "0:0", "0:00", "nan", "NaN", "None"}:
            return False
        if ":" in text:
            parts = [p.strip() for p in text.split(":") if p.strip() != ""]
            if not parts:
                return False
            try:
                return any(int(p) > 0 for p in parts)
            except ValueError:
                return True
        try:
            return float(text.replace(",", ".")) > 0
        except ValueError:
            return True

    def _reset_kpis(self) -> None:
        self._set_kpi_value(self.kpi_employees_value, "-")
        self._set_kpi_value(self.kpi_days_value, "-")
        self._set_kpi_value(self.kpi_late_value, "-")
        self._set_kpi_value(self.kpi_absent_value, "-")
        self._set_kpi_value(self.kpi_overtime_value, "-")

    def _refresh_kpis_from_daily_df(self, daily_df: pd.DataFrame) -> None:
        if daily_df.empty:
            self._reset_kpis()
            return

        employees = int(daily_df["ID de persona"].astype(str).str.strip().replace("", pd.NA).dropna().nunique()) if "ID de persona" in daily_df.columns else 0
        days = int(len(daily_df))
        status_series = daily_df["Estado"].astype(str).str.strip().str.lower() if "Estado" in daily_df.columns else pd.Series(dtype=str)
        late_count = int(status_series.isin(["tarde", "tardanza"]).sum())
        absent_count = int(status_series.isin(["ausente", "ausencia"]).sum())
        overtime_minutes = (
            pd.to_numeric(daily_df["Minutos extra"], errors="coerce").fillna(0).astype(int).sum()
            if "Minutos extra" in daily_df.columns
            else 0
        )

        self._set_kpi_value(self.kpi_employees_value, str(employees))
        self._set_kpi_value(self.kpi_days_value, str(days))
        self._set_kpi_value(self.kpi_late_value, str(late_count))
        self._set_kpi_value(self.kpi_absent_value, str(absent_count))
        self._set_kpi_value(self.kpi_overtime_value, self._minutes_to_hhmm(int(overtime_minutes)))

    def _refresh_kpis_from_source(self, source_path: Path) -> None:
        try:
            df = load_hikvision_excel(source_path)
        except Exception:
            self._reset_kpis()
            return

        if df.empty:
            self._reset_kpis()
            return

        employees = int(df["employee_id"].astype(str).str.strip().replace("", pd.NA).dropna().nunique()) if "employee_id" in df.columns else 0
        days = int(pd.to_datetime(df["work_date_raw"], errors="coerce").dropna().dt.date.nunique()) if "work_date_raw" in df.columns else 0
        late_count = int(df["late_raw"].map(self._marker_is_positive).sum()) if "late_raw" in df.columns else 0
        absent_count = int(df["absent_raw"].map(self._marker_is_positive).sum()) if "absent_raw" in df.columns else 0

        self._set_kpi_value(self.kpi_employees_value, str(employees))
        self._set_kpi_value(self.kpi_days_value, str(days))
        self._set_kpi_value(self.kpi_late_value, str(late_count))
        self._set_kpi_value(self.kpi_absent_value, str(absent_count))
        self._set_kpi_value(self.kpi_overtime_value, "--:--")

    def _refresh_kpis_from_output(self, output_path: Path) -> None:
        try:
            daily_df = pd.read_excel(output_path, sheet_name="Diario")
        except Exception:
            return
        self._refresh_kpis_from_daily_df(daily_df)

    def _refresh_context_panel(self) -> None:
        file_value = self.selected_file.name if self.selected_file else "Sin archivo"
        exc_value = self.exceptions_file.name if self.exceptions_file else "Sin archivo"
        manual_text = self.exceptions_manual_box.get("1.0", "end").strip() if self.exceptions_manual_box else ""
        manual_value = f"{self._count_manual_exceptions(manual_text)} lineas"
        output_value = self.output_file.name if self.output_file else "No generado"

        self._set_context_value(self.context_file_value, file_value)
        self._set_context_value(self.context_exceptions_value, exc_value)
        self._set_context_value(self.context_manual_value, manual_value)
        self._set_context_value(self.context_output_value, output_value)

    def _start_status_pulse(self) -> None:
        if self._status_pulse_job is not None:
            return

        def _pulse() -> None:
            self._status_pulse_on = not self._status_pulse_on
            self.status_badge.configure(
                fg_color=COLOR_STATUS_WORKING_BG_ALT if self._status_pulse_on else COLOR_STATUS_WORKING_BG
            )
            self._status_pulse_job = self.after(500, _pulse)

        self._status_pulse_on = False
        _pulse()

    def _stop_status_pulse(self) -> None:
        if self._status_pulse_job is not None:
            self.after_cancel(self._status_pulse_job)
            self._status_pulse_job = None
        self._status_pulse_on = False

    def _mount_logo(self) -> None:
        logo_path = Path(__file__).resolve().parent.parent / "logo.png"
        if not logo_path.exists():
            return

        try:
            raw_logo = tk.PhotoImage(file=str(logo_path))
        except Exception:
            return

        max_width = 150
        max_height = 30
        ratio = max(raw_logo.width() / max_width, raw_logo.height() / max_height, 1)
        scale = max(1, math.ceil(ratio))

        self._logo_image = raw_logo.subsample(scale, scale) if scale > 1 else raw_logo
        self._logo_badge = ctk.CTkFrame(
            self.navbar,
            corner_radius=12,
            border_width=1,
            border_color=COLOR_BORDER,
            fg_color=("#F8FBFF", "#111827"),
            width=190,
            height=40,
        )
        self._logo_badge.grid(row=0, column=4, padx=(8, 14), pady=(10, 8), sticky="e")
        self._logo_badge.grid_propagate(False)

        self._logo_label = ctk.CTkLabel(
            self._logo_badge,
            text="",
            image=self._logo_image,
            fg_color="transparent",
        )
        self._logo_label.place(relx=0.5, rely=0.5, anchor="center")

    def _build_report_page(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        page = ctk.CTkFrame(parent, corner_radius=0, fg_color="transparent")
        page.grid_columnconfigure(0, weight=0)
        page.grid_columnconfigure(1, weight=1)
        page.grid_rowconfigure(0, weight=1)

        sidebar = ctk.CTkFrame(
            page,
            width=370,
            corner_radius=18,
            border_width=1,
            border_color=COLOR_BORDER,
            fg_color=COLOR_PANEL,
        )
        sidebar.grid(row=0, column=0, sticky="nsw", padx=(14, 10), pady=12)
        sidebar.grid_propagate(False)
        sidebar.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            sidebar,
            text="Procesador Hikvision",
            font=self.font_title,
            text_color=COLOR_TEXT,
        )
        title.grid(row=0, column=0, padx=20, pady=(22, 8), sticky="w")

        subtitle = ctk.CTkLabel(
            sidebar,
            text="Importa fichadas y genera reportes\ndiarios/mensuales.",
            justify="left",
            text_color=COLOR_TEXT_MUTED,
            font=self.font_subtitle,
        )
        subtitle.grid(row=1, column=0, padx=20, pady=(0, 16), sticky="w")

        self.status_badge = ctk.CTkLabel(
            sidebar,
            text="Estado: listo",
            corner_radius=8,
            fg_color=COLOR_STATUS_READY_BG,
            text_color=COLOR_STATUS_READY_TEXT,
            padx=10,
            pady=6,
            font=self.font_body,
        )
        self.status_badge.grid(row=2, column=0, padx=20, pady=(0, 14), sticky="w")

        select_label = ctk.CTkLabel(
            sidebar, text="Archivo de entrada", font=self.font_button, text_color=COLOR_TEXT
        )
        select_label.grid(row=3, column=0, padx=20, pady=(0, 8), sticky="w")

        self.file_entry = ctk.CTkEntry(sidebar, placeholder_text="Ningun archivo seleccionado")
        self.file_entry.grid(row=4, column=0, padx=20, pady=(0, 8), sticky="ew")
        self.file_entry.configure(border_color=COLOR_BORDER, fg_color=COLOR_PANEL_SOFT, text_color=COLOR_TEXT)
        self.file_entry.configure(state="disabled")

        self.file_meta = ctk.CTkLabel(
            sidebar,
            text="",
            justify="left",
            text_color=COLOR_TEXT_MUTED,
            font=self.font_body,
        )
        self.file_meta.grid(row=5, column=0, padx=20, pady=(0, 12), sticky="w")

        self.select_button = ctk.CTkButton(
            sidebar,
            text="Seleccionar archivo",
            height=38,
            command=self.select_file,
        )
        self._style_primary_button(self.select_button, height=38)
        self.select_button.grid(row=6, column=0, padx=20, pady=(0, 10), sticky="ew")

        self.exceptions_summary_label = ctk.CTkLabel(
            sidebar,
            text="Excepciones: sin configurar",
            justify="left",
            text_color=COLOR_TEXT_MUTED,
            font=self.font_body,
        )
        self.exceptions_summary_label.grid(row=7, column=0, padx=20, pady=(0, 8), sticky="w")

        goto_exceptions = ctk.CTkButton(
            sidebar,
            text="Ir a Excepciones",
            height=34,
            command=lambda: self._show_page("exceptions"),
        )
        self._style_secondary_button(goto_exceptions, height=34)
        goto_exceptions.grid(row=8, column=0, padx=20, pady=(0, 10), sticky="ew")

        self.process_button = ctk.CTkButton(
            sidebar,
            text="Procesar reporte",
            height=40,
            command=self.process_file,
            state="disabled",
        )
        self._style_primary_button(self.process_button, height=40)
        self.process_button.grid(row=9, column=0, padx=20, pady=(2, 10), sticky="ew")

        self.open_button = ctk.CTkButton(
            sidebar,
            text="Abrir reporte generado",
            height=36,
            command=self.open_output_file,
            state="disabled",
        )
        self._style_primary_button(self.open_button, height=36)
        self.open_button.grid(row=10, column=0, padx=20, pady=(0, 8), sticky="ew")

        self.clear_button = ctk.CTkButton(
            sidebar,
            text="Limpiar bitacora",
            height=36,
            command=self.clear_log,
        )
        self._style_secondary_button(self.clear_button, height=36)
        self.clear_button.grid(row=11, column=0, padx=20, pady=(0, 14), sticky="ew")

        self.progress = ctk.CTkProgressBar(sidebar, mode="determinate")
        self.progress.configure(progress_color=COLOR_PRIMARY, fg_color=("#DCEAF6", "#1F2937"))
        self.progress.grid(row=12, column=0, padx=20, pady=(0, 8), sticky="ew")
        self.progress.set(0)

        self.progress_text = ctk.CTkLabel(sidebar, text="Progreso: 0%", text_color=COLOR_TEXT_MUTED, font=self.font_body)
        self.progress_text.grid(row=13, column=0, padx=20, pady=(0, 4), sticky="w")

        self.progress_step = ctk.CTkLabel(
            sidebar,
            text="Paso: esperando archivo",
            text_color=COLOR_TEXT_MUTED,
            font=self.font_body,
        )
        self.progress_step.grid(row=14, column=0, padx=20, pady=(0, 4), sticky="w")

        self.elapsed_text = ctk.CTkLabel(sidebar, text="Tiempo: 00:00", text_color=COLOR_TEXT_MUTED, font=self.font_body)
        self.elapsed_text.grid(row=15, column=0, padx=20, pady=(0, 10), sticky="w")

        info = ctk.CTkLabel(
            sidebar,
            text="Entrada: .xls/.xlsx",
            text_color=COLOR_TEXT_MUTED,
            justify="left",
            font=self.font_body,
        )
        info.grid(row=16, column=0, padx=20, pady=(0, 12), sticky="w")

        main = ctk.CTkFrame(
            page,
            corner_radius=18,
            border_width=1,
            border_color=COLOR_BORDER,
            fg_color=COLOR_PANEL,
        )
        main.grid(row=0, column=1, sticky="nsew", padx=(8, 14), pady=12)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(4, weight=1)

        header = ctk.CTkLabel(
            main,
            text="Bitacora de ejecucion",
            font=self.font_section,
            text_color=COLOR_TEXT,
        )
        header.grid(row=0, column=0, sticky="w", pady=(2, 6))

        helper = ctk.CTkLabel(
            main,
            text="Aca vas a ver cada paso del procesamiento y errores detectados.",
            text_color=COLOR_TEXT_MUTED,
            font=self.font_subtitle,
        )
        helper.grid(row=1, column=0, sticky="w", pady=(0, 8))

        kpi_grid = ctk.CTkFrame(main, fg_color="transparent")
        kpi_grid.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        for idx in range(5):
            kpi_grid.grid_columnconfigure(idx, weight=1)

        def _add_kpi(col: int, title: str) -> ctk.CTkLabel:
            card = ctk.CTkFrame(
                kpi_grid,
                corner_radius=12,
                border_width=1,
                border_color=COLOR_BORDER,
                fg_color=COLOR_PANEL_SOFT,
            )
            card.grid(row=0, column=col, padx=(0 if col == 0 else 6, 0), sticky="ew")
            ctk.CTkLabel(card, text=title, text_color=COLOR_TEXT_MUTED, font=self.font_body).pack(
                anchor="w", padx=10, pady=(8, 0)
            )
            value = ctk.CTkLabel(card, text="-", text_color=COLOR_TEXT, font=ctk.CTkFont(family="Bahnschrift", size=22, weight="bold"))
            value.pack(anchor="w", padx=10, pady=(2, 10))
            return value

        self.kpi_employees_value = _add_kpi(0, "Empleados")
        self.kpi_days_value = _add_kpi(1, "Jornadas")
        self.kpi_late_value = _add_kpi(2, "Tardanzas")
        self.kpi_absent_value = _add_kpi(3, "Ausencias")
        self.kpi_overtime_value = _add_kpi(4, "Horas Extra")

        context_card = ctk.CTkFrame(
            main,
            corner_radius=12,
            border_width=1,
            border_color=COLOR_BORDER,
            fg_color=COLOR_PANEL_SOFT,
        )
        context_card.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        context_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(context_card, text="Archivo:", text_color=COLOR_TEXT_MUTED, font=self.font_body).grid(
            row=0, column=0, padx=(12, 8), pady=(10, 4), sticky="w"
        )
        self.context_file_value = ctk.CTkLabel(context_card, text="Sin archivo", text_color=COLOR_TEXT, font=self.font_body)
        self.context_file_value.grid(row=0, column=1, padx=(0, 12), pady=(10, 4), sticky="w")

        ctk.CTkLabel(context_card, text="Excepciones:", text_color=COLOR_TEXT_MUTED, font=self.font_body).grid(
            row=1, column=0, padx=(12, 8), pady=4, sticky="w"
        )
        self.context_exceptions_value = ctk.CTkLabel(context_card, text="Sin archivo", text_color=COLOR_TEXT, font=self.font_body)
        self.context_exceptions_value.grid(row=1, column=1, padx=(0, 12), pady=4, sticky="w")

        ctk.CTkLabel(context_card, text="Manual pendiente:", text_color=COLOR_TEXT_MUTED, font=self.font_body).grid(
            row=2, column=0, padx=(12, 8), pady=4, sticky="w"
        )
        self.context_manual_value = ctk.CTkLabel(context_card, text="0 lineas", text_color=COLOR_TEXT, font=self.font_body)
        self.context_manual_value.grid(row=2, column=1, padx=(0, 12), pady=4, sticky="w")

        ctk.CTkLabel(context_card, text="Ultimo reporte:", text_color=COLOR_TEXT_MUTED, font=self.font_body).grid(
            row=3, column=0, padx=(12, 8), pady=(4, 10), sticky="w"
        )
        self.context_output_value = ctk.CTkLabel(context_card, text="No generado", text_color=COLOR_TEXT, font=self.font_body)
        self.context_output_value.grid(row=3, column=1, padx=(0, 12), pady=(4, 10), sticky="w")

        self.log_box = ctk.CTkTextbox(main, corner_radius=10, border_width=1)
        self.log_box.configure(border_color=COLOR_BORDER, fg_color=COLOR_PANEL_SOFT, text_color=COLOR_TEXT, font=self.font_mono)
        self.log_box.grid(row=4, column=0, sticky="nsew")
        self.log_box.insert("end", "Listo para procesar fichadas.\n")
        self.log_box.configure(state="disabled")

        watermark = ctk.CTkLabel(
            page,
            text="PORTA",
            font=ctk.CTkFont(size=116, weight="bold"),
            text_color=("#DCE6F2", "#172436"),
        )
        watermark.place(relx=0.64, rely=0.52, anchor="center")
        watermark.lower()

        return page

    def _build_exceptions_page(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        page = ctk.CTkFrame(parent, corner_radius=0, fg_color="transparent")
        page.grid_columnconfigure(0, weight=0)
        page.grid_columnconfigure(1, weight=1)
        page.grid_rowconfigure(0, weight=1)

        sidebar = ctk.CTkFrame(
            page,
            width=370,
            corner_radius=18,
            border_width=1,
            border_color=COLOR_BORDER,
            fg_color=COLOR_PANEL,
        )
        sidebar.grid(row=0, column=0, sticky="nsw", padx=(14, 10), pady=12)
        sidebar.grid_propagate(False)
        sidebar.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            sidebar,
            text="Excepciones",
            font=self.font_title,
            text_color=COLOR_TEXT,
        )
        title.grid(row=0, column=0, padx=20, pady=(22, 8), sticky="w")

        subtitle = ctk.CTkLabel(
            sidebar,
            text="Configura feriados, vacaciones,\nenfermedad y permisos.",
            justify="left",
            text_color=COLOR_TEXT_MUTED,
            font=self.font_subtitle,
        )
        subtitle.grid(row=1, column=0, padx=20, pady=(0, 16), sticky="w")

        file_label = ctk.CTkLabel(
            sidebar,
            text="Archivo de excepciones",
            font=self.font_button,
            text_color=COLOR_TEXT,
        )
        file_label.grid(row=2, column=0, padx=20, pady=(0, 8), sticky="w")

        self.exceptions_file_entry = ctk.CTkEntry(sidebar, placeholder_text="Sin archivo de excepciones")
        self.exceptions_file_entry.grid(row=3, column=0, padx=20, pady=(0, 8), sticky="ew")
        self.exceptions_file_entry.configure(border_color=COLOR_BORDER, fg_color=COLOR_PANEL_SOFT, text_color=COLOR_TEXT)
        self.exceptions_file_entry.configure(state="disabled")

        self.select_exceptions_button = ctk.CTkButton(
            sidebar,
            text="Seleccionar archivo",
            height=34,
            command=self.select_exceptions_file,
        )
        self._style_primary_button(self.select_exceptions_button, height=34)
        self.select_exceptions_button.grid(row=4, column=0, padx=20, pady=(0, 8), sticky="ew")

        self.clear_exceptions_button = ctk.CTkButton(
            sidebar,
            text="Usar archivo por defecto",
            height=34,
            command=self.clear_exceptions_file,
        )
        self._style_secondary_button(self.clear_exceptions_button, height=34)
        self.clear_exceptions_button.grid(row=5, column=0, padx=20, pady=(0, 8), sticky="ew")

        self.save_exceptions_button = ctk.CTkButton(
            sidebar,
            text="Guardar configuracion",
            height=36,
            command=self.save_exceptions_config,
        )
        self._style_primary_button(self.save_exceptions_button, height=36)
        self.save_exceptions_button.grid(row=6, column=0, padx=20, pady=(6, 8), sticky="ew")

        go_report = ctk.CTkButton(
            sidebar,
            text="Volver a Reporte",
            height=34,
            command=lambda: self._show_page("report"),
        )
        self._style_secondary_button(go_report, height=34)
        go_report.grid(row=7, column=0, padx=20, pady=(0, 8), sticky="ew")

        info = ctk.CTkLabel(
            sidebar,
            text="Formatos: .csv/.xls/.xlsx",
            text_color=COLOR_TEXT_MUTED,
            justify="left",
            font=self.font_body,
        )
        info.grid(row=8, column=0, padx=20, pady=(0, 12), sticky="w")

        main = ctk.CTkFrame(
            page,
            corner_radius=18,
            border_width=1,
            border_color=COLOR_BORDER,
            fg_color=COLOR_PANEL,
        )
        main.grid(row=0, column=1, sticky="nsew", padx=(8, 14), pady=12)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(4, weight=1)

        header = ctk.CTkLabel(
            main,
            text="Carga Manual",
            font=self.font_section,
            text_color=COLOR_TEXT,
        )
        header.grid(row=0, column=0, sticky="w", pady=(2, 6))

        helper = ctk.CTkLabel(
            main,
            text=(
                "Carga rapida por campos o pega varias lineas.\n"
                "Soporta rango Fecha inicio/fin (ejemplo: vacaciones).\n"
                "Formato manual: ID|YYYY-MM-DD|TIPO|DETALLE"
            ),
            justify="left",
            text_color=COLOR_TEXT_MUTED,
            font=self.font_subtitle,
        )
        helper.grid(row=1, column=0, sticky="w", pady=(0, 8))

        quick_card = ctk.CTkFrame(
            main,
            corner_radius=12,
            border_width=1,
            border_color=COLOR_BORDER,
            fg_color=COLOR_PANEL_SOFT,
        )
        quick_card.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        quick_card.grid_columnconfigure(0, weight=1)
        quick_card.grid_columnconfigure(1, weight=1)
        quick_card.grid_columnconfigure(2, weight=1)
        quick_card.grid_columnconfigure(3, weight=2)

        quick_title = ctk.CTkLabel(
            quick_card,
            text="Carga Rapida",
            font=self.font_button,
            text_color=COLOR_TEXT,
        )
        quick_title.grid(row=0, column=0, columnspan=4, sticky="w", padx=12, pady=(10, 8))

        self.quick_employee_menu = ctk.CTkOptionMenu(
            quick_card,
            values=self._employee_selector_values,
        )
        self.quick_employee_menu.configure(
            fg_color=COLOR_PANEL,
            button_color=COLOR_PRIMARY,
            button_hover_color=COLOR_PRIMARY_HOVER,
            text_color=COLOR_TEXT,
            font=self.font_body,
        )
        self.quick_employee_menu.grid(row=1, column=0, padx=(12, 6), pady=(0, 8), sticky="ew")
        self.quick_employee_menu.set("Todos")

        date_row = ctk.CTkFrame(quick_card, fg_color="transparent")
        date_row.grid(row=1, column=1, padx=6, pady=(0, 8), sticky="ew")
        date_row.grid_columnconfigure(0, weight=1)
        date_row.grid_columnconfigure(1, weight=1)
        date_row.grid_columnconfigure(2, weight=0)

        self.quick_date_entry = ctk.CTkEntry(
            date_row,
            placeholder_text="Inicio (YYYY-MM-DD o DD/MM/YYYY)",
        )
        self.quick_date_entry.configure(border_color=COLOR_BORDER, fg_color=COLOR_PANEL, text_color=COLOR_TEXT)
        self.quick_date_entry.grid(row=0, column=0, sticky="ew")

        self.quick_date_end_entry = ctk.CTkEntry(
            date_row,
            placeholder_text="Fin (opcional)",
        )
        self.quick_date_end_entry.configure(border_color=COLOR_BORDER, fg_color=COLOR_PANEL, text_color=COLOR_TEXT)
        self.quick_date_end_entry.grid(row=0, column=1, padx=(6, 0), sticky="ew")

        self.quick_date_button = ctk.CTkButton(
            date_row,
            text="Calendario",
            width=96,
            height=28,
            command=self.open_date_picker,
        )
        self._style_secondary_button(self.quick_date_button, height=28)
        self.quick_date_button.grid(row=0, column=2, padx=(6, 0), sticky="e")

        self.quick_type_menu = ctk.CTkOptionMenu(
            quick_card,
            values=["Feriado", "Vacaciones", "Enfermedad", "Permiso", "Ausencia justificada", "Otro"],
        )
        self.quick_type_menu.configure(
            fg_color=COLOR_PANEL,
            button_color=COLOR_PRIMARY,
            button_hover_color=COLOR_PRIMARY_HOVER,
            text_color=COLOR_TEXT,
            font=self.font_body,
        )
        self.quick_type_menu.grid(row=1, column=2, padx=6, pady=(0, 8), sticky="ew")
        self.quick_type_menu.set("Feriado")

        self.quick_detail_entry = ctk.CTkEntry(quick_card, placeholder_text="Detalle (opcional)")
        self.quick_detail_entry.configure(border_color=COLOR_BORDER, fg_color=COLOR_PANEL, text_color=COLOR_TEXT)
        self.quick_detail_entry.grid(row=1, column=3, padx=(6, 12), pady=(0, 8), sticky="ew")

        self.quick_add_button = ctk.CTkButton(
            quick_card,
            text="Agregar a carga manual",
            height=32,
            command=self.add_manual_exception_from_fields,
        )
        self._style_primary_button(self.quick_add_button, height=32)
        self.quick_add_button.grid(row=2, column=0, columnspan=2, padx=(12, 6), pady=(0, 10), sticky="ew")

        self.quick_remove_last_button = ctk.CTkButton(
            quick_card,
            text="Quitar ultima linea",
            height=32,
            command=self.remove_last_manual_exception_line,
        )
        self._style_secondary_button(self.quick_remove_last_button, height=32)
        self.quick_remove_last_button.grid(row=2, column=2, padx=6, pady=(0, 10), sticky="ew")

        self.quick_clear_manual_button = ctk.CTkButton(
            quick_card,
            text="Limpiar manual",
            height=32,
            command=self.clear_manual_exceptions_box,
        )
        self._style_secondary_button(self.quick_clear_manual_button, height=32)
        self.quick_clear_manual_button.grid(row=2, column=3, padx=(6, 12), pady=(0, 10), sticky="ew")

        legend_card = ctk.CTkFrame(
            main,
            corner_radius=12,
            border_width=1,
            border_color=COLOR_BORDER,
            fg_color=COLOR_PANEL_SOFT,
        )
        legend_card.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        legend_card.grid_columnconfigure((0, 1, 2), weight=1)

        ctk.CTkLabel(
            legend_card,
            text="Leyenda Rapida de Estados",
            font=self.font_button,
            text_color=COLOR_TEXT,
        ).grid(row=0, column=0, columnspan=3, padx=12, pady=(8, 6), sticky="w")

        legend_items = [
            ("Feriado", "#A9DF8F"),
            ("Vacaciones", "#FFF59D"),
            ("Ausente", "#EF9A9A"),
            ("Tardanza", "#2E7D32"),
            ("Domingo", "#D9D9D9"),
            ("Licencia", "#FFCC80"),
        ]
        for idx, (name, color) in enumerate(legend_items):
            row = 1 + (idx // 3)
            col = idx % 3
            chip = ctk.CTkFrame(legend_card, corner_radius=8, fg_color=color)
            chip.grid(row=row, column=col, padx=8, pady=6, sticky="ew")
            ctk.CTkLabel(
                chip,
                text=name,
                text_color="#102A43",
                font=self.font_body,
            ).pack(padx=8, pady=4)

        self.exceptions_manual_box = ctk.CTkTextbox(main, corner_radius=10, border_width=1)
        self.exceptions_manual_box.configure(
            border_color=COLOR_BORDER,
            fg_color=COLOR_PANEL_SOFT,
            text_color=COLOR_TEXT,
            font=self.font_mono,
        )
        self.exceptions_manual_box.grid(row=4, column=0, sticky="nsew")
        self.exceptions_manual_box.insert(
            "end",
            "# Formato por linea: ID|YYYY-MM-DD|TIPO|DETALLE\n"
            "# Ejemplo: 20|2026-05-01|Feriado|Dia del trabajador\n",
        )

        watermark = ctk.CTkLabel(
            page,
            text="PORTA",
            font=ctk.CTkFont(size=116, weight="bold"),
            text_color=("#DCE6F2", "#172436"),
        )
        watermark.place(relx=0.64, rely=0.52, anchor="center")
        watermark.lower()

        return page

    def _show_page(self, page: str) -> None:
        self.current_page = page
        active_tab_fg = COLOR_TAB_ACTIVE
        inactive_tab_fg = "transparent"
        active_text = COLOR_TAB_TEXT_ACTIVE
        inactive_text = COLOR_TEXT_MUTED
        active_line = COLOR_PRIMARY
        hidden_line = ("#DCE8F6", "#0E1A2B")

        if page == "report":
            self.report_page.tkraise()
            self.report_tab_button.configure(fg_color=active_tab_fg, text_color=active_text)
            self.exceptions_tab_button.configure(fg_color=inactive_tab_fg, text_color=inactive_text)
            self.report_indicator.configure(fg_color=active_line)
            self.exceptions_indicator.configure(fg_color=hidden_line)
        else:
            self.exceptions_page.tkraise()
            self.exceptions_tab_button.configure(fg_color=active_tab_fg, text_color=active_text)
            self.report_tab_button.configure(fg_color=inactive_tab_fg, text_color=inactive_text)
            self.exceptions_indicator.configure(fg_color=active_line)
            self.report_indicator.configure(fg_color=hidden_line)

    def _set_status(self, text: str, level: str = "ready") -> None:
        palette = {
            "ready": {"fg": COLOR_STATUS_READY_BG, "text": COLOR_STATUS_READY_TEXT},
            "working": {"fg": COLOR_STATUS_WORKING_BG, "text": COLOR_STATUS_WORKING_TEXT},
            "error": {"fg": COLOR_STATUS_ERROR_BG, "text": COLOR_STATUS_ERROR_TEXT},
        }
        colors = palette.get(level, palette["ready"])
        self.status_badge.configure(
            text=f"Estado: {text}",
            fg_color=colors["fg"],
            text_color=colors["text"],
        )
        if level == "working":
            self._start_status_pulse()
        else:
            self._stop_status_pulse()

    def _set_file_entry(self, text: str) -> None:
        self.file_entry.configure(state="normal")
        self.file_entry.delete(0, "end")
        self.file_entry.insert(0, text)
        self.file_entry.configure(state="disabled")

    def _set_exceptions_file_entry(self, text: str) -> None:
        if self.exceptions_file_entry is None:
            return
        self.exceptions_file_entry.configure(state="normal")
        self.exceptions_file_entry.delete(0, "end")
        if text:
            self.exceptions_file_entry.insert(0, text)
        self.exceptions_file_entry.configure(state="disabled")

    def _set_processing_state(self, enabled: bool) -> None:
        self.select_button.configure(state="normal" if enabled else "disabled")
        self.process_button.configure(state="normal" if enabled and self.selected_file else "disabled")
        self.clear_button.configure(state="normal" if enabled else "disabled")

        self.select_exceptions_button.configure(state="normal" if enabled else "disabled")
        self.clear_exceptions_button.configure(state="normal" if enabled else "disabled")
        self.save_exceptions_button.configure(state="normal" if enabled else "disabled")
        if self.exceptions_manual_box is not None:
            self.exceptions_manual_box.configure(state="normal" if enabled else "disabled")
        if self.quick_employee_menu is not None:
            self.quick_employee_menu.configure(state="normal" if enabled else "disabled")
        if self.quick_date_entry is not None:
            self.quick_date_entry.configure(state="normal" if enabled else "disabled")
        if self.quick_date_end_entry is not None:
            self.quick_date_end_entry.configure(state="normal" if enabled else "disabled")
        if self.quick_date_button is not None:
            self.quick_date_button.configure(state="normal" if enabled else "disabled")
        if self.quick_type_menu is not None:
            self.quick_type_menu.configure(state="normal" if enabled else "disabled")
        if self.quick_detail_entry is not None:
            self.quick_detail_entry.configure(state="normal" if enabled else "disabled")
        if self.quick_add_button is not None:
            self.quick_add_button.configure(state="normal" if enabled else "disabled")
        if self.quick_remove_last_button is not None:
            self.quick_remove_last_button.configure(state="normal" if enabled else "disabled")
        if self.quick_clear_manual_button is not None:
            self.quick_clear_manual_button.configure(state="normal" if enabled else "disabled")

        self.report_tab_button.configure(state="normal" if enabled else "disabled")
        self.exceptions_tab_button.configure(state="normal" if enabled else "disabled")

    def _set_progress(self, percent: int, step_text: str) -> None:
        bounded = max(0, min(100, int(percent)))
        self.progress.set(bounded / 100)
        self.progress_text.configure(text=f"Progreso: {bounded}%")
        self.progress_step.configure(text=f"Paso: {step_text}")

    def _update_elapsed(self) -> None:
        if self._started_at is None:
            self.elapsed_text.configure(text="Tiempo: 00:00")
            return
        elapsed = int(time.perf_counter() - self._started_at)
        minutes, seconds = divmod(elapsed, 60)
        self.elapsed_text.configure(text=f"Tiempo: {minutes:02d}:{seconds:02d}")

    def _count_manual_exceptions(self, text: str) -> int:
        lines = [line.strip() for line in text.splitlines()]
        return sum(1 for line in lines if line and not line.startswith("#"))

    def _refresh_exceptions_summary(self) -> None:
        file_text = self.exceptions_file.name if self.exceptions_file else "sin archivo"
        manual_text = self.exceptions_manual_box.get("1.0", "end").strip() if self.exceptions_manual_box else ""
        manual_count = self._count_manual_exceptions(manual_text)
        if self.exceptions_summary_label is not None:
            self.exceptions_summary_label.configure(
                text=f"Excepciones: {file_text}\nCarga manual: {manual_count} lineas"
            )
        self._refresh_context_panel()

    def _default_exceptions_path(self) -> Path:
        return Path(__file__).resolve().parent.parent / DEFAULT_EXCEPTIONS_FILENAME

    def _initialize_default_exceptions_file(self) -> None:
        default_path = self._default_exceptions_path()
        try:
            ensure_exceptions_file(default_path)
        except ExceptionConfigError as exc:
            self.log(f"No se pudo inicializar archivo de excepciones por defecto: {exc}")
            return

        self.exceptions_file = default_path
        self._set_exceptions_file_entry(str(self.exceptions_file))
        self._refresh_exceptions_summary()
        self.log(f"Archivo de excepciones por defecto: {self.exceptions_file}")

    def _manual_text_from_box(self) -> str:
        return self.exceptions_manual_box.get("1.0", "end").strip() if self.exceptions_manual_box else ""

    def _reset_manual_box_to_template(self) -> None:
        if self.exceptions_manual_box is None:
            return
        self.exceptions_manual_box.delete("1.0", "end")
        self.exceptions_manual_box.insert(
            "end",
            "# Formato por linea: ID|YYYY-MM-DD|TIPO|DETALLE\n"
            "# Ejemplo: 20|2026-05-01|Feriado|Dia del trabajador\n",
        )
        self._refresh_exceptions_summary()

    def _persist_manual_exceptions(self) -> int:
        if self.exceptions_file is None:
            raise ExceptionConfigError("No hay archivo de excepciones seleccionado.")

        manual_text = self._manual_text_from_box()
        if self._count_manual_exceptions(manual_text) == 0:
            return 0

        added_count = append_manual_exceptions_to_file(self.exceptions_file, manual_text)
        self._reset_manual_box_to_template()
        return added_count

    def _set_employee_selector_values(self, labels: list[str]) -> None:
        if not labels:
            labels = ["Todos"]

        self._employee_selector_values = labels
        if self.quick_employee_menu is not None:
            self.quick_employee_menu.configure(values=labels)
            if "Todos" in labels:
                self.quick_employee_menu.set("Todos")
            else:
                self.quick_employee_menu.set(labels[0])

    def _normalize_label(self, text: object) -> str:
        ascii_text = unicodedata.normalize("NFKD", str(text)).encode("ascii", "ignore").decode("ascii")
        return " ".join(ascii_text.strip().lower().replace("_", " ").split())

    def _clean_employee_id(self, value: object) -> str:
        if pd.isna(value):
            return ""
        text = str(value).strip()
        if text.lower() in {"", "nan", "none", "nat", "-"}:
            return ""
        if text.endswith(".0") and text[:-2].isdigit():
            return text[:-2]
        return text

    def _extract_employee_pairs(
        self,
        df: pd.DataFrame,
        id_aliases: list[str],
        name_aliases: list[str],
    ) -> list[tuple[str, str]]:
        if df.empty:
            return []

        columns = [str(col) for col in df.columns]
        normalized = {self._normalize_label(col): col for col in columns}

        def _find_col(aliases: list[str]) -> str | None:
            for alias in aliases:
                key = self._normalize_label(alias)
                if key in normalized:
                    return normalized[key]
            for col in columns:
                col_key = self._normalize_label(col)
                if any(self._normalize_label(alias) in col_key for alias in aliases):
                    return col
            return None

        id_col = _find_col(id_aliases)
        name_col = _find_col(name_aliases)
        if id_col is None:
            return []

        by_id: dict[str, str] = {}
        for _, row in df.iterrows():
            emp_id = self._clean_employee_id(row.get(id_col, ""))
            if not emp_id:
                continue
            name_value = ""
            if name_col is not None:
                raw_name = str(row.get(name_col, "")).strip()
                if raw_name.lower() not in {"", "nan", "none", "nat", "-"}:
                    name_value = raw_name
            if emp_id not in by_id or (not by_id[emp_id] and name_value):
                by_id[emp_id] = name_value

        pairs = [(emp_id, by_id[emp_id] or f"ID {emp_id}") for emp_id in by_id]
        pairs.sort(key=lambda item: (item[1].lower(), item[0]))
        return pairs

    def _apply_employee_pairs(self, pairs: list[tuple[str, str]]) -> None:
        self._employee_label_to_id = {"Todos": ""}
        labels = ["Todos"]
        used_labels: set[str] = set()

        for emp_id, emp_name in pairs:
            label = emp_name
            if label in used_labels:
                label = f"{emp_name} ({emp_id})"
            used_labels.add(label)
            self._employee_label_to_id[label] = emp_id
            labels.append(label)

        self._set_employee_selector_values(labels)

    def _refresh_employee_options(self, source_path: Path | None = None) -> None:
        by_id: dict[str, str] = {}

        info_path = Path(__file__).resolve().parent.parent / "info.xlsx"
        if info_path.exists():
            try:
                info_df = pd.read_excel(info_path)
                info_pairs = self._extract_employee_pairs(
                    info_df,
                    id_aliases=["id", "legajo", "id de persona"],
                    name_aliases=["nombre", "name", "empleado"],
                )
                for emp_id, emp_name in info_pairs:
                    by_id[emp_id] = emp_name
            except Exception as exc:
                self.log(f"No se pudo cargar lista de trabajadores desde info.xlsx: {exc}")

        try:
            if source_path is not None:
                source_df = load_hikvision_excel(source_path)
                source_pairs = self._extract_employee_pairs(
                    source_df,
                    id_aliases=["employee_id", "id", "id de persona"],
                    name_aliases=["employee_name", "nombre", "name"],
                )
                for emp_id, emp_name in source_pairs:
                    if emp_id not in by_id:
                        by_id[emp_id] = emp_name
        except Exception as exc:
            self.log(f"No se pudo cargar lista de trabajadores para excepciones: {exc}")

        pairs = [(emp_id, by_id[emp_id]) for emp_id in by_id]
        pairs.sort(key=lambda item: (item[1].lower(), item[0]))
        self._apply_employee_pairs(pairs)

    def _load_employee_options_from_source(self, source_path: Path) -> None:
        self._refresh_employee_options(source_path)

    def open_date_picker(self) -> None:
        if self.quick_date_entry is None or self.quick_date_end_entry is None:
            return

        if Calendar is None:
            messagebox.showwarning(
                "Calendario no disponible",
                "Instala dependencia 'tkcalendar' para usar el selector visual de fecha.",
            )
            return

        seed_start = self._normalize_input_date(self.quick_date_entry.get().strip())
        seed_end = self._normalize_input_date(self.quick_date_end_entry.get().strip())
        seed_start_dt = datetime.strptime(seed_start, "%Y-%m-%d") if seed_start else datetime.now()
        seed_end_dt = datetime.strptime(seed_end, "%Y-%m-%d") if seed_end else seed_start_dt

        popup = ctk.CTkToplevel(self)
        popup.title("Seleccionar rango de fechas")
        popup.geometry("700x390")
        popup.resizable(False, False)
        popup.transient(self)
        popup.grab_set()

        body = ctk.CTkFrame(popup, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=12, pady=12)

        title_row = ctk.CTkFrame(body, fg_color="transparent")
        title_row.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(title_row, text="Inicio", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(46, 0))
        ctk.CTkLabel(title_row, text="Fin", font=ctk.CTkFont(size=12, weight="bold")).pack(side="right", padx=(0, 62))

        calendars_row = ctk.CTkFrame(body, fg_color="transparent")
        calendars_row.pack(fill="both", expand=True)
        calendars_row.grid_columnconfigure(0, weight=1)
        calendars_row.grid_columnconfigure(1, weight=1)

        calendar_start = Calendar(
            calendars_row,
            selectmode="day",
            year=seed_start_dt.year,
            month=seed_start_dt.month,
            day=seed_start_dt.day,
            date_pattern="yyyy-mm-dd",
        )
        calendar_start.grid(row=0, column=0, padx=(0, 8), sticky="nsew")

        calendar_end = Calendar(
            calendars_row,
            selectmode="day",
            year=seed_end_dt.year,
            month=seed_end_dt.month,
            day=seed_end_dt.day,
            date_pattern="yyyy-mm-dd",
        )
        calendar_end.grid(row=0, column=1, padx=(8, 0), sticky="nsew")

        actions = ctk.CTkFrame(body, fg_color="transparent")
        actions.pack(fill="x", pady=(10, 0))

        def _confirm() -> None:
            picked_start = calendar_start.get_date()
            picked_end = calendar_end.get_date()
            start_dt = datetime.strptime(picked_start, "%Y-%m-%d")
            end_dt = datetime.strptime(picked_end, "%Y-%m-%d")
            if end_dt < start_dt:
                messagebox.showwarning(
                    "Rango invalido",
                    "La fecha fin no puede ser menor a la fecha inicio.",
                    parent=popup,
                )
                return
            self.quick_date_entry.delete(0, "end")
            self.quick_date_entry.insert(0, picked_start)
            self.quick_date_end_entry.delete(0, "end")
            self.quick_date_end_entry.insert(0, picked_end)
            popup.destroy()

        ctk.CTkButton(
            actions,
            text="Cancelar",
            fg_color="transparent",
            border_width=1,
            text_color=("gray20", "gray90"),
            command=popup.destroy,
        ).pack(side="left")
        ctk.CTkButton(actions, text="Seleccionar", command=_confirm).pack(side="right")

    def _normalize_input_date(self, raw_value: str) -> str | None:
        text = raw_value.strip()
        if not text:
            return None

        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                parsed = datetime.strptime(text, fmt)
                return parsed.strftime("%Y-%m-%d")
            except ValueError:
                continue

        return None

    def add_manual_exception_from_fields(self) -> None:
        if self.exceptions_manual_box is None:
            return

        selected_label = self.quick_employee_menu.get().strip() if self.quick_employee_menu else "Todos"
        employee = self._employee_label_to_id.get(selected_label, "")
        date_start_value = self.quick_date_entry.get().strip() if self.quick_date_entry else ""
        date_end_value = self.quick_date_end_entry.get().strip() if self.quick_date_end_entry else ""
        exception_type = self.quick_type_menu.get().strip() if self.quick_type_menu else ""
        details = self.quick_detail_entry.get().strip() if self.quick_detail_entry else ""

        normalized_start = self._normalize_input_date(date_start_value)
        if normalized_start is None:
            messagebox.showwarning(
                "Fecha invalida",
                "Fecha inicio invalida. Usa YYYY-MM-DD o DD/MM/YYYY.",
            )
            return

        normalized_end = normalized_start
        if date_end_value:
            normalized_end = self._normalize_input_date(date_end_value)
            if normalized_end is None:
                messagebox.showwarning(
                    "Fecha invalida",
                    "Fecha fin invalida. Usa YYYY-MM-DD o DD/MM/YYYY.",
                )
                return

        start_dt = datetime.strptime(normalized_start, "%Y-%m-%d")
        end_dt = datetime.strptime(normalized_end, "%Y-%m-%d")
        if end_dt < start_dt:
            messagebox.showwarning(
                "Rango invalido",
                "La fecha fin no puede ser menor a la fecha inicio.",
            )
            return

        if not exception_type:
            messagebox.showwarning("Tipo faltante", "Selecciona un tipo de excepcion.")
            return

        current = self.exceptions_manual_box.get("1.0", "end").rstrip()
        if current and not current.endswith("\n"):
            self.exceptions_manual_box.insert("end", "\n")
        generated_lines: list[str] = []
        cursor = start_dt
        while cursor <= end_dt:
            new_line = f"{employee}|{cursor.strftime('%Y-%m-%d')}|{exception_type}|{details}".rstrip()
            self.exceptions_manual_box.insert("end", new_line + "\n")
            generated_lines.append(new_line)
            cursor += timedelta(days=1)
        self.exceptions_manual_box.see("end")

        if self.quick_date_entry is not None:
            self.quick_date_entry.delete(0, "end")
        if self.quick_date_end_entry is not None:
            self.quick_date_end_entry.delete(0, "end")
        if self.quick_detail_entry is not None:
            self.quick_detail_entry.delete(0, "end")
        if self.quick_employee_menu is not None and "Todos" in self._employee_selector_values:
            self.quick_employee_menu.set("Todos")

        self._refresh_exceptions_summary()
        self.log(
            f"Excepcion agregada en rango: {normalized_start} a {normalized_end} "
            f"({len(generated_lines)} dias)."
        )

    def remove_last_manual_exception_line(self) -> None:
        if self.exceptions_manual_box is None:
            return

        lines = self.exceptions_manual_box.get("1.0", "end").splitlines()
        while lines and (not lines[-1].strip() or lines[-1].strip().startswith("#")):
            lines.pop()
        if not lines:
            return

        removed = lines.pop()
        rebuilt = "\n".join(lines)
        if rebuilt:
            rebuilt += "\n"
        self.exceptions_manual_box.delete("1.0", "end")
        self.exceptions_manual_box.insert("1.0", rebuilt)
        self._refresh_exceptions_summary()
        self.log(f"Se removio la ultima linea manual: {removed}")

    def clear_manual_exceptions_box(self) -> None:
        self._reset_manual_box_to_template()
        self.log("Carga manual de excepciones reiniciada.")

    def _push_progress(self, percent: int, message: str) -> None:
        self._progress_queue.put(("progress", (percent, message)))

    def _worker_run(self, source: Path, exceptions_file: Path | None, manual_text: str) -> None:
        try:
            output = self.processor.process_file(
                source,
                progress_callback=self._push_progress,
                exceptions_file=exceptions_file,
                manual_exceptions_text=manual_text,
            )
            self._progress_queue.put(("done", output))
        except Exception as exc:
            self._progress_queue.put(("error", exc))

    def _poll_worker(self) -> None:
        finished = False
        while not self._progress_queue.empty():
            event, payload = self._progress_queue.get_nowait()
            if event == "progress":
                percent, text = payload
                self._set_progress(int(percent), str(text))
                self.log(str(text))
            elif event == "done":
                output = Path(payload)
                self.output_file = output
                self._set_progress(100, "Proceso completado.")
                self._set_status("completado", "ready")
                self.log(f"Reporte generado: {output}")
                self._refresh_kpis_from_output(output)
                self.open_button.configure(state="normal")
                self._set_processing_state(True)
                self._refresh_context_panel()
                messagebox.showinfo("Proceso completado", f"Excel generado en:\n{output}")
                finished = True
            elif event == "error":
                self._set_status("error", "error")
                self.log(f"Error: {payload}")
                self._set_processing_state(True)
                self._refresh_context_panel()
                messagebox.showerror("Error", str(payload))
                finished = True

        self._update_elapsed()

        if finished:
            self._worker_thread = None
            self._started_at = None
            return

        if self._worker_thread and self._worker_thread.is_alive():
            self.after(120, self._poll_worker)
        else:
            self._worker_thread = None
            self._started_at = None
            self._set_processing_state(True)

    def log(self, message: str) -> None:
        if self.log_box is None:
            return
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{message}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def clear_log(self) -> None:
        if self.log_box is None:
            return
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.insert("end", "Bitacora limpia.\n")
        self.log_box.configure(state="disabled")

    def open_output_file(self) -> None:
        if not self.output_file:
            return
        try:
            os.startfile(self.output_file)
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{exc}")

    def select_file(self) -> None:
        filepath = filedialog.askopenfilename(
            title="Seleccionar exportacion Hikvision",
            filetypes=[("Excel", "*.xlsx *.xls")],
        )
        if not filepath:
            return

        self.selected_file = Path(filepath)
        self.output_file = None
        self.open_button.configure(state="disabled")
        self._set_file_entry(str(self.selected_file))
        self._refresh_context_panel()

        stat = self.selected_file.stat()
        size_mb = stat.st_size / (1024 * 1024)
        modified = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M")

        self.file_meta.configure(
            text=(
                f"Nombre: {self.selected_file.name}\n"
                f"Tamanio: {size_mb:.2f} MB\n"
                f"Modificado: {modified}"
            )
        )

        self.process_button.configure(state="normal")
        self._set_status("archivo cargado", "ready")
        self._set_progress(0, "Listo para iniciar.")
        self.elapsed_text.configure(text="Tiempo: 00:00")
        self._refresh_kpis_from_source(self.selected_file)
        self._load_employee_options_from_source(self.selected_file)
        self.log(f"Archivo seleccionado: {self.selected_file}")

    def select_exceptions_file(self) -> None:
        filepath = filedialog.askopenfilename(
            title="Seleccionar archivo de excepciones",
            filetypes=[("Excepciones", "*.csv *.xlsx *.xls")],
        )
        if not filepath:
            return

        selected = Path(filepath)
        try:
            self.exceptions_file = ensure_exceptions_file(selected)
        except ExceptionConfigError as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._set_exceptions_file_entry(str(self.exceptions_file))
        self._refresh_exceptions_summary()

    def clear_exceptions_file(self) -> None:
        self._initialize_default_exceptions_file()
        self.log("Se restablecio el archivo de excepciones por defecto.")

    def save_exceptions_config(self) -> None:
        if self.exceptions_file is None:
            messagebox.showwarning("Atencion", "No hay archivo de excepciones seleccionado.")
            return
        try:
            added_count = self._persist_manual_exceptions()
        except ExceptionConfigError as exc:
            messagebox.showerror("Error", str(exc))
            return

        self._refresh_exceptions_summary()
        if added_count > 0:
            self.log(
                f"Configuracion guardada. Se agregaron {added_count} excepciones manuales en {self.exceptions_file.name}."
            )
        else:
            self.log("Configuracion guardada. No habia nuevas excepciones manuales para agregar.")

    def process_file(self) -> None:
        if not self.selected_file:
            messagebox.showwarning("Atencion", "Primero selecciona un archivo.")
            return
        if self._worker_thread and self._worker_thread.is_alive():
            return

        if self.exceptions_file is None:
            self._initialize_default_exceptions_file()

        try:
            added_count = self._persist_manual_exceptions()
        except ExceptionConfigError as exc:
            messagebox.showerror("Error", str(exc))
            return

        if added_count > 0 and self.exceptions_file is not None:
            self.log(
                f"Se guardaron {added_count} excepciones manuales en {self.exceptions_file.name} antes de procesar."
            )

        self.output_file = None
        self.open_button.configure(state="disabled")
        self._refresh_context_panel()
        self._set_status("procesando", "working")
        self._set_progress(2, "Preparando ejecucion...")
        self._started_at = time.perf_counter()
        self._set_processing_state(False)
        self.log("Iniciando procesamiento...")

        self._worker_thread = threading.Thread(
            target=self._worker_run,
            args=(self.selected_file, self.exceptions_file, ""),
            daemon=True,
        )
        self._worker_thread.start()
        self.after(120, self._poll_worker)
