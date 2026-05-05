from __future__ import annotations

import os
import queue
import threading
import time
from datetime import datetime
import math
from pathlib import Path

import customtkinter as ctk
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


class HikvisionApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Hikvision Hours Processor")
        self.geometry("1140x720")
        self.minsize(1020, 620)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("dark-blue")

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
        self.quick_date_button: ctk.CTkButton | None = None
        self.quick_type_menu: ctk.CTkOptionMenu | None = None
        self.quick_detail_entry: ctk.CTkEntry | None = None
        self.quick_add_button: ctk.CTkButton | None = None
        self.quick_remove_last_button: ctk.CTkButton | None = None
        self.quick_clear_manual_button: ctk.CTkButton | None = None
        self.exceptions_summary_label: ctk.CTkLabel | None = None
        self._logo_image: tk.PhotoImage | None = None
        self._logo_badge: ctk.CTkFrame | None = None
        self._logo_label: ctk.CTkLabel | None = None
        self._employee_selector_values: list[str] = ["Todos"]
        self._employee_label_to_id: dict[str, str] = {"Todos": ""}

        self._build_ui()
        self._initialize_default_exceptions_file()

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.navbar = ctk.CTkFrame(self, corner_radius=0, height=60, fg_color=("#F4F6F8", "#111827"))
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
            text_color=("#111827", "#E5E7EB"),
        )
        self.nav_title.grid(row=0, column=0, padx=(14, 18), pady=(10, 8), sticky="w")

        self.report_tab_button = ctk.CTkButton(
            self.navbar,
            text="Reporte",
            width=110,
            height=32,
            corner_radius=6,
            fg_color="transparent",
            hover_color=("#E8EEF5", "#1F2937"),
            text_color=("#374151", "#D1D5DB"),
            font=ctk.CTkFont(size=13, weight="bold"),
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
            hover_color=("#E8EEF5", "#1F2937"),
            text_color=("#374151", "#D1D5DB"),
            font=ctk.CTkFont(size=13, weight="bold"),
            command=lambda: self._show_page("exceptions"),
        )
        self.exceptions_tab_button.grid(row=0, column=2, padx=(0, 10), pady=(12, 8), sticky="w")

        self.report_indicator = ctk.CTkFrame(
            self.navbar,
            height=3,
            width=110,
            corner_radius=0,
            fg_color=("#1F6AA5", "#60A5FA"),
        )
        self.report_indicator.grid(row=1, column=1, padx=(0, 6), sticky="sw")

        self.exceptions_indicator = ctk.CTkFrame(
            self.navbar,
            height=3,
            width=120,
            corner_radius=0,
            fg_color=("#1F6AA5", "#60A5FA"),
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

        self._refresh_exceptions_summary()
        self._show_page("report")

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
            border_color=("#D4DCE6", "#374151"),
            fg_color=("#FFFFFF", "#0F172A"),
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

        sidebar = ctk.CTkFrame(page, width=360, corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nsw")
        sidebar.grid_propagate(False)
        sidebar.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            sidebar,
            text="Procesador Hikvision",
            font=ctk.CTkFont(size=24, weight="bold"),
        )
        title.grid(row=0, column=0, padx=20, pady=(22, 8), sticky="w")

        subtitle = ctk.CTkLabel(
            sidebar,
            text="Importa fichadas y genera reportes\ndiarios/mensuales.",
            justify="left",
            text_color=("gray20", "gray70"),
        )
        subtitle.grid(row=1, column=0, padx=20, pady=(0, 16), sticky="w")

        self.status_badge = ctk.CTkLabel(
            sidebar,
            text="Estado: listo",
            corner_radius=8,
            fg_color=("#D9F2E5", "#153828"),
            text_color=("#14532D", "#A7F3D0"),
            padx=10,
            pady=6,
        )
        self.status_badge.grid(row=2, column=0, padx=20, pady=(0, 14), sticky="w")

        select_label = ctk.CTkLabel(
            sidebar, text="Archivo de entrada", font=ctk.CTkFont(size=13, weight="bold")
        )
        select_label.grid(row=3, column=0, padx=20, pady=(0, 8), sticky="w")

        self.file_entry = ctk.CTkEntry(sidebar, placeholder_text="Ningun archivo seleccionado")
        self.file_entry.grid(row=4, column=0, padx=20, pady=(0, 8), sticky="ew")
        self.file_entry.configure(state="disabled")

        self.file_meta = ctk.CTkLabel(
            sidebar,
            text="",
            justify="left",
            text_color=("gray35", "gray65"),
        )
        self.file_meta.grid(row=5, column=0, padx=20, pady=(0, 12), sticky="w")

        self.select_button = ctk.CTkButton(
            sidebar,
            text="Seleccionar archivo",
            height=38,
            command=self.select_file,
        )
        self.select_button.grid(row=6, column=0, padx=20, pady=(0, 10), sticky="ew")

        self.exceptions_summary_label = ctk.CTkLabel(
            sidebar,
            text="Excepciones: sin configurar",
            justify="left",
            text_color=("gray35", "gray65"),
        )
        self.exceptions_summary_label.grid(row=7, column=0, padx=20, pady=(0, 8), sticky="w")

        goto_exceptions = ctk.CTkButton(
            sidebar,
            text="Ir a Excepciones",
            height=34,
            fg_color="transparent",
            border_width=1,
            text_color=("gray20", "gray90"),
            command=lambda: self._show_page("exceptions"),
        )
        goto_exceptions.grid(row=8, column=0, padx=20, pady=(0, 10), sticky="ew")

        self.process_button = ctk.CTkButton(
            sidebar,
            text="Procesar reporte",
            height=40,
            command=self.process_file,
            state="disabled",
        )
        self.process_button.grid(row=9, column=0, padx=20, pady=(2, 10), sticky="ew")

        self.open_button = ctk.CTkButton(
            sidebar,
            text="Abrir reporte generado",
            height=36,
            command=self.open_output_file,
            state="disabled",
        )
        self.open_button.grid(row=10, column=0, padx=20, pady=(0, 8), sticky="ew")

        self.clear_button = ctk.CTkButton(
            sidebar,
            text="Limpiar bitacora",
            height=36,
            fg_color="transparent",
            border_width=1,
            text_color=("gray20", "gray90"),
            command=self.clear_log,
        )
        self.clear_button.grid(row=11, column=0, padx=20, pady=(0, 14), sticky="ew")

        self.progress = ctk.CTkProgressBar(sidebar, mode="determinate")
        self.progress.grid(row=12, column=0, padx=20, pady=(0, 8), sticky="ew")
        self.progress.set(0)

        self.progress_text = ctk.CTkLabel(sidebar, text="Progreso: 0%", text_color=("gray35", "gray65"))
        self.progress_text.grid(row=13, column=0, padx=20, pady=(0, 4), sticky="w")

        self.progress_step = ctk.CTkLabel(
            sidebar,
            text="Paso: esperando archivo",
            text_color=("gray35", "gray65"),
        )
        self.progress_step.grid(row=14, column=0, padx=20, pady=(0, 4), sticky="w")

        self.elapsed_text = ctk.CTkLabel(sidebar, text="Tiempo: 00:00", text_color=("gray35", "gray65"))
        self.elapsed_text.grid(row=15, column=0, padx=20, pady=(0, 10), sticky="w")

        info = ctk.CTkLabel(
            sidebar,
            text="Entrada: .xls/.xlsx",
            text_color=("gray35", "gray65"),
            justify="left",
        )
        info.grid(row=16, column=0, padx=20, pady=(0, 12), sticky="w")

        main = ctk.CTkFrame(page, corner_radius=0, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew", padx=(16, 18), pady=18)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(4, weight=1)

        header = ctk.CTkLabel(
            main,
            text="Bitacora de ejecucion",
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        header.grid(row=0, column=0, sticky="w", pady=(2, 6))

        helper = ctk.CTkLabel(
            main,
            text="Aca vas a ver cada paso del procesamiento y errores detectados.",
            text_color=("gray35", "gray65"),
        )
        helper.grid(row=1, column=0, sticky="w", pady=(0, 8))

        self.log_box = ctk.CTkTextbox(main, corner_radius=10, border_width=1)
        self.log_box.grid(row=2, column=0, sticky="nsew")
        self.log_box.insert("end", "Listo para procesar fichadas.\n")
        self.log_box.configure(state="disabled")

        watermark = ctk.CTkLabel(
            page,
            text="PORTA",
            font=ctk.CTkFont(size=116, weight="bold"),
            text_color=("#E7ECF4", "#202A38"),
        )
        watermark.place(relx=0.64, rely=0.52, anchor="center")
        watermark.lower()

        return page

    def _build_exceptions_page(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        page = ctk.CTkFrame(parent, corner_radius=0, fg_color="transparent")
        page.grid_columnconfigure(0, weight=0)
        page.grid_columnconfigure(1, weight=1)
        page.grid_rowconfigure(0, weight=1)

        sidebar = ctk.CTkFrame(page, width=360, corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nsw")
        sidebar.grid_propagate(False)
        sidebar.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            sidebar,
            text="Excepciones",
            font=ctk.CTkFont(size=24, weight="bold"),
        )
        title.grid(row=0, column=0, padx=20, pady=(22, 8), sticky="w")

        subtitle = ctk.CTkLabel(
            sidebar,
            text="Configura feriados, vacaciones,\nenfermedad y permisos.",
            justify="left",
            text_color=("gray20", "gray70"),
        )
        subtitle.grid(row=1, column=0, padx=20, pady=(0, 16), sticky="w")

        file_label = ctk.CTkLabel(
            sidebar,
            text="Archivo de excepciones",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        file_label.grid(row=2, column=0, padx=20, pady=(0, 8), sticky="w")

        self.exceptions_file_entry = ctk.CTkEntry(sidebar, placeholder_text="Sin archivo de excepciones")
        self.exceptions_file_entry.grid(row=3, column=0, padx=20, pady=(0, 8), sticky="ew")
        self.exceptions_file_entry.configure(state="disabled")

        self.select_exceptions_button = ctk.CTkButton(
            sidebar,
            text="Seleccionar archivo",
            height=34,
            command=self.select_exceptions_file,
        )
        self.select_exceptions_button.grid(row=4, column=0, padx=20, pady=(0, 8), sticky="ew")

        self.clear_exceptions_button = ctk.CTkButton(
            sidebar,
            text="Usar archivo por defecto",
            height=34,
            fg_color="transparent",
            border_width=1,
            text_color=("gray20", "gray90"),
            command=self.clear_exceptions_file,
        )
        self.clear_exceptions_button.grid(row=5, column=0, padx=20, pady=(0, 8), sticky="ew")

        self.save_exceptions_button = ctk.CTkButton(
            sidebar,
            text="Guardar configuracion",
            height=36,
            command=self.save_exceptions_config,
        )
        self.save_exceptions_button.grid(row=6, column=0, padx=20, pady=(6, 8), sticky="ew")

        go_report = ctk.CTkButton(
            sidebar,
            text="Volver a Reporte",
            height=34,
            fg_color="transparent",
            border_width=1,
            text_color=("gray20", "gray90"),
            command=lambda: self._show_page("report"),
        )
        go_report.grid(row=7, column=0, padx=20, pady=(0, 8), sticky="ew")

        info = ctk.CTkLabel(
            sidebar,
            text="Formatos: .csv/.xls/.xlsx",
            text_color=("gray35", "gray65"),
            justify="left",
        )
        info.grid(row=8, column=0, padx=20, pady=(0, 12), sticky="w")

        main = ctk.CTkFrame(page, corner_radius=0, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew", padx=(16, 18), pady=18)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(2, weight=1)

        header = ctk.CTkLabel(
            main,
            text="Carga Manual",
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        header.grid(row=0, column=0, sticky="w", pady=(2, 6))

        helper = ctk.CTkLabel(
            main,
            text=(
                "Carga rapida por campos o pega varias lineas.\n"
                "Formato manual: ID|YYYY-MM-DD|TIPO|DETALLE"
            ),
            justify="left",
            text_color=("gray35", "gray65"),
        )
        helper.grid(row=1, column=0, sticky="w", pady=(0, 8))

        quick_card = ctk.CTkFrame(main, corner_radius=10, border_width=1)
        quick_card.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        quick_card.grid_columnconfigure(0, weight=1)
        quick_card.grid_columnconfigure(1, weight=1)
        quick_card.grid_columnconfigure(2, weight=1)
        quick_card.grid_columnconfigure(3, weight=2)

        quick_title = ctk.CTkLabel(
            quick_card,
            text="Carga Rapida",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        quick_title.grid(row=0, column=0, columnspan=4, sticky="w", padx=12, pady=(10, 8))

        self.quick_employee_menu = ctk.CTkOptionMenu(
            quick_card,
            values=self._employee_selector_values,
        )
        self.quick_employee_menu.grid(row=1, column=0, padx=(12, 6), pady=(0, 8), sticky="ew")
        self.quick_employee_menu.set("Todos")

        date_row = ctk.CTkFrame(quick_card, fg_color="transparent")
        date_row.grid(row=1, column=1, padx=6, pady=(0, 8), sticky="ew")
        date_row.grid_columnconfigure(0, weight=1)
        date_row.grid_columnconfigure(1, weight=0)

        self.quick_date_entry = ctk.CTkEntry(date_row, placeholder_text="Fecha (YYYY-MM-DD o DD/MM/YYYY)")
        self.quick_date_entry.grid(row=0, column=0, sticky="ew")

        self.quick_date_button = ctk.CTkButton(
            date_row,
            text="Calendario",
            width=96,
            height=28,
            command=self.open_date_picker,
        )
        self.quick_date_button.grid(row=0, column=1, padx=(6, 0), sticky="e")

        self.quick_type_menu = ctk.CTkOptionMenu(
            quick_card,
            values=["Feriado", "Vacaciones", "Enfermedad", "Permiso", "Ausencia justificada", "Otro"],
        )
        self.quick_type_menu.grid(row=1, column=2, padx=6, pady=(0, 8), sticky="ew")
        self.quick_type_menu.set("Feriado")

        self.quick_detail_entry = ctk.CTkEntry(quick_card, placeholder_text="Detalle (opcional)")
        self.quick_detail_entry.grid(row=1, column=3, padx=(6, 12), pady=(0, 8), sticky="ew")

        self.quick_add_button = ctk.CTkButton(
            quick_card,
            text="Agregar a carga manual",
            height=32,
            command=self.add_manual_exception_from_fields,
        )
        self.quick_add_button.grid(row=2, column=0, columnspan=2, padx=(12, 6), pady=(0, 10), sticky="ew")

        self.quick_remove_last_button = ctk.CTkButton(
            quick_card,
            text="Quitar ultima linea",
            height=32,
            fg_color="transparent",
            border_width=1,
            text_color=("gray20", "gray90"),
            command=self.remove_last_manual_exception_line,
        )
        self.quick_remove_last_button.grid(row=2, column=2, padx=6, pady=(0, 10), sticky="ew")

        self.quick_clear_manual_button = ctk.CTkButton(
            quick_card,
            text="Limpiar manual",
            height=32,
            fg_color="transparent",
            border_width=1,
            text_color=("gray20", "gray90"),
            command=self.clear_manual_exceptions_box,
        )
        self.quick_clear_manual_button.grid(row=2, column=3, padx=(6, 12), pady=(0, 10), sticky="ew")

        self.exceptions_manual_box = ctk.CTkTextbox(main, corner_radius=10, border_width=1)
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
            text_color=("#E7ECF4", "#202A38"),
        )
        watermark.place(relx=0.64, rely=0.52, anchor="center")
        watermark.lower()

        return page

    def _show_page(self, page: str) -> None:
        self.current_page = page
        active_tab_fg = ("#E5EEF7", "#1E293B")
        inactive_tab_fg = "transparent"
        active_text = ("#0F4C81", "#93C5FD")
        inactive_text = ("#374151", "#D1D5DB")
        active_line = ("#1F6AA5", "#60A5FA")
        hidden_line = ("#F4F6F8", "#111827")

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
            "ready": {"fg": ("#D9F2E5", "#153828"), "text": ("#14532D", "#A7F3D0")},
            "working": {"fg": ("#FEF3C7", "#3F3208"), "text": ("#92400E", "#FDE68A")},
            "error": {"fg": ("#FEE2E2", "#3D1111"), "text": ("#991B1B", "#FCA5A5")},
        }
        colors = palette.get(level, palette["ready"])
        self.status_badge.configure(
            text=f"Estado: {text}",
            fg_color=colors["fg"],
            text_color=colors["text"],
        )

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

    def _load_employee_options_from_source(self, source_path: Path) -> None:
        self._employee_label_to_id = {"Todos": ""}
        labels = ["Todos"]

        try:
            df = load_hikvision_excel(source_path)
            if not df.empty:
                employees = (
                    df[["employee_id", "employee_name"]]
                    .dropna()
                    .astype(str)
                    .assign(
                        employee_id=lambda x: x["employee_id"].str.strip(),
                        employee_name=lambda x: x["employee_name"].str.strip(),
                    )
                )
                employees = employees[
                    (employees["employee_id"] != "") & (employees["employee_name"] != "")
                ].drop_duplicates(subset=["employee_id"])
                employees = employees.sort_values(["employee_name", "employee_id"], kind="stable")

                used_labels: set[str] = set()
                for _, row in employees.iterrows():
                    emp_id = row["employee_id"]
                    emp_name = row["employee_name"]
                    label = emp_name
                    if label in used_labels:
                        label = f"{emp_name} ({emp_id})"
                    used_labels.add(label)
                    self._employee_label_to_id[label] = emp_id
                    labels.append(label)
        except Exception as exc:
            self.log(f"No se pudo cargar lista de trabajadores para excepciones: {exc}")

        self._set_employee_selector_values(labels)

    def open_date_picker(self) -> None:
        if self.quick_date_entry is None:
            return

        if Calendar is None:
            messagebox.showwarning(
                "Calendario no disponible",
                "Instala dependencia 'tkcalendar' para usar el selector visual de fecha.",
            )
            return

        seed = self._normalize_input_date(self.quick_date_entry.get().strip() if self.quick_date_entry else "")
        if seed:
            seed_dt = datetime.strptime(seed, "%Y-%m-%d")
        else:
            seed_dt = datetime.now()

        popup = ctk.CTkToplevel(self)
        popup.title("Seleccionar fecha")
        popup.geometry("320x360")
        popup.resizable(False, False)
        popup.transient(self)
        popup.grab_set()

        body = ctk.CTkFrame(popup, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=12, pady=12)

        calendar = Calendar(
            body,
            selectmode="day",
            year=seed_dt.year,
            month=seed_dt.month,
            day=seed_dt.day,
            date_pattern="yyyy-mm-dd",
        )
        calendar.pack(fill="both", expand=True)

        actions = ctk.CTkFrame(body, fg_color="transparent")
        actions.pack(fill="x", pady=(10, 0))

        def _confirm() -> None:
            picked = calendar.get_date()
            if self.quick_date_entry is not None:
                self.quick_date_entry.delete(0, "end")
                self.quick_date_entry.insert(0, picked)
            popup.destroy()

        ctk.CTkButton(actions, text="Cancelar", fg_color="transparent", border_width=1, text_color=("gray20", "gray90"), command=popup.destroy).pack(side="left")
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
        date_value = self.quick_date_entry.get().strip() if self.quick_date_entry else ""
        exception_type = self.quick_type_menu.get().strip() if self.quick_type_menu else ""
        details = self.quick_detail_entry.get().strip() if self.quick_detail_entry else ""

        normalized_date = self._normalize_input_date(date_value)
        if normalized_date is None:
            messagebox.showwarning(
                "Fecha invalida",
                "Usa YYYY-MM-DD o DD/MM/YYYY.",
            )
            return

        if not exception_type:
            messagebox.showwarning("Tipo faltante", "Selecciona un tipo de excepcion.")
            return

        new_line = f"{employee}|{normalized_date}|{exception_type}|{details}".rstrip()
        current = self.exceptions_manual_box.get("1.0", "end").rstrip()
        if current and not current.endswith("\n"):
            self.exceptions_manual_box.insert("end", "\n")
        self.exceptions_manual_box.insert("end", new_line + "\n")
        self.exceptions_manual_box.see("end")

        if self.quick_date_entry is not None:
            self.quick_date_entry.delete(0, "end")
        if self.quick_detail_entry is not None:
            self.quick_detail_entry.delete(0, "end")
        if self.quick_employee_menu is not None and "Todos" in self._employee_selector_values:
            self.quick_employee_menu.set("Todos")

        self._refresh_exceptions_summary()
        self.log(f"Excepcion agregada: {new_line}")

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
                self.open_button.configure(state="normal")
                self._set_processing_state(True)
                messagebox.showinfo("Proceso completado", f"Excel generado en:\n{output}")
                finished = True
            elif event == "error":
                self._set_status("error", "error")
                self.log(f"Error: {payload}")
                self._set_processing_state(True)
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
