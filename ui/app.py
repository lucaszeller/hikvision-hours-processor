from __future__ import annotations

import os
import queue
import threading
import time
from datetime import datetime
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog, messagebox

from services.processor import ProcessorService


class HikvisionApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Hikvision Hours Processor")
        self.geometry("980x620")
        self.minsize(920, 560)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("dark-blue")

        self.processor = ProcessorService()
        self.selected_file: Path | None = None
        self.output_file: Path | None = None

        self._worker_thread: threading.Thread | None = None
        self._progress_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._started_at: float | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        sidebar = ctk.CTkFrame(self, width=320, corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nsw")
        sidebar.grid_propagate(False)
        sidebar.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            sidebar,
            text="Procesador Hikvision",
            font=ctk.CTkFont(size=24, weight="bold"),
        )
        title.grid(row=0, column=0, padx=20, pady=(26, 8), sticky="w")

        subtitle = ctk.CTkLabel(
            sidebar,
            text="Importa fichadas y genera el reporte\ncon horas diarias y resumen mensual.",
            justify="left",
            text_color=("gray20", "gray70"),
        )
        subtitle.grid(row=1, column=0, padx=20, pady=(0, 18), sticky="w")

        self.status_badge = ctk.CTkLabel(
            sidebar,
            text="Estado: listo",
            corner_radius=8,
            fg_color=("#D9F2E5", "#153828"),
            text_color=("#14532D", "#A7F3D0"),
            padx=10,
            pady=6,
        )
        self.status_badge.grid(row=2, column=0, padx=20, pady=(0, 18), sticky="w")

        select_label = ctk.CTkLabel(
            sidebar, text="Archivo de entrada", font=ctk.CTkFont(size=13, weight="bold")
        )
        select_label.grid(row=3, column=0, padx=20, pady=(0, 8), sticky="w")

        self.file_entry = ctk.CTkEntry(sidebar, placeholder_text="Ningun archivo seleccionado")
        self.file_entry.grid(row=4, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.file_entry.configure(state="disabled")

        self.file_meta = ctk.CTkLabel(
            sidebar,
            text="",
            justify="left",
            text_color=("gray35", "gray65"),
        )
        self.file_meta.grid(row=5, column=0, padx=20, pady=(0, 18), sticky="w")

        self.select_button = ctk.CTkButton(
            sidebar,
            text="Seleccionar archivo",
            height=40,
            command=self.select_file,
        )
        self.select_button.grid(row=6, column=0, padx=20, pady=(0, 10), sticky="ew")

        self.process_button = ctk.CTkButton(
            sidebar,
            text="Procesar reporte",
            height=40,
            command=self.process_file,
            state="disabled",
        )
        self.process_button.grid(row=7, column=0, padx=20, pady=(0, 10), sticky="ew")

        self.open_button = ctk.CTkButton(
            sidebar,
            text="Abrir reporte generado",
            height=38,
            command=self.open_output_file,
            state="disabled",
        )
        self.open_button.grid(row=8, column=0, padx=20, pady=(0, 10), sticky="ew")

        self.clear_button = ctk.CTkButton(
            sidebar,
            text="Limpiar bitacora",
            height=38,
            fg_color="transparent",
            border_width=1,
            text_color=("gray20", "gray90"),
            command=self.clear_log,
        )
        self.clear_button.grid(row=9, column=0, padx=20, pady=(0, 18), sticky="ew")

        self.progress = ctk.CTkProgressBar(sidebar, mode="determinate")
        self.progress.grid(row=10, column=0, padx=20, pady=(0, 8), sticky="ew")
        self.progress.set(0)

        self.progress_text = ctk.CTkLabel(sidebar, text="Progreso: 0%", text_color=("gray35", "gray65"))
        self.progress_text.grid(row=11, column=0, padx=20, pady=(0, 4), sticky="w")

        self.progress_step = ctk.CTkLabel(sidebar, text="Paso: esperando archivo", text_color=("gray35", "gray65"))
        self.progress_step.grid(row=12, column=0, padx=20, pady=(0, 4), sticky="w")

        self.elapsed_text = ctk.CTkLabel(sidebar, text="Tiempo: 00:00", text_color=("gray35", "gray65"))
        self.elapsed_text.grid(row=13, column=0, padx=20, pady=(0, 14), sticky="w")

        info = ctk.CTkLabel(
            sidebar,
            text="Formato compatible: .xls / .xlsx",
            text_color=("gray35", "gray65"),
        )
        info.grid(row=14, column=0, padx=20, pady=(0, 16), sticky="w")

        main = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew", padx=(16, 18), pady=18)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(2, weight=1)

        header = ctk.CTkLabel(
            main,
            text="Bitacora de ejecucion",
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        header.grid(row=0, column=0, sticky="w", pady=(4, 6))

        helper = ctk.CTkLabel(
            main,
            text="Aca vas a ver cada paso del procesamiento y cualquier error detectado.",
            text_color=("gray35", "gray65"),
        )
        helper.grid(row=1, column=0, sticky="w", pady=(0, 12))

        self.log_box = ctk.CTkTextbox(main, corner_radius=10, border_width=1)
        self.log_box.grid(row=2, column=0, sticky="nsew")
        self.log_box.insert("end", "Listo para procesar fichadas.\n")
        self.log_box.configure(state="disabled")

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

    def _set_processing_state(self, enabled: bool) -> None:
        self.select_button.configure(state="normal" if enabled else "disabled")
        self.process_button.configure(state="normal" if enabled and self.selected_file else "disabled")
        self.clear_button.configure(state="normal" if enabled else "disabled")

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

    def _push_progress(self, percent: int, message: str) -> None:
        self._progress_queue.put(("progress", (percent, message)))

    def _worker_run(self, source: Path) -> None:
        try:
            output = self.processor.process_file(source, progress_callback=self._push_progress)
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
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{message}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def clear_log(self) -> None:
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
        self.log(f"Archivo seleccionado: {self.selected_file}")

    def process_file(self) -> None:
        if not self.selected_file:
            messagebox.showwarning("Atencion", "Primero selecciona un archivo.")
            return
        if self._worker_thread and self._worker_thread.is_alive():
            return

        self.output_file = None
        self.open_button.configure(state="disabled")
        self._set_status("procesando", "working")
        self._set_progress(2, "Preparando ejecucion...")
        self._started_at = time.perf_counter()
        self._set_processing_state(False)
        self.log("Iniciando procesamiento...")

        self._worker_thread = threading.Thread(
            target=self._worker_run,
            args=(self.selected_file,),
            daemon=True,
        )
        self._worker_thread.start()
        self.after(120, self._poll_worker)
