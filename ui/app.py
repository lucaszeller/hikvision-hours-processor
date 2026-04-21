from __future__ import annotations

from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog, messagebox

from services.processor import ProcessorService


class HikvisionApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Hikvision Hours Processor")
        self.geometry("760x480")

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.processor = ProcessorService()
        self.selected_file: Path | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        title = ctk.CTkLabel(self, text="Procesador de fichadas Hikvision", font=ctk.CTkFont(size=22, weight="bold"))
        title.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        actions = ctk.CTkFrame(self)
        actions.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        actions.grid_columnconfigure(1, weight=1)

        self.select_button = ctk.CTkButton(actions, text="Seleccionar archivo", command=self.select_file)
        self.select_button.grid(row=0, column=0, padx=10, pady=10)

        self.file_label = ctk.CTkLabel(actions, text="Ningún archivo seleccionado", anchor="w")
        self.file_label.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        self.process_button = ctk.CTkButton(actions, text="Procesar", command=self.process_file, state="disabled")
        self.process_button.grid(row=0, column=2, padx=10, pady=10)

        self.log_box = ctk.CTkTextbox(self)
        self.log_box.grid(row=3, column=0, padx=20, pady=(10, 20), sticky="nsew")
        self.log_box.insert("end", "Listo para procesar fichadas.\n")
        self.log_box.configure(state="disabled")

    def log(self, message: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{message}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def select_file(self) -> None:
        filepath = filedialog.askopenfilename(
            title="Seleccionar exportación Hikvision",
            filetypes=[("Excel", "*.xlsx *.xls")],
        )
        if not filepath:
            return

        self.selected_file = Path(filepath)
        self.file_label.configure(text=str(self.selected_file))
        self.process_button.configure(state="normal")
        self.log(f"Archivo seleccionado: {self.selected_file}")

    def process_file(self) -> None:
        if not self.selected_file:
            messagebox.showwarning("Atención", "Primero seleccioná un archivo.")
            return

        self.log("Iniciando procesamiento...")
        try:
            output = self.processor.process_file(self.selected_file)
            self.log(f"Reporte generado: {output}")
            messagebox.showinfo("Proceso completado", f"Excel generado en:\n{output}")
        except Exception as exc:
            self.log(f"Error: {exc}")
            messagebox.showerror("Error", str(exc))
