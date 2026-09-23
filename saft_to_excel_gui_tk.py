"""
saft_to_excel_gui_tk.py -- the tkinter/ttk front-end for saft_to_excel.py.

Built on tkinter (ships inside Python itself, so a PyInstaller build of
this file needs no extra GUI toolkit bundled) with sv_ttk (a small,
pure-Python theme package) replacing stock ttk/tk's dated default styling
with a modern Windows look. See _enable_windows_dpi_awareness and
UI_FONT_SIZE for the DPI-awareness/scaling handling this needed.

The conversion runs in a child process (subprocess.Popen, not in-process),
so Cancelar can just kill it outright rather than needing cooperative
cancellation support inside convert() itself. The child process is this
same file, relaunched with `--convert <saft.xml> <output.xlsx>` -- works
identically run from source and once frozen by PyInstaller, where
sys.executable is the exe itself and there's no separate interpreter or
script file to invoke.

Usage:
    python saft_to_excel_gui_tk.py
"""

from __future__ import annotations

import os
import subprocess
import sys
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import sv_ttk

import saft_to_excel

# Suppresses the child conversion process's own console window popping up
# behind this GUI on Windows; harmless no-op signature on other platforms.
_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

# Set explicitly rather than left to Tk's own default (9pt) or its DPI
# auto-detection: `tk scaling` reported the correct value even in the
# frozen exe (measured: 2.667 on a 200%-scaled display, matching the
# display's real DPI) but the rendered result still didn't look right --
# controlling the point size directly here is a lot easier to tune by eye
# than chasing exactly why Tk's auto-scaling under-delivers. Segoe UI is
# Windows' own UI font. Settled at 8pt after a few rounds of trying it live.
UI_FONT_FAMILY = "Segoe UI"
UI_FONT_SIZE = 8


def _worker_command(saft_path: str, output_path: str) -> list[str]:
    """Args to relaunch this same file in --convert (no GUI) mode, for
    subprocess.Popen. Frozen (PyInstaller), sys.executable is this program
    itself, so no script path is passed; from source, sys.executable is
    the Python interpreter, which does need one.
    """
    args = ["--convert", saft_path, output_path]
    if getattr(sys, "frozen", False):
        return [sys.executable, *args]
    return [sys.executable, str(Path(__file__).resolve()), *args]


def _run_worker(saft_path: str, output_path: str) -> int:
    saft_to_excel.convert(saft_path, output_path)
    return 0


class MainWindow:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title("Conversor SAF-T → Excel")
        root.resizable(False, False)

        default_font = (UI_FONT_FAMILY, UI_FONT_SIZE)
        style = ttk.Style(root)
        style.configure(".", font=default_font)
        for widget_style in ("TLabel", "TButton", "TEntry", "TCheckbutton"):
            style.configure(widget_style, font=default_font)
        root.option_add("*Font", default_font)  # covers plain tk widgets too, e.g. messagebox dialogs

        self.process: subprocess.Popen | None = None
        self.cancel_requested = False
        self.start_time = 0.0
        self.output_path: Path | None = None

        self.saft_var = tk.StringVar(value="(nenhum)")
        self.output_var = tk.StringVar(value="(nenhuma)")
        self.open_after_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(
            value='Escolha o SAF-T e a pasta de destino e clique em "Gerar Excel".'
        )

        main = ttk.Frame(root, padding=10)
        main.grid(row=0, column=0, sticky="nsew")

        top_row = ttk.Frame(main)
        top_row.grid(row=0, column=0, sticky="w")

        saft_col = ttk.Frame(top_row)
        saft_col.grid(row=0, column=0, padx=(0, 10))
        ttk.Label(saft_col, text="Ficheiro SAF-T").grid(row=0, column=0, columnspan=2, sticky="w")
        self.saft_entry = ttk.Entry(saft_col, textvariable=self.saft_var, width=24, state="readonly")
        self.saft_entry.grid(row=1, column=0)
        ttk.Button(saft_col, text="...", width=3, command=self._browse_saft).grid(row=1, column=1, padx=(4, 0))

        output_col = ttk.Frame(top_row)
        output_col.grid(row=0, column=1, padx=(0, 10))
        ttk.Label(output_col, text="Pasta de destino").grid(row=0, column=0, columnspan=2, sticky="w")
        self.output_entry = ttk.Entry(output_col, textvariable=self.output_var, width=24, state="readonly")
        self.output_entry.grid(row=1, column=0)
        ttk.Button(output_col, text="...", width=3, command=self._browse_output).grid(row=1, column=1, padx=(4, 0))

        buttons_col = ttk.Frame(top_row)
        buttons_col.grid(row=0, column=2, sticky="n")
        self.generate_button = ttk.Button(buttons_col, text="Gerar Excel", command=self._on_generate)
        self.generate_button.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        self.cancel_button = ttk.Button(buttons_col, text="Cancelar", command=self._on_cancel, state="disabled")
        self.cancel_button.grid(row=1, column=0, sticky="ew")

        self.open_after_check = ttk.Checkbutton(
            main, text='Abrir ficheiro Excel quando terminar a conversão', variable=self.open_after_var,
        )
        self.open_after_check.grid(row=1, column=0, sticky="w", pady=(8, 0))

        # Always present, idle (stopped, empty) or running (started,
        # indeterminate) -- never removed from the layout, since toggling a
        # widget's visibility around a run caused a window-geometry shift
        # in testing.
        self.progress = ttk.Progressbar(main, mode="indeterminate")
        self.progress.grid(row=2, column=0, sticky="ew", pady=(8, 4))
        main.columnconfigure(0, weight=1)

        self.status_label = ttk.Label(main, textvariable=self.status_var, wraplength=650)
        self.status_label.grid(row=3, column=0, sticky="w")

    def _browse_saft(self) -> None:
        path = filedialog.askopenfilename(
            title="Escolher ficheiro SAF-T", filetypes=[("Ficheiros SAF-T", "*.xml")]
        )
        if not path:
            return
        self.saft_var.set(path)
        if self.output_var.get() in ("", "(nenhuma)"):
            self.output_var.set(str(Path(path).parent))

    def _browse_output(self) -> None:
        path = filedialog.askdirectory(title="Escolher pasta de destino")
        if path:
            self.output_var.set(path)

    def _on_generate(self) -> None:
        saft_path = self.saft_var.get().strip()
        output_dir = self.output_var.get().strip()
        if not saft_path or saft_path == "(nenhum)" or not output_dir or output_dir == "(nenhuma)":
            messagebox.showwarning("Dados em falta", "Escolha o ficheiro SAF-T e a pasta de destino.")
            return

        self.output_path = Path(output_dir) / (Path(saft_path).stem + ".xlsx")
        self.cancel_requested = False
        self.generate_button.config(state="disabled")
        self.cancel_button.config(state="normal")
        self.open_after_check.config(state="disabled")
        self.progress.start(12)
        self.status_var.set("A gerar...")
        self.start_time = time.monotonic()

        command = _worker_command(saft_path, str(self.output_path))
        self.process = subprocess.Popen(
            command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=_NO_WINDOW
        )
        self.root.after(200, self._poll_process)

    def _poll_process(self) -> None:
        if self.process is None:
            return
        exit_code = self.process.poll()
        if exit_code is None:
            elapsed = int(time.monotonic() - self.start_time)
            self.status_var.set(f"A gerar... ({elapsed}s)")
            self.root.after(500, self._poll_process)
            return
        self._on_finished(exit_code)

    def _on_cancel(self) -> None:
        if self.process is not None and self.process.poll() is None:
            self.cancel_requested = True
            self.process.terminate()

    def _on_finished(self, exit_code: int) -> None:
        self.progress.stop()
        self.generate_button.config(state="normal")
        self.cancel_button.config(state="disabled")
        self.open_after_check.config(state="normal")
        elapsed = int(time.monotonic() - self.start_time)

        if self.cancel_requested:
            self.status_var.set("Cancelado.")
            self._delete_partial_output()
        elif exit_code == 0:
            self.status_var.set(f"Concluído em {elapsed}s: {self.output_path.name}")
            if self.open_after_var.get():
                self._open_output_file()
        else:
            stderr = self.process.stderr.read().decode("utf-8", "replace") if self.process.stderr else ""
            self.status_var.set("Erro ao gerar o ficheiro.")
            messagebox.showerror("Erro", stderr[-2000:] or "Falha desconhecida.")
            self._delete_partial_output()

        self.process = None

    def _open_output_file(self) -> None:
        if self.output_path is None:
            return
        try:
            os.startfile(self.output_path)  # noqa: S606 -- opens with the OS-registered default app (Excel)
        except OSError:
            pass

    def _delete_partial_output(self) -> None:
        # A killed or failed run can leave a truncated .xlsx behind (the
        # child process writes straight to this path, not atomically) --
        # best-effort cleanup so a half-written file doesn't sit there
        # looking like a real result.
        if self.output_path is None:
            return
        try:
            if self.output_path.exists():
                self.output_path.unlink()
        except OSError:
            pass


def _enable_windows_dpi_awareness() -> int:
    """Tcl/Tk doesn't declare itself DPI-aware, so on any >100%-scaled
    Windows display (the default on most machines today) Windows silently
    bitmap-stretches the whole window to compensate -- the blurry look.
    Qt/PySide6 declares awareness itself, which is why only this tkinter
    build needs this. Telling Windows "I'll handle scaling myself" turns
    that off -- but then Tk itself still assumes a plain 96 DPI (its
    default 1.333 pixels-per-point scale) unless told the real value, which
    is what made everything render crisp but *too small* right after this
    was added: stopping the blurry stretch without also telling Tk the
    real DPI just trades one wrong size for another. Returns the detected
    system DPI (96 if detection fails or this isn't Windows) for main() to
    pass to `tk scaling`.
    """
    if sys.platform != "win32":
        return 96
    import ctypes

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()  # fallback for older Windows
        except Exception:
            pass
    try:
        return ctypes.windll.user32.GetDpiForSystem()  # Windows 10 1607+
    except Exception:
        return 96


def main() -> int:
    if len(sys.argv) >= 4 and sys.argv[1] == "--convert":
        return _run_worker(sys.argv[2], sys.argv[3])

    dpi = _enable_windows_dpi_awareness()
    root = tk.Tk()
    root.tk.call("tk", "scaling", dpi / 72.0)
    sv_ttk.set_theme("light")  # the ordinary grey Windows look, not forced dark
    MainWindow(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
