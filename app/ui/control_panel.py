"""
Left-side control panel: field list, action buttons, progress, log.
"""
import tkinter as tk
from tkinter import ttk

from app.constants import C
from app.ui.widgets import flat_button, label, separator


class ControlPanel(tk.Frame):
    """
    Owns:
      self.fields_frame   -- inner Frame where field rows are injected
      self.progress       -- ttk.Progressbar
      self.info_text      -- tk.Text (log)
    """

    def __init__(self, parent, preview_cmd, generate_cmd):
        super().__init__(
            parent, width=300, bg=C["surface"],
            highlightthickness=1, highlightbackground=C["border"],
        )
        self.pack(side="left", fill="y", padx=(0, 10), pady=6)
        self.pack_propagate(False)
        self._build(preview_cmd, generate_cmd)

    def _build(self, preview_cmd, generate_cmd):
        # Header
        hdr = tk.Frame(self, bg=C["surface"], pady=10)
        hdr.pack(fill="x", padx=14)
        label(hdr, "Certificate Fields", font_size=10, bold=True,
              bg=C["surface"]).pack(side="left")
        separator(self)

        # Scrollable field list
        self.fields_frame = tk.Frame(self, bg=C["surface"])
        self.fields_frame.pack(fill="x")

        separator(self, padx=14)

        # Preview / Generate buttons
        btn_row = tk.Frame(self, bg=C["surface"], pady=10)
        btn_row.pack(fill="x", padx=14)
        flat_button(
            btn_row, "Preview", preview_cmd,
            C["success"], C["success2"],
            font_size=9, bold=True, padx=14, pady=7,
        ).pack(side="left", fill="x", expand=True, padx=(0, 5))
        flat_button(
            btn_row, "Generate", generate_cmd,
            C["accent"], C["accent2"],
            font_size=9, bold=True, padx=14, pady=7,
        ).pack(side="left", fill="x", expand=True)

        # Progress bar
        prog = tk.Frame(self, bg=C["surface"])
        prog.pack(fill="x", padx=14, pady=(0, 8))
        self.progress = ttk.Progressbar(
            prog, orient="horizontal", mode="determinate",
            style="Thin.Horizontal.TProgressbar",
        )
        self.progress.pack(fill="x")

        separator(self)

        # Log header
        log_hdr = tk.Frame(self, bg=C["surface"], pady=8)
        log_hdr.pack(fill="x", padx=14)
        label(log_hdr, "Generation Log", font_size=9, bold=True,
              color=C["subtext"], bg=C["surface"]).pack(side="left")

        # Log text area
        log_wrap = tk.Frame(self, bg=C["surface"])
        log_wrap.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        self.info_text = tk.Text(
            log_wrap, height=8, wrap=tk.WORD,
            font=("Consolas", 8), bg=C["log_bg"], fg=C["subtext"],
            relief="flat", bd=0, padx=8, pady=6, state="disabled",
            highlightthickness=1, highlightbackground=C["border"],
        )
        self.info_text.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(
            log_wrap, orient="vertical",
            command=self.info_text.yview,
            style="Flat.Vertical.TScrollbar",
        )
        sb.pack(side="right", fill="y")
        self.info_text.configure(yscrollcommand=sb.set)

    def append_log(self, msg: str, clear: bool = False) -> None:
        self.info_text.configure(state="normal")
        if clear:
            self.info_text.delete("1.0", tk.END)
        self.info_text.insert(tk.END, msg + "\n")
        self.info_text.see(tk.END)
        self.info_text.configure(state="disabled")

    def set_progress(self, pct: float) -> None:
        self.progress.configure(value=pct)
