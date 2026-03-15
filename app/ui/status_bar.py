"""
Thin status bar that sits at the bottom of the window.
"""
import tkinter as tk

from app.constants import C


class StatusBar(tk.Frame):

    def __init__(self, parent):
        super().__init__(parent, bg=C["nav"], height=24)
        self.pack(side="bottom", fill="x")
        self.pack_propagate(False)

        self._var = tk.StringVar(value="Ready")
        tk.Label(
            self, textvariable=self._var,
            bg=C["nav"], fg="#8892b0",
            font=("Segoe UI", 8), anchor="w", padx=14,
        ).pack(fill="x", expand=True)

    def set(self, msg: str) -> None:
        self._var.set(msg)
