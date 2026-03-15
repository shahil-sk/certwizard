"""
Top navigation bar widget.
"""
import tkinter as tk

from app.constants import C, APP_TITLE, APP_VERSION
from app.ui.widgets import flat_button, label


class NavBar(tk.Frame):
    """
    A fixed-height dark navbar.
    Exposes save_cmd / load_cmd / load_template_cmd / load_excel_cmd
    so the caller (App) wires up its own methods.
    """

    def __init__(
        self,
        parent,
        save_cmd,
        load_cmd,
        load_template_cmd,
        load_excel_cmd,
    ):
        super().__init__(parent, bg=C["nav"], height=48)
        self.pack(fill="x")
        self.pack_propagate(False)
        self._build(save_cmd, load_cmd, load_template_cmd, load_excel_cmd)

    def _build(self, save_cmd, load_cmd, load_template_cmd, load_excel_cmd):
        label(self, APP_TITLE, font_size=13, bold=True,
              color="#ffffff", bg=C["nav"], padx=18).pack(side="left")
        label(self, f"v{APP_VERSION}", font_size=8,
              color="#8892b0", bg=C["nav"]).pack(side="left", pady=(2, 0))

        right = tk.Frame(self, bg=C["nav"])
        right.pack(side="right", padx=12)

        proj_btn = tk.Menubutton(
            right, text="Project", bg=C["nav"], fg="#c9d1d9",
            relief="flat", font=("Segoe UI", 9), padx=10,
            cursor="hand2", activebackground="#2d3250",
            activeforeground="white", bd=0,
        )
        proj_btn.menu = tk.Menu(
            proj_btn, tearoff=0,
            bg=C["surface"], fg=C["text"],
            activebackground=C["accent"], activeforeground="white",
            font=("Segoe UI", 9), bd=0, relief="flat",
        )
        proj_btn["menu"] = proj_btn.menu
        proj_btn.menu.add_command(label="Save Project", command=save_cmd)
        proj_btn.menu.add_command(label="Load Project", command=load_cmd)
        proj_btn.pack(side="left", padx=4)

        for text, cmd in (
            ("Load Template", load_template_cmd),
            ("Load Excel",    load_excel_cmd),
        ):
            flat_button(right, text, cmd, C["accent"], C["accent2"],
                        padx=12, pady=6).pack(side="left", padx=4)
