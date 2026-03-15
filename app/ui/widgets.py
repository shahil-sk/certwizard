"""
Reusable tkinter widget factory functions.
Import C from constants before calling these.
"""
import tkinter as tk
from tkinter import ttk

from app.constants import C


def flat_button(
    parent,
    text: str,
    command,
    bg: str,
    active_bg: str,
    font_size: int = 9,
    bold: bool = False,
    **kw,
) -> tk.Button:
    weight = "bold" if bold else "normal"
    return tk.Button(
        parent, text=text, command=command,
        bg=bg, fg=C["white"], relief="flat", cursor="hand2",
        font=("Segoe UI", font_size, weight),
        activebackground=active_bg, activeforeground=C["white"],
        bd=0, highlightthickness=0, **kw,
    )


def label(
    parent,
    text: str,
    font_size: int = 9,
    bold: bool = False,
    color: str | None = None,
    bg: str | None = None,
    **kw,
) -> tk.Label:
    weight = "bold" if bold else "normal"
    return tk.Label(
        parent, text=text,
        font=("Segoe UI", font_size, weight),
        fg=color or C["text"],
        bg=bg or C["surface"],
        **kw,
    )


def separator(parent, bg: str | None = None, padx: int = 14) -> tk.Frame:
    f = tk.Frame(parent, bg=bg or C["border"], height=1)
    f.pack(fill="x", padx=padx)
    return f


def setup_ttk_styles() -> None:
    s = ttk.Style()
    s.theme_use("clam")
    s.configure(
        "Thin.Horizontal.TProgressbar",
        troughcolor=C["border"], background=C["accent"],
        thickness=6, borderwidth=0,
    )
    s.configure(
        "Flat.TCombobox",
        fieldbackground=C["surface"], background=C["surface"],
        foreground=C["text"], borderwidth=1, relief="flat",
    )
    s.map("Flat.TCombobox", fieldbackground=[("readonly", C["surface"])])
    s.configure(
        "Flat.Vertical.TScrollbar",
        background=C["border"], troughcolor=C["surface"],
        borderwidth=0, arrowsize=12,
    )
