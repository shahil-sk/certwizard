"""
All Toplevel dialog windows.
"""
import tkinter as tk
from tkinter import colorchooser

from PIL import Image, ImageTk

from app.constants import C
from app.helpers import cmyk_to_hex
from app.ui.widgets import flat_button, label


def show_preview(
    parent: tk.Tk,
    img: Image.Image,
) -> None:
    """Open a resizable preview window for the first certificate."""
    win = tk.Toplevel(parent)
    win.title("Preview")
    win.transient(parent)
    win.grab_set()
    win.resizable(True, True)
    win.minsize(400, 300)
    win.configure(bg=C["bg"])

    pw = max(parent.winfo_width() - 100, 600)
    ph = max(parent.winfo_height() - 100, 400)
    win.geometry(f"{pw}x{ph}")

    iw, ih  = img.size
    ratio   = min(pw / iw, ph / ih)
    photo   = ImageTk.PhotoImage(
        img.resize((int(iw * ratio), int(ih * ratio)), Image.LANCZOS))

    lbl = tk.Label(win, image=photo, bg=C["bg"], bd=0)
    lbl.image = photo
    lbl.pack(fill="both", expand=True, padx=12, pady=12)

    flat_button(win, "Close", win.destroy,
                C["accent"], C["accent2"],
                font_size=9, bold=True,
                padx=24, pady=6).pack(pady=(0, 12))

    win.bind("<Escape>", lambda e: win.destroy())
    win.lift()
    win.focus_set()


def pick_color_rgb(
    parent: tk.Tk,
    field: str,
    font_settings: dict,
) -> None:
    """Open the system RGB colour picker and apply the result."""
    cur  = font_settings[field]["color"].get()
    init = cur if cur.startswith("#") else "#000000"
    chosen = colorchooser.askcolor(
        title=f"Pick color for {field}", initialcolor=init, parent=parent)
    if chosen[1]:
        font_settings[field]["color"].set(chosen[1])


def pick_color_cmyk(
    parent: tk.Tk,
    field: str,
    font_settings: dict,
) -> None:
    """Open a CMYK slider dialog and apply the result."""
    cur = font_settings[field]["color"].get()
    try:
        vals = list(map(float, cur[5:-1].split(",")))
    except Exception:
        vals = [0.0, 0.0, 0.0, 0.0]

    win = tk.Toplevel(parent)
    win.title(f"CMYK color for {field}")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()
    win.configure(bg=C["surface"])

    cvars = {ch: tk.DoubleVar(value=v)
             for ch, v in zip("CMYK", vals)}

    preview = tk.Label(win, width=22, height=4,
                       bg=C["surface"], relief="flat", bd=0)
    preview.grid(row=4, column=0, columnspan=2, padx=16, pady=10)

    def refresh(*_):
        cmyk = "cmyk({:.2f},{:.2f},{:.2f},{:.2f})".format(
            cvars["C"].get(), cvars["M"].get(),
            cvars["Y"].get(), cvars["K"].get())
        font_settings[field]["color"].set(cmyk)
        preview.config(bg=cmyk_to_hex(cmyk))

    for row, (lbl_text, ch) in enumerate(
            zip(("Cyan", "Magenta", "Yellow", "Black"), "CMYK")):
        tk.Label(
            win, text=lbl_text, font=("Segoe UI", 9),
            bg=C["surface"], fg=C["text"], anchor="e",
        ).grid(row=row, column=0, padx=(16, 8), pady=5, sticky="e")
        tk.Scale(
            win, from_=0, to=1, resolution=0.01,
            variable=cvars[ch], command=refresh,
            orient="horizontal", length=200,
            bg=C["surface"], fg=C["text"],
            troughcolor=C["border"], highlightthickness=0, relief="flat",
        ).grid(row=row, column=1, padx=(0, 16))

    refresh()
    flat_button(win, "Apply", win.destroy,
                C["accent"], C["accent2"],
                font_size=9, padx=24, pady=5).grid(
        row=5, column=0, columnspan=2, pady=(4, 14))
    parent.wait_window(win)
