"""
CertificateApp  -- the top-level controller.
"""
import os
import platform
import threading

import tkinter as tk
from tkinter import filedialog, messagebox

from PIL import Image, ImageTk

from app.constants import C, APP_TITLE, DEFAULT_FONT_SIZE
from app.helpers import resource_path, cmyk_to_hex
from app.font_manager import load_available_fonts
from app import excel_loader, generator, project_io
from app.ui.widgets import setup_ttk_styles
from app.ui.navbar import NavBar
from app.ui.status_bar import StatusBar
from app.ui.control_panel import ControlPanel
from app.ui.field_row import FieldList
from app.ui.canvas_area import CanvasArea
from app.ui.dialogs import show_preview, pick_color_rgb, pick_color_cmyk

_DEFAULT_ALIGN = "center"


def _make_field_settings(default_font: str) -> dict:
    """Return a fresh font_settings sub-dict for one field."""
    return {
        "size":           tk.IntVar(value=DEFAULT_FONT_SIZE),
        "color":          tk.StringVar(value="#000000"),
        "font_name":      tk.StringVar(value=default_font),
        "align":          tk.StringVar(value=_DEFAULT_ALIGN),
        "opacity":        tk.IntVar(value=100),
        "shadow":         tk.BooleanVar(value=False),
        "shadow_offset":  tk.IntVar(value=4),
        "outline":        tk.BooleanVar(value=False),
        "outline_width":  tk.IntVar(value=2),
    }


class CertificateApp:

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.minsize(980, 640)
        self.root.configure(bg=C["bg"])

        self.original_image  = None
        self.template_path   = None
        self.excel_path      = None
        self.excel_data:     list = []
        self.fields:         list = []
        self.field_vars:     dict = {}
        self.font_settings:  dict = {}
        self.color_space     = tk.StringVar(value="RGB")
        self._gen_lock       = threading.Lock()

        self.available_fonts = load_available_fonts()

        setup_ttk_styles()
        self._build_ui()
        self._set_icon()

    def _build_ui(self):
        NavBar(
            self.root,
            save_cmd=self.save_project,
            load_cmd=self.load_project,
            load_template_cmd=self.load_template,
            load_excel_cmd=self.load_excel,
        )
        self._status_bar = StatusBar(self.root)
        body = tk.Frame(self.root, bg=C["bg"])
        body.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        self._panel = ControlPanel(
            body,
            preview_cmd=self.preview_certificate,
            generate_cmd=self.generate_certificates,
        )
        self._field_list  = FieldList(self._panel.fields_frame)
        self._canvas_area = CanvasArea(body)

    def _set_icon(self):
        for p in map(resource_path,
                     ("icon.ico")):
            if not os.path.exists(p): continue
            try:
                if platform.system() == "Windows" and p.endswith(".ico"):
                    self.root.iconbitmap(p)
                else:
                    photo = ImageTk.PhotoImage(Image.open(p))
                    self.root.iconphoto(True, photo)
                    self._icon_ref = photo
                return
            except Exception:
                pass

    def _status(self, msg: str, ok: bool = True) -> None:
        self._status_bar.set(msg, ok)
        self.root.update_idletasks()

    def _log(self, msg, clear=False):
        self.root.after(0, lambda m=msg, c=clear: self._panel.append_log(m, c))

    # ------------------------------------------------------------------
    def load_template(self, file_path=None):
        if not file_path:
            file_path = filedialog.askopenfilename(
                title="Select certificate template",
                filetypes=[("Images", "*.png *.jpg *.jpeg")])
        if not file_path: return
        try:
            self.template_path  = file_path
            self.original_image = Image.open(file_path).convert("RGBA")
            self._canvas_area.load_image(self.original_image)
            for f in self.fields:
                self._canvas_area.create_placeholder(f)
            self._status(f"Template: {os.path.basename(file_path)}")
        except Exception as exc:
            messagebox.showerror("Error", f"Cannot load template:\n{exc}")

    def load_excel(self, file_path=None):
        if not file_path:
            file_path = filedialog.askopenfilename(
                title="Select data file",
                filetypes=[("Excel / CSV", "*.xlsx *.csv")])
        if not file_path: return
        try:
            header, rows = excel_loader.read(file_path)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return

        self.excel_path  = file_path
        self.fields      = header
        self.excel_data  = rows

        default_font = next(iter(self.available_fonts))
        self.field_vars    = {f: tk.BooleanVar(value=True) for f in header}
        self.font_settings = {f: _make_field_settings(default_font) for f in header}

        self._canvas_area.font_settings   = self.font_settings
        self._canvas_area.available_fonts = self.available_fonts
        self._canvas_area.excel_data      = self.excel_data
        self._canvas_area.fields          = self.fields

        self._field_list.rebuild(
            self.fields, self.field_vars, self.font_settings,
            list(self.available_fonts.keys()),
            update_cb=self._on_field_update,
            color_cb=self._on_pick_color,
        )
        if self.original_image:
            for f in self.fields:
                self._canvas_area.create_placeholder(f)

        ext = os.path.splitext(file_path)[1].upper()
        self._status(f"{ext} loaded: {len(rows)} records, {len(header)} fields")

    def _on_field_update(self, field):
        self._canvas_area.update_placeholder(field)
        self._status(f"Updated: {field}")

    def _on_pick_color(self, field):
        if self.color_space.get() == "RGB":
            pick_color_rgb(self.root, field, self.font_settings)
        else:
            pick_color_cmyk(self.root, field, self.font_settings)
        self._canvas_area.update_placeholder(field)
        swatch = self.font_settings[field].get("_swatch")
        if swatch:
            col = self.font_settings[field]["color"].get()
            if col.startswith("cmyk("):
                col = cmyk_to_hex(col)
            try:
                swatch.config(bg=col)
            except Exception:
                pass

    def preview_certificate(self):
        if not self.original_image:
            messagebox.showwarning("CertWizard", "Load a template first."); return
        if not self.excel_data:
            messagebox.showwarning("CertWizard", "Load student data first."); return
        from app.image_renderer import draw_text_on_image
        img = draw_text_on_image(
            self.original_image.copy(),
            self.fields, self.field_vars, self.font_settings,
            self.available_fonts, self.excel_data[0],
            self._canvas_area.get_scaled_positions(),
        )
        show_preview(self.root, img)

    def generate_certificates(self):
        if not self.excel_data:
            messagebox.showwarning("CertWizard", "Load student data first."); return
        if not self.original_image:
            messagebox.showwarning("CertWizard", "Load a template first."); return
        if not self._gen_lock.acquire(blocking=False):
            messagebox.showwarning("CertWizard", "Generation already in progress."); return

        use_cmyk = messagebox.askyesno(
            "Color mode", "Generate in CMYK?\n\nYes = CMYK    No = RGB")
        self.color_space.set("CMYK" if use_cmyk else "RGB")
        out_dir = filedialog.askdirectory(title="Select output folder")
        if not out_dir:
            self._gen_lock.release(); return

        generator.run(
            excel_data=self.excel_data,
            fields=self.fields,
            field_vars=self.field_vars,
            font_settings=self.font_settings,
            available_fonts=self.available_fonts,
            original_image=self.original_image,
            positions=self._canvas_area.get_scaled_positions(),
            out_dir=out_dir,
            color_mode="CMYK" if use_cmyk else "RGB",
            filename_pattern=self._panel.filename_pattern.get(),
            on_progress=lambda pct: self.root.after(
                0, lambda v=pct: self._panel.set_progress(v)),
            on_log=lambda msg, clr: self.root.after(
                0, lambda m=msg, c=clr: self._panel.append_log(m, c)),
            on_done=lambda cnt, tot: self.root.after(
                0, lambda: messagebox.showinfo(
                    "CertWizard", f"{cnt} certificate(s) generated!")),
            lock=self._gen_lock,
        )

    # ------------------------------------------------------------------
    def save_project(self):
        if not self.original_image:
            messagebox.showwarning("Warning", "No template loaded."); return
        try:
            data = project_io.serialise(
                template_path=self.template_path,
                excel_path=self.excel_path,
                color_space=self.color_space.get(),
                positions=self._canvas_area.get_scaled_positions(),
                fields=self.fields,
                font_settings=self.font_settings,
                field_vars=self.field_vars,
                filename_pattern=self._panel.filename_pattern.get(),
            )
            path = filedialog.asksaveasfilename(
                defaultextension=".certwiz",
                filetypes=[("CertWizard Project", "*.certwiz")])
            if path:
                project_io.save(path, data)
                messagebox.showinfo("Saved", "Project saved.")
        except Exception as exc:
            messagebox.showerror("Error", f"Save failed:\n{exc}")

    def load_project(self):
        path = filedialog.askopenfilename(
            filetypes=[("CertWizard Project", "*.certwiz")])
        if not path: return
        try:
            data = project_io.load(path)
        except Exception as exc:
            messagebox.showerror("Error", f"Load failed:\n{exc}"); return

        self._canvas_area.clear()
        tpl = data.get("template_path", "")
        if tpl and os.path.exists(tpl):   self.load_template(tpl)
        elif tpl: messagebox.showwarning("Warning", f"Template not found:\n{tpl}")
        xl = data.get("excel_path", "")
        if xl and os.path.exists(xl):     self.load_excel(xl)
        elif xl: messagebox.showwarning("Warning", f"Excel not found:\n{xl}")

        self.color_space.set(data.get("color_space", "RGB"))
        self._panel.filename_pattern.set(data.get("filename_pattern", ""))

        fs = data.get("field_settings", {})
        for f in self.fields:
            if f not in fs: continue
            d = fs[f]
            self.font_settings[f]["size"].set(d.get("size",     DEFAULT_FONT_SIZE))
            self.font_settings[f]["color"].set(d.get("color",   "#000000"))
            self.font_settings[f]["font_name"].set(
                d.get("font_name", next(iter(self.available_fonts))))
            self.font_settings[f]["align"].set(d.get("align",   _DEFAULT_ALIGN))
            self.font_settings[f]["opacity"].set(d.get("opacity", 100))
            self.font_settings[f]["shadow"].set(d.get("shadow",  False))
            self.font_settings[f]["shadow_offset"].set(d.get("shadow_offset", 4))
            self.font_settings[f]["outline"].set(d.get("outline", False))
            self.font_settings[f]["outline_width"].set(d.get("outline_width", 2))
            self.field_vars[f].set(d.get("visible", True))

        for field, (sx, sy) in data.get("positions", {}).items():
            if field in self.fields:
                self._canvas_area.draw_placeholder(
                    field,
                    sx / self._canvas_area.scale_x,
                    sy / self._canvas_area.scale_y,
                )
        messagebox.showinfo("Loaded", "Project loaded.")

    # legacy aliases
    def _update_status(self, msg): self._status(msg)
    def update_status(self, msg):  self._status(msg)
    def update_info(self, msg, clear=False): self._log(msg, clear)
    def get_placeholder_positions(self): return self._canvas_area.get_scaled_positions()
    def update_preview(self, field):     self._on_field_update(field)
