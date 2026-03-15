"""
CertificateApp  -- the top-level controller.

Wires together all UI components and domain modules.
Contains no raw widget creation beyond what belongs here.
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


class CertificateApp:

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.minsize(980, 640)
        self.root.configure(bg=C["bg"])

        # domain state
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

    # ------------------------------------------------------------------
    # UI assembly
    # ------------------------------------------------------------------
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
        self._field_list = FieldList(self._panel.fields_frame)

        self._canvas_area = CanvasArea(body)

    # ------------------------------------------------------------------
    # Icon
    # ------------------------------------------------------------------
    def _set_icon(self):
        for p in map(resource_path,
                     ("certgen.ico", "certgen.png", "icon.ico", "icon.png")):
            if not os.path.exists(p):
                continue
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

    # ------------------------------------------------------------------
    # Status / log helpers
    # ------------------------------------------------------------------
    def _status(self, msg: str) -> None:
        self._status_bar.set(msg)
        self.root.update_idletasks()

    def _log(self, msg: str, clear: bool = False) -> None:
        self.root.after(
            0, lambda m=msg, c=clear: self._panel.append_log(m, c))

    # ------------------------------------------------------------------
    # Template loading
    # ------------------------------------------------------------------
    def load_template(self, file_path=None):
        if not file_path:
            file_path = filedialog.askopenfilename(
                title="Select certificate template",
                filetypes=[("Images", "*.png *.jpg *.jpeg")])
        if not file_path:
            return
        try:
            self.template_path  = file_path
            self.original_image = Image.open(file_path).convert("RGBA")
            self._canvas_area.load_image(self.original_image)
            for f in self.fields:
                self._canvas_area.create_placeholder(f)
            self._status(f"Template: {os.path.basename(file_path)}")
        except Exception as exc:
            messagebox.showerror("Error", f"Cannot load template:\n{exc}")

    # ------------------------------------------------------------------
    # Excel loading
    # ------------------------------------------------------------------
    def load_excel(self, file_path=None):
        if not file_path:
            file_path = filedialog.askopenfilename(
                title="Select Excel file",
                filetypes=[("Excel", "*.xlsx")])
        if not file_path:
            return
        try:
            header, rows = excel_loader.read(file_path)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return

        self.excel_path  = file_path
        self.fields      = header
        self.excel_data  = rows

        default_font = next(iter(self.available_fonts))
        self.field_vars   = {f: tk.BooleanVar(value=True) for f in header}
        self.font_settings = {
            f: {
                "size":      tk.IntVar(value=DEFAULT_FONT_SIZE),
                "color":     tk.StringVar(value="#000000"),
                "font_name": tk.StringVar(value=default_font),
            } for f in header
        }

        # Push state into canvas so it can render placeholders
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

        self._status(f"Excel loaded: {len(rows)} records, {len(header)} fields")

    # ------------------------------------------------------------------
    # Field callbacks
    # ------------------------------------------------------------------
    def _on_field_update(self, field: str) -> None:
        self._canvas_area.update_placeholder(field)
        self._status(f"Updated field: {field}")

    def _on_pick_color(self, field: str) -> None:
        if self.color_space.get() == "RGB":
            pick_color_rgb(self.root, field, self.font_settings)
        else:
            pick_color_cmyk(self.root, field, self.font_settings)
        self._canvas_area.update_placeholder(field)
        # Refresh the colour swatch in the field row
        swatch = self.font_settings[field].get("_swatch")
        if swatch:
            col = self.font_settings[field]["color"].get()
            if col.startswith("cmyk("):
                col = cmyk_to_hex(col)
            try:
                swatch.config(bg=col)
            except Exception:
                pass

    # ------------------------------------------------------------------
    # Preview
    # ------------------------------------------------------------------
    def preview_certificate(self):
        if not self.original_image:
            messagebox.showwarning("CertWizard", "Load a template first.")
            return
        if not self.excel_data:
            messagebox.showwarning("CertWizard", "Load student data first.")
            return

        from app.image_renderer import draw_text_on_image
        img = draw_text_on_image(
            self.original_image.copy(),
            self.fields, self.field_vars, self.font_settings,
            self.available_fonts, self.excel_data[0],
            self._canvas_area.get_scaled_positions(),
        )
        show_preview(self.root, img)

    # ------------------------------------------------------------------
    # Generation
    # ------------------------------------------------------------------
    def generate_certificates(self):
        if not self.excel_data:
            messagebox.showwarning("CertWizard", "Load student data first.")
            return
        if not self.original_image:
            messagebox.showwarning("CertWizard", "Load a template first.")
            return
        if not self._gen_lock.acquire(blocking=False):
            messagebox.showwarning("CertWizard", "Generation already in progress.")
            return

        use_cmyk = messagebox.askyesno(
            "Color mode",
            "Generate in CMYK color space?\n\nYes  =  CMYK    No  =  RGB")
        self.color_space.set("CMYK" if use_cmyk else "RGB")

        out_dir = filedialog.askdirectory(title="Select output folder")
        if not out_dir:
            self._gen_lock.release()
            return

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
            on_progress=lambda pct: self.root.after(
                0, lambda v=pct: self._panel.set_progress(v)),
            on_log=lambda msg, clr: self.root.after(
                0, lambda m=msg, c=clr: self._panel.append_log(m, c)),
            on_done=lambda cnt, tot: self.root.after(
                0, lambda: messagebox.showinfo(
                    "CertWizard",
                    f"{cnt} certificate(s) generated successfully.")),
            lock=self._gen_lock,
        )

    # ------------------------------------------------------------------
    # Project save / load
    # ------------------------------------------------------------------
    def save_project(self):
        if not self.original_image:
            messagebox.showwarning("Warning", "No template loaded.")
            return
        try:
            data = project_io.serialise(
                template_path=self.template_path,
                excel_path=self.excel_path,
                color_space=self.color_space.get(),
                positions=self._canvas_area.get_scaled_positions(),
                fields=self.fields,
                font_settings=self.font_settings,
                field_vars=self.field_vars,
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
        if not path:
            return
        try:
            data = project_io.load(path)
        except Exception as exc:
            messagebox.showerror("Error", f"Load failed:\n{exc}")
            return

        self._canvas_area.clear()

        tpl = data.get("template_path", "")
        if tpl and os.path.exists(tpl):
            self.load_template(tpl)
        elif tpl:
            messagebox.showwarning("Warning", f"Template not found:\n{tpl}")

        xl = data.get("excel_path", "")
        if xl and os.path.exists(xl):
            self.load_excel(xl)
        elif xl:
            messagebox.showwarning("Warning", f"Excel file not found:\n{xl}")

        self.color_space.set(data.get("color_space", "RGB"))

        fs = data.get("field_settings", {})
        for f in self.fields:
            if f in fs:
                self.font_settings[f]["size"].set(fs[f].get("size", DEFAULT_FONT_SIZE))
                self.font_settings[f]["color"].set(fs[f].get("color", "#000000"))
                self.font_settings[f]["font_name"].set(
                    fs[f].get("font_name", next(iter(self.available_fonts))))
                self.field_vars[f].set(fs[f].get("visible", True))

        for field, (sx, sy) in data.get("positions", {}).items():
            if field in self.fields:
                self._canvas_area.draw_placeholder(
                    field,
                    sx / self._canvas_area.scale_x,
                    sy / self._canvas_area.scale_y,
                )

        messagebox.showinfo("Loaded", "Project loaded.")

    # ------------------------------------------------------------------
    # Legacy aliases (keeps old .certwiz / external references working)
    # ------------------------------------------------------------------
    def _update_status(self, msg): self._status(msg)
    def update_status(self, msg):  self._status(msg)
    def update_info(self, msg, clear=False): self._log(msg, clear)
    def get_placeholder_positions(self): return self._canvas_area.get_scaled_positions()
    def update_preview(self, field):     self._on_field_update(field)
