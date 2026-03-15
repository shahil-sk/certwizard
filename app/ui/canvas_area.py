"""
Right-side canvas frame: template + draggable text placeholders.
Features:
  - Scroll (mousewheel + scrollbars)
  - Row preview switcher (← row N/total →) in a toolbar above the canvas
  - Ctrl+Scroll zoom (25% – 200%)
"""
import platform
import tkinter as tk
from tkinter import ttk

from PIL import Image, ImageTk

from app.constants import C, CANVAS_MAX_W, CANVAS_MAX_H
from app.image_renderer import render_placeholder

_LANCZOS = Image.Resampling.LANCZOS


class CanvasArea(tk.Frame):

    def __init__(self, parent):
        super().__init__(
            parent, bg=C["surface"],
            highlightthickness=1, highlightbackground=C["border"],
        )
        self.pack(side="left", fill="both", expand=True, pady=6)

        self._scale_x = self._scale_y = 1.0
        self._zoom    = 1.0
        self._pil_image = None          # original PIL image (unscaled)
        self._display_image = None
        self._ph_images: dict = {}
        self._placeholders: dict = {}
        self._drag: dict = {}

        # preview row state
        self._preview_idx = 0

        # injected by App
        self.font_settings:   dict = {}
        self.available_fonts: dict = {}
        self.excel_data:      list = []
        self.fields:          list = []

        self._build_toolbar()
        self._build_canvas()

    # ------------------------------------------------------------------
    # Toolbar (row preview switcher)
    # ------------------------------------------------------------------
    def _build_toolbar(self) -> None:
        tb = tk.Frame(self, bg=C["nav"], height=32)
        tb.pack(fill="x")
        tb.pack_propagate(False)

        # Left: zoom label
        self._zoom_label = tk.Label(
            tb, text="100%",
            font=("Segoe UI", 8), fg=C["subtext"], bg=C["nav"],
        )
        self._zoom_label.pack(side="left", padx=10)

        # Centre: row navigator
        nav = tk.Frame(tb, bg=C["nav"])
        nav.pack(side="left", expand=True)

        _btn = lambda txt, cmd: tk.Button(
            nav, text=txt, command=cmd,
            bg=C["nav"], fg=C["subtext"],
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI", 9),
            activebackground=C["surface"],
            activeforeground=C["text"],
            padx=8, pady=2,
        )
        _btn("◄", self._prev_row).pack(side="left")
        self._row_label = tk.Label(
            nav, text="row —",
            font=("Segoe UI", 8), fg=C["text"], bg=C["nav"],
            width=12,
        )
        self._row_label.pack(side="left", padx=4)
        _btn("►", self._next_row).pack(side="left")

        # Right: hint
        tk.Label(
            tb, text="Ctrl+scroll to zoom",
            font=("Segoe UI", 7), fg=C["muted"], bg=C["nav"],
        ).pack(side="right", padx=10)

    # ------------------------------------------------------------------
    # Canvas
    # ------------------------------------------------------------------
    def _build_canvas(self) -> None:
        self._h_scroll = ttk.Scrollbar(
            self, orient="horizontal",
            style="Dark.Horizontal.TScrollbar")
        self._v_scroll = ttk.Scrollbar(
            self, orient="vertical",
            style="Dark.Vertical.TScrollbar")

        self._canvas = tk.Canvas(
            self, bg=C["canvas_bg"], highlightthickness=0,
            xscrollcommand=self._h_scroll.set,
            yscrollcommand=self._v_scroll.set,
        )
        self._h_scroll.config(command=self._canvas.xview)
        self._v_scroll.config(command=self._canvas.yview)

        self._h_scroll.pack(side="bottom", fill="x")
        self._v_scroll.pack(side="right",  fill="y")
        self._canvas.pack(fill="both", expand=True, padx=1, pady=1)

        self._canvas.bind("<Enter>", self._bind_scroll)
        self._canvas.bind("<Leave>", self._unbind_scroll)

    # ------------------------------------------------------------------
    # Properties
    # ------------------------------------------------------------------
    @property
    def scale_x(self): return self._scale_x

    @property
    def scale_y(self): return self._scale_y

    # ------------------------------------------------------------------
    # Image loading
    # ------------------------------------------------------------------
    def load_image(self, pil_image) -> None:
        self._pil_image = pil_image
        self._preview_idx = 0
        self._redraw_image()

    def _redraw_image(self) -> None:
        if not self._pil_image:
            return
        ow, oh = self._pil_image.size
        base_ratio = min(CANVAS_MAX_W / ow, CANVAS_MAX_H / oh)
        ratio = base_ratio * self._zoom
        nw, nh = max(int(ow * ratio), 1), max(int(oh * ratio), 1)
        self._scale_x = ow / nw * self._zoom
        self._scale_y = oh / nh * self._zoom
        # Keep the "logical" scale for position mapping:
        # positions are stored relative to original image pixels.
        # scale_x = original_px / canvas_px
        self._scale_x = ow / (nw / self._zoom * self._zoom)
        # simplify: scale is always original/canvas at base zoom
        base_nw = max(int(ow * base_ratio), 1)
        base_nh = max(int(oh * base_ratio), 1)
        self._scale_x = ow / base_nw
        self._scale_y = oh / base_nh

        self._display_image = ImageTk.PhotoImage(
            self._pil_image.resize((nw, nh), _LANCZOS))
        self._canvas.config(
            width=min(nw, CANVAS_MAX_W),
            height=min(nh, CANVAS_MAX_H),
            scrollregion=(0, 0, nw, nh),
        )
        self._canvas.delete("all")
        self._canvas.create_image(0, 0, image=self._display_image, anchor="nw")
        self._zoom_label.config(text=f"{int(self._zoom * 100)}%")

        for field, data in list(self._placeholders.items()):
            self.draw_placeholder(field, data["x"], data["y"])

    # ------------------------------------------------------------------
    # Row preview switcher
    # ------------------------------------------------------------------
    def _update_row_label(self) -> None:
        total = len(self.excel_data)
        if total:
            self._row_label.config(
                text=f"row {self._preview_idx + 1} / {total}")
        else:
            self._row_label.config(text="row —")

    def _prev_row(self) -> None:
        if not self.excel_data:
            return
        self._preview_idx = (self._preview_idx - 1) % len(self.excel_data)
        self._update_row_label()
        self._refresh_all_placeholders()

    def _next_row(self) -> None:
        if not self.excel_data:
            return
        self._preview_idx = (self._preview_idx + 1) % len(self.excel_data)
        self._update_row_label()
        self._refresh_all_placeholders()

    def _refresh_all_placeholders(self) -> None:
        """Redraw every placeholder using the current preview row."""
        for field, data in list(self._placeholders.items()):
            self.draw_placeholder(field, data["x"], data["y"])

    def _current_row(self) -> dict:
        if self.excel_data:
            return self.excel_data[self._preview_idx]
        return {}

    # ------------------------------------------------------------------
    # Placeholder management
    # ------------------------------------------------------------------
    def draw_placeholder(self, field: str, x: float, y: float) -> None:
        if field in self._placeholders:
            self._canvas.delete(self._placeholders[field]["item"])

        # Build a temporary font_settings override that uses the current row
        # value as the preview text (render_placeholder already does this
        # via excel_data[0] — we temporarily swap excel_data).
        saved = self.excel_data
        row   = self._current_row()
        self.excel_data = [row] if row else saved

        ph_img = render_placeholder(
            field, self.font_settings, self.available_fonts,
            self.excel_data, self._scale_x, self._scale_y,
        )
        self.excel_data = saved

        photo = ImageTk.PhotoImage(ph_img)
        self._ph_images[field] = photo

        item = self._canvas.create_image(x, y, image=photo, anchor="center")
        self._canvas.tag_bind(item, "<Button-1>",
                              lambda e, i=item: self._drag_start(e, i))
        self._canvas.tag_bind(item, "<B1-Motion>",
                              lambda e, i=item: self._drag_move(e, i))
        self._placeholders[field] = {"item": item, "x": x, "y": y}
        self._update_row_label()

    def create_placeholder(self, field: str, x=None, y=None) -> None:
        if field not in self.fields:
            return
        if x is None or y is None:
            cw  = self._canvas.winfo_width() or 800
            idx = self.fields.index(field)
            x, y = cw // 2, 50 + idx * 60
        self.draw_placeholder(field, x, y)

    def update_placeholder(self, field: str) -> None:
        if field in self._placeholders:
            p = self._placeholders[field]
            self.draw_placeholder(field, p["x"], p["y"])

    def get_scaled_positions(self) -> dict:
        return {
            f: (d["x"] * self._scale_x, d["y"] * self._scale_y)
            for f, d in self._placeholders.items()
        }

    def clear(self) -> None:
        self._canvas.delete("all")
        self._placeholders.clear()
        self._ph_images.clear()
        self._preview_idx = 0
        self._update_row_label()

    # ------------------------------------------------------------------
    # Scroll + Zoom
    # ------------------------------------------------------------------
    def _bind_scroll(self, _event=None) -> None:
        system = platform.system()
        if system == "Windows":
            self._canvas.bind_all("<MouseWheel>",          self._on_mousewheel)
            self._canvas.bind_all("<Shift-MouseWheel>",    self._scroll_x)
        elif system == "Darwin":
            self._canvas.bind_all("<MouseWheel>",          self._on_mousewheel_mac)
            self._canvas.bind_all("<Shift-MouseWheel>",    self._scroll_x_mac)
        else:
            self._canvas.bind_all("<Button-4>",            self._scroll_up)
            self._canvas.bind_all("<Button-5>",            self._scroll_down)
            self._canvas.bind_all("<Shift-Button-4>",      self._scroll_left)
            self._canvas.bind_all("<Shift-Button-5>",      self._scroll_right)
            self._canvas.bind_all("<Control-Button-4>",    self._zoom_in)
            self._canvas.bind_all("<Control-Button-5>",    self._zoom_out)

    def _unbind_scroll(self, _event=None) -> None:
        for seq in ("<MouseWheel>", "<Shift-MouseWheel>",
                    "<Button-4>", "<Button-5>",
                    "<Shift-Button-4>", "<Shift-Button-5>",
                    "<Control-Button-4>", "<Control-Button-5>"):
            try:
                self._canvas.unbind_all(seq)
            except Exception:
                pass

    def _on_mousewheel(self, event):
        if event.state & 0x0004:   # Ctrl held
            if event.delta > 0:
                self._zoom_in()
            else:
                self._zoom_out()
        else:
            self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_mac(self, event):
        if event.state & 0x0004:
            if event.delta > 0:
                self._zoom_in()
            else:
                self._zoom_out()
        else:
            self._canvas.yview_scroll(int(-1 * event.delta), "units")

    def _scroll_x(self, event):
        self._canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def _scroll_x_mac(self, event):
        self._canvas.xview_scroll(int(-1 * event.delta), "units")

    def _scroll_up(self,    _e): self._canvas.yview_scroll(-1, "units")
    def _scroll_down(self,  _e): self._canvas.yview_scroll( 1, "units")
    def _scroll_left(self,  _e): self._canvas.xview_scroll(-1, "units")
    def _scroll_right(self, _e): self._canvas.xview_scroll( 1, "units")

    def _zoom_in(self, _e=None):
        if self._zoom < 2.0:
            self._zoom = round(min(self._zoom + 0.1, 2.0), 1)
            self._redraw_image()

    def _zoom_out(self, _e=None):
        if self._zoom > 0.3:
            self._zoom = round(max(self._zoom - 0.1, 0.3), 1)
            self._redraw_image()

    # ------------------------------------------------------------------
    # Drag
    # ------------------------------------------------------------------
    def _drag_start(self, event, item) -> None:
        self._drag = {"item": item, "x": event.x, "y": event.y}

    def _drag_move(self, event, item) -> None:
        dx = event.x - self._drag["x"]
        dy = event.y - self._drag["y"]
        self._canvas.move(item, dx, dy)
        self._drag["x"] = event.x
        self._drag["y"] = event.y
        for f, d in self._placeholders.items():
            if d["item"] == item:
                d["x"] += dx
                d["y"] += dy
                break
