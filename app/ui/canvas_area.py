"""
Right-side canvas frame: displays the template and draggable text placeholders.
"""
import tkinter as tk

from PIL import ImageTk

from app.constants import C, CANVAS_MAX_W, CANVAS_MAX_H
from app.image_renderer import render_placeholder


class CanvasArea(tk.Frame):
    """
    Owns the tk.Canvas and all placeholder management.
    Exposes:
      load_image(pil_image)        -- draw/redraw background template
      draw_placeholder(field, x, y)
      create_placeholder(field, x=None, y=None)
      update_placeholder(field)
      get_scaled_positions() -> dict
      scale_x / scale_y           -- read-only properties
    """

    def __init__(self, parent):
        super().__init__(
            parent, bg=C["surface"],
            highlightthickness=1, highlightbackground=C["border"],
        )
        self.pack(side="left", fill="both", expand=True, pady=6)

        self._canvas = tk.Canvas(self, bg=C["bg"], highlightthickness=0)
        self._canvas.pack(fill="both", expand=True, padx=1, pady=1)

        self._scale_x = self._scale_y = 1.0
        self._display_image = None   # keep ImageTk ref alive
        self._ph_images: dict = {}   # field -> ImageTk.PhotoImage
        self._placeholders: dict = {}  # field -> {item, x, y}
        self._drag: dict = {}

        # state injected by App after Excel load
        self.font_settings:  dict = {}
        self.available_fonts: dict = {}
        self.excel_data:      list = []
        self.fields:          list = []

    # ------------------------------------------------------------------
    @property
    def scale_x(self): return self._scale_x

    @property
    def scale_y(self): return self._scale_y

    # ------------------------------------------------------------------
    def load_image(self, pil_image) -> None:
        """Scale image to fit canvas area and redraw it."""
        ow, oh = pil_image.size
        ratio  = min(CANVAS_MAX_W / ow, CANVAS_MAX_H / oh)
        nw, nh = int(ow * ratio), int(oh * ratio)
        self._scale_x, self._scale_y = ow / nw, oh / nh

        self._display_image = ImageTk.PhotoImage(
            pil_image.resize((nw, nh), pil_image.LANCZOS))
        self._canvas.config(width=nw, height=nh)
        self._canvas.delete("all")
        self._canvas.create_image(0, 0, image=self._display_image, anchor="nw")

        for field, data in list(self._placeholders.items()):
            self.draw_placeholder(field, data["x"], data["y"])

    def draw_placeholder(self, field: str, x: float, y: float) -> None:
        if field in self._placeholders:
            self._canvas.delete(self._placeholders[field]["item"])

        ph_img = render_placeholder(
            field, self.font_settings, self.available_fonts,
            self.excel_data, self._scale_x, self._scale_y,
        )
        photo = ImageTk.PhotoImage(ph_img)
        self._ph_images[field] = photo

        item = self._canvas.create_image(x, y, image=photo, anchor="center")
        self._canvas.tag_bind(item, "<Button-1>",
                              lambda e, i=item: self._drag_start(e, i))
        self._canvas.tag_bind(item, "<B1-Motion>",
                              lambda e, i=item: self._drag_move(e, i))
        self._placeholders[field] = {"item": item, "x": x, "y": y}

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
