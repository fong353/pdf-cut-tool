import tkinter as tk

import pymupdf
from PIL import Image, ImageTk

from pdf_ops import PT_PER_MM

OVERLAY_TAG = 'overlay'
LINE_COLOR = '#e60000'


class PreviewPane(tk.Frame):
    """PDF 首页位图预览 + 矢量叠加层（切割线/外框）。"""

    def __init__(self, parent, width=420, height=580):
        super().__init__(parent)
        self.canvas = tk.Canvas(
            self, width=width, height=height,
            bg='#ededed', highlightthickness=1, highlightbackground='#b5b5b5',
        )
        self.canvas.pack(fill='both', expand=True)

        self._pdf_path = None
        self._photo = None
        self.page_w_mm = 0.0
        self.page_h_mm = 0.0
        self._img_x0 = 0
        self._img_y0 = 0
        self._img_w = 0
        self._img_h = 0
        self._mm_to_px = 1.0

        # 由 App 赋值：位图重绘完成后回调，让 App 重画 overlay
        self.on_render = None

        self.canvas.bind('<Configure>', self._on_resize)
        self._resize_job = None

    def load_pdf(self, path):
        self._pdf_path = path
        self._render()

    def clear(self):
        self._pdf_path = None
        self._photo = None
        self.canvas.delete('all')
        self.page_w_mm = 0.0
        self.page_h_mm = 0.0

    def mm_to_canvas_x(self, x_mm):
        return self._img_x0 + x_mm * self._mm_to_px

    def mm_to_canvas_y(self, y_mm):
        return self._img_y0 + y_mm * self._mm_to_px

    def clear_overlay(self):
        self.canvas.delete(OVERLAY_TAG)

    def set_overlay(self, cuts_mm, inset_mm=None, add_border=False):
        """清空叠加层，按 mm 坐标重画竖切线与外框。"""
        self.canvas.delete(OVERLAY_TAG)
        if not self.page_w_mm:
            return

        if add_border and inset_mm is not None:
            bx0 = self.mm_to_canvas_x(inset_mm)
            by0 = self.mm_to_canvas_y(inset_mm)
            bx1 = self.mm_to_canvas_x(self.page_w_mm - inset_mm)
            by1 = self.mm_to_canvas_y(self.page_h_mm - inset_mm)
            self.canvas.create_rectangle(
                bx0, by0, bx1, by1,
                outline=LINE_COLOR, width=1, tags=OVERLAY_TAG,
            )

        for x_mm in cuts_mm:
            px = self.mm_to_canvas_x(x_mm)
            self.canvas.create_line(
                px, self._img_y0, px, self._img_y0 + self._img_h,
                fill=LINE_COLOR, width=1, tags=OVERLAY_TAG,
            )

    def set_overlay_crop(self, mark_len_mm, gap_mm):
        """清空叠加层，在 MediaBox 4 角画 L 形裁切标记（向外）。"""
        self.canvas.delete(OVERLAY_TAG)
        if not self.page_w_mm:
            return
        color = '#000000'
        gap_px = gap_mm * self._mm_to_px
        len_px = mark_len_mm * self._mm_to_px
        # (mm 坐标, 方向 sx/sy：sx 控制 x 延伸方向, sy 控制 PDF 的 y 方向)
        for (x_mm, y_mm, sx, sy) in [
            (0, 0, -1, -1),
            (self.page_w_mm, 0, +1, -1),
            (0, self.page_h_mm, -1, +1),
            (self.page_w_mm, self.page_h_mm, +1, +1),
        ]:
            cx = self.mm_to_canvas_x(x_mm)
            cy = self.mm_to_canvas_y(y_mm)
            # 水平臂
            self.canvas.create_line(
                cx + sx * gap_px, cy,
                cx + sx * (gap_px + len_px), cy,
                fill=color, width=1, tags=OVERLAY_TAG,
            )
            # 垂直臂：canvas y 向下增长，PDF y 向上增长，sy 反向映射
            self.canvas.create_line(
                cx, cy - sy * gap_px,
                cx, cy - sy * (gap_px + len_px),
                fill=color, width=1, tags=OVERLAY_TAG,
            )

    def set_overlay_circle(self, cx_mm, cy_mm, r_mm):
        """清空叠加层，画一个圆。"""
        self.canvas.delete(OVERLAY_TAG)
        if not self.page_w_mm or r_mm <= 0:
            return
        cx = self.mm_to_canvas_x(cx_mm)
        cy = self.mm_to_canvas_y(cy_mm)
        r_px = r_mm * self._mm_to_px
        self.canvas.create_oval(
            cx - r_px, cy - r_px, cx + r_px, cy + r_px,
            outline=LINE_COLOR, width=1, tags=OVERLAY_TAG,
        )

    def _on_resize(self, _event):
        if not self._pdf_path:
            return
        # 防抖，resize 频繁触发时只重绘最后一次
        if self._resize_job is not None:
            self.after_cancel(self._resize_job)
        self._resize_job = self.after(80, self._render)

    def _render(self):
        self._resize_job = None
        if not self._pdf_path:
            return
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw < 10 or ch < 10:
            return

        self.canvas.delete('all')
        doc = pymupdf.open(self._pdf_path)
        try:
            page = doc[0]
            rect = page.rect
            pw_pt, ph_pt = rect.width, rect.height
            self.page_w_mm = pw_pt / PT_PER_MM
            self.page_h_mm = ph_pt / PT_PER_MM

            scale = min(cw / pw_pt, ch / ph_pt) * 0.95
            pix = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
            img = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
            self._photo = ImageTk.PhotoImage(img)
            self._img_w = pix.width
            self._img_h = pix.height
            self._img_x0 = (cw - self._img_w) // 2
            self._img_y0 = (ch - self._img_h) // 2
            self._mm_to_px = self._img_w / self.page_w_mm

            self.canvas.create_image(
                self._img_x0, self._img_y0, anchor='nw', image=self._photo,
            )
        finally:
            doc.close()

        if self.on_render:
            self.on_render()
