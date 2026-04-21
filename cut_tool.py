import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font

import pikepdf

from pdf_ops import (
    PT_PER_MM, add_vertical_cuts, add_circle_cut,
    is_image, image_to_pdf, read_image_size_mm,
)
from preview import PreviewPane


class App:
    def __init__(self, root):
        self.root = root
        self.root.title('PDF CUT 切割线工具')
        self.root.minsize(900, 680)
        self.root.geometry('1000x720')
        self.src_path = None
        self.pdf_path = None
        self._temp_pdf = None
        self.page_w_mm = 0.0
        self.page_h_mm = 0.0
        self.active_tool = 'vertical'
        self._init_style()
        self._build()
        self.root.protocol('WM_DELETE_WINDOW', self._on_close)

    def _init_style(self):
        BG = '#ededed'
        BG_INPUT = '#ffffff'
        FG = '#1a1a1a'
        FG_MUTED = '#666'
        BORDER = '#c2c2c2'
        ACCENT = '#0a84ff'

        self.BG = BG
        self.BG_INPUT = BG_INPUT
        self.FG = FG
        self.FG_MUTED = FG_MUTED

        self.root.configure(bg=BG)

        style = ttk.Style()
        style.theme_use('clam')

        default_font = font.nametofont('TkDefaultFont')
        fam = default_font.actual('family')
        sz = default_font.actual('size')
        big_font = (fam, sz + 2, 'bold')
        bold_font = (fam, sz, 'bold')
        tab_font = (fam, sz + 1)

        style.configure('.', background=BG, foreground=FG, bordercolor=BORDER)
        style.configure('TFrame', background=BG)
        style.configure('TLabel', background=BG, foreground=FG)
        style.configure('TSeparator', background=BORDER)

        style.configure('TLabelframe', background=BG, bordercolor=BORDER, padding=10)
        style.configure('TLabelframe.Label', background=BG, foreground=FG, font=bold_font)

        style.configure('TRadiobutton', background=BG, foreground=FG)
        style.map('TRadiobutton',
                  background=[('active', BG)],
                  indicatorcolor=[('selected', ACCENT), ('!selected', BG_INPUT)])

        style.configure('TEntry', fieldbackground=BG_INPUT, foreground=FG,
                        bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER)
        style.map('TEntry', bordercolor=[('focus', ACCENT)])

        style.configure('TButton', background='#dcdcdc', foreground=FG,
                        bordercolor=BORDER, focusthickness=0, padding=(10, 6))
        style.map('TButton',
                  background=[('active', '#cfcfcf'), ('pressed', '#c2c2c2')])

        style.configure('Primary.TButton', font=big_font, padding=(14, 10),
                        background=ACCENT, foreground='#ffffff', bordercolor=ACCENT)
        style.map('Primary.TButton',
                  background=[('active', '#0066cc'), ('pressed', '#0052a8')])

        style.configure('Status.TLabel', background=BG, foreground=FG_MUTED)

        # Notebook 样式
        style.configure('TNotebook', background=BG, borderwidth=0, tabmargins=(0, 0, 0, 0))
        style.configure('TNotebook.Tab', background='#d6d6d6', foreground=FG,
                        padding=(18, 9), font=tab_font, bordercolor=BORDER)
        style.map('TNotebook.Tab',
                  background=[('selected', BG), ('active', '#e3e3e3')],
                  foreground=[('selected', FG)])

    def _build(self):
        # ── 顶部状态条 ─────────────────────────────
        top = ttk.Frame(self.root, padding=(12, 10, 12, 6))
        top.pack(fill='x')
        ttk.Button(top, text='选择 PDF / 图片', command=self._open).pack(side='left')
        self.lbl_file = ttk.Label(top, text='未选择文件', foreground='#888')
        self.lbl_file.pack(side='left', padx=10)
        self.lbl_size = ttk.Label(top, text='', style='Status.TLabel')
        self.lbl_size.pack(side='right')

        ttk.Separator(self.root).pack(fill='x')

        # ── 主体：左标签页 / 右预览 ─────────────────
        body = ttk.Frame(self.root)
        body.pack(fill='both', expand=True)

        left = ttk.Frame(body, padding=(12, 10, 6, 10))
        left.pack(side='left', fill='both')

        right = ttk.Frame(body, padding=(6, 10, 12, 10))
        right.pack(side='right', fill='both', expand=True)

        self.notebook = ttk.Notebook(left)
        self.notebook.pack(fill='both', expand=True)

        self.tab_vertical = ttk.Frame(self.notebook, padding=12)
        self.tab_circle = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(self.tab_vertical, text='竖切折页')
        self.notebook.add(self.tab_circle, text='圆切贴纸')
        self.notebook.bind('<<NotebookTabChanged>>', self._on_tab_change)

        self._build_vertical_tab(self.tab_vertical)
        self._build_circle_tab(self.tab_circle)

        self.preview = PreviewPane(right)
        self.preview.pack(fill='both', expand=True)
        self.preview.on_render = self._redraw_overlay

    def _build_vertical_tab(self, parent):
        self.mode = tk.StringVar(value='equal')
        row_mode = ttk.Frame(parent)
        row_mode.pack(fill='x', pady=(0, 8))
        ttk.Radiobutton(row_mode, text='等分', variable=self.mode,
                        value='equal', command=self._on_mode).pack(side='left')
        ttk.Radiobutton(row_mode, text='自定义 (mm)', variable=self.mode,
                        value='custom', command=self._on_mode).pack(side='left', padx=14)

        self.f_input = ttk.Frame(parent)
        self.f_input.pack(fill='x')

        self.f_equal = ttk.Frame(self.f_input)
        ttk.Label(self.f_equal, text='份数').pack(side='left')
        self.entry_n = ttk.Entry(self.f_equal, width=6)
        self.entry_n.insert(0, '3')
        self.entry_n.pack(side='left', padx=(6, 0))
        self.entry_n.bind('<KeyRelease>', lambda _: self._refresh())

        self.f_custom = ttk.Frame(self.f_input)
        ttk.Label(self.f_custom, text='段宽').pack(side='left')
        self.entry_custom = ttk.Entry(self.f_custom, width=26)
        self.entry_custom.pack(side='left', padx=(6, 0))
        self.entry_custom.bind('<KeyRelease>', lambda _: self._refresh())
        ttk.Label(self.f_custom, text='mm，逗号分隔',
                  foreground='#888').pack(side='left', padx=(6, 0))

        self.f_equal.pack(fill='x')

        row_bleed = ttk.Frame(parent)
        row_bleed.pack(fill='x', pady=(10, 0))
        ttk.Label(row_bleed, text='外框出血').pack(side='left')
        self.entry_inset = ttk.Entry(row_bleed, width=5)
        self.entry_inset.insert(0, '0.3')
        self.entry_inset.pack(side='left', padx=(6, 0))
        self.entry_inset.bind('<KeyRelease>', lambda _: self._refresh())
        ttk.Label(row_bleed, text='mm / 边', foreground='#888').pack(side='left', padx=(4, 0))

        ttk.Button(parent, text='生成竖切 PDF', style='Primary.TButton',
                   command=self._generate_vertical).pack(fill='x', pady=(14, 10))

        self.txt_vertical = self._make_info_text(parent, '  切割位置  ')

    def _build_circle_tab(self, parent):
        row = ttk.Frame(parent)
        row.pack(fill='x')
        ttk.Label(row, text='出血').pack(side='left')
        self.entry_circle_bleed = ttk.Entry(row, width=5)
        self.entry_circle_bleed.insert(0, '2')
        self.entry_circle_bleed.pack(side='left', padx=(6, 0))
        self.entry_circle_bleed.bind('<KeyRelease>', lambda _: self._refresh())
        ttk.Label(row, text='mm（圆距离图边的距离）',
                  foreground='#888').pack(side='left', padx=(6, 0))

        ttk.Button(parent, text='生成圆切 PDF', style='Primary.TButton',
                   command=self._generate_circle).pack(fill='x', pady=(14, 10))

        self.txt_circle = self._make_info_text(parent, '  圆切参数  ')

    def _make_info_text(self, parent, title):
        frame = ttk.LabelFrame(parent, text=title)
        frame.pack(fill='both', expand=True, pady=(6, 0))
        is_aqua = self.root.tk.call('tk', 'windowingsystem') == 'aqua'
        txt = tk.Text(frame, width=36, height=8, state='disabled',
                      bg=self.BG_INPUT, fg=self.FG, relief='flat', bd=0,
                      highlightthickness=1, highlightbackground='#c2c2c2',
                      insertbackground=self.FG,
                      font=('Menlo', 11) if is_aqua else ('Consolas', 10))
        txt.tag_config('warn', foreground='#c0392b')
        txt.tag_config('header', foreground=self.FG_MUTED)
        txt.pack(fill='both', expand=True)
        return txt

    # ── 事件 ───────────────────────────────────
    def _on_tab_change(self, _event=None):
        idx = self.notebook.index('current')
        self.active_tool = 'vertical' if idx == 0 else 'circle'
        self._refresh()

    def _on_mode(self):
        if self.mode.get() == 'equal':
            self.f_custom.pack_forget()
            self.f_equal.pack(fill='x')
        else:
            self.f_equal.pack_forget()
            self.f_custom.pack(fill='x')
        self._refresh()

    # ── 文件打开 ───────────────────────────────
    def _open(self):
        path = filedialog.askopenfilename(filetypes=[
            ('PDF 或图片', '*.pdf *.png *.jpg *.jpeg *.tif *.tiff'),
            ('PDF 文件', '*.pdf'),
            ('图片 (PNG/JPG/TIFF)', '*.png *.jpg *.jpeg *.tif *.tiff'),
        ])
        if not path:
            return

        self._cleanup_temp()

        try:
            if is_image(path):
                w_mm, h_mm = read_image_size_mm(path)
                working_pdf = image_to_pdf(path, w_mm, h_mm)
                self._temp_pdf = working_pdf
            else:
                with pikepdf.open(path) as pdf:
                    mb = pdf.pages[0].mediabox
                    w_mm = (float(mb[2]) - float(mb[0])) / PT_PER_MM
                    h_mm = (float(mb[3]) - float(mb[1])) / PT_PER_MM
                working_pdf = path
        except Exception as e:
            messagebox.showerror('打开失败', str(e))
            return

        self.src_path = path
        self.pdf_path = working_pdf
        self.page_w_mm = w_mm
        self.page_h_mm = h_mm
        self.lbl_file.config(text=os.path.basename(path), foreground='#222')
        self.lbl_size.config(text=f'{w_mm:.1f} × {h_mm:.1f} mm')

        # 方图自动跳圆切 tab，非方图跳竖切
        is_square = abs(w_mm - h_mm) < 1.0
        self.notebook.select(1 if is_square else 0)  # 触发 <<NotebookTabChanged>>

        self.preview.load_pdf(working_pdf)
        self._refresh()

    def _cleanup_temp(self):
        if self._temp_pdf and os.path.exists(self._temp_pdf):
            try:
                os.unlink(self._temp_pdf)
            except OSError:
                pass
        self._temp_pdf = None

    def _on_close(self):
        self._cleanup_temp()
        self.root.destroy()

    # ── 竖切参数计算 ───────────────────────────
    def _parse_custom_segments(self):
        raw = self.entry_custom.get()
        try:
            segs = [float(p.strip()) for p in raw.split(',') if p.strip()]
        except ValueError:
            return None
        if not segs:
            return None
        cuts = []
        total = 0.0
        for s in segs:
            total += s
            cuts.append(total)
        remaining = self.page_w_mm - total
        return segs, cuts, remaining

    def _positions(self):
        if self.mode.get() == 'equal':
            try:
                n = int(self.entry_n.get())
                if n < 2 or self.page_w_mm == 0:
                    return []
                step = self.page_w_mm / n
                return [step * i for i in range(1, n)]
            except ValueError:
                return []
        else:
            result = self._parse_custom_segments()
            if result is None:
                return []
            _, cuts, _ = result
            return cuts

    def _current_inset_preview(self):
        try:
            return float(self.entry_inset.get())
        except ValueError:
            return 0.3

    def _current_circle_bleed_preview(self):
        try:
            return float(self.entry_circle_bleed.get())
        except ValueError:
            return 2.0

    # ── 预览/文本刷新 ──────────────────────────
    def _refresh(self):
        for txt in (self.txt_vertical, self.txt_circle):
            txt.config(state='normal')
            txt.delete('1.0', 'end')

        if not self.page_w_mm:
            msg = '（请先选择 PDF 或图片）'
            self.txt_vertical.insert('1.0', msg, 'header')
            self.txt_circle.insert('1.0', msg, 'header')
        else:
            self._fill_vertical_info()
            self._fill_circle_info()

        for txt in (self.txt_vertical, self.txt_circle):
            txt.config(state='disabled')
        self._redraw_overlay()

    def _fill_vertical_info(self):
        txt = self.txt_vertical
        if self.mode.get() == 'equal':
            positions = self._positions()
            if positions:
                lines = [f'  {i+1:>2}.  x = {p:>7.2f} mm' for i, p in enumerate(positions)]
                txt.insert('1.0', '\n'.join(lines))
            else:
                txt.insert('1.0', '（请输入份数）', 'header')
        else:
            result = self._parse_custom_segments()
            if result:
                segs, cuts, remaining = result
                warn_remaining = 0 < remaining < 100
                txt.insert('end', '  段宽：\n', 'header')
                seg_lines = [f'    {i+1:>2}.  {s:>7.2f} mm' for i, s in enumerate(segs)]
                txt.insert('end', '\n'.join(seg_lines))
                if remaining != 0.0:
                    txt.insert('end', '\n\n  剩余：', 'header')
                    rem_text = f' {remaining:.2f} mm'
                    if warn_remaining:
                        rem_text += '   ⚠ 尾料不足 10cm'
                    txt.insert('end', rem_text, 'warn' if warn_remaining else '')
                if cuts:
                    txt.insert('end', '\n\n  切割位置：\n', 'header')
                    txt.insert('end', '\n'.join(f'    x = {x:>7.2f} mm' for x in cuts))
            else:
                txt.insert('1.0', '（输入段宽，如: 200, 200, 200）', 'header')

    def _fill_circle_info(self):
        txt = self.txt_circle
        bleed = self._current_circle_bleed_preview()
        side = min(self.page_w_mm, self.page_h_mm)
        r = side / 2 - bleed
        if r <= 0:
            txt.insert('1.0', f'  出血 {bleed}mm 过大，圆半径 ≤ 0', 'warn')
            return
        cx = self.page_w_mm / 2
        cy = self.page_h_mm / 2
        txt.insert('end',
                   f'  圆心:   ({cx:.2f}, {cy:.2f}) mm\n'
                   f'  半径:   {r:.2f} mm\n'
                   f'  直径:   {r * 2:.2f} mm\n'
                   f'  出血:   {bleed:.2f} mm')

    def _redraw_overlay(self):
        if not self.page_w_mm:
            self.preview.clear_overlay()
            return
        if self.active_tool == 'circle':
            bleed = self._current_circle_bleed_preview()
            r = min(self.page_w_mm, self.page_h_mm) / 2 - bleed
            self.preview.set_overlay_circle(
                self.page_w_mm / 2, self.page_h_mm / 2, r,
            )
        else:
            self.preview.set_overlay(
                self._positions(),
                inset_mm=self._current_inset_preview(),
                add_border=True,
            )

    # ── 生成 ──────────────────────────────────
    def _parse_strict(self, entry, label):
        raw = entry.get().strip()
        if not raw:
            raise ValueError(f'{label} 未填写')
        try:
            return float(raw)
        except ValueError:
            raise ValueError(f'{label} 不是有效数字：{raw!r}')

    def _generate_vertical(self):
        if not self.pdf_path:
            messagebox.showwarning('提示', '请先选择文件')
            return
        positions = self._positions()
        if not positions:
            messagebox.showwarning('提示', '请输入有效参数')
            return
        try:
            inset_mm = self._parse_strict(self.entry_inset, '外框出血')
        except ValueError as e:
            messagebox.showerror('参数错误', str(e))
            return
        out = self._make_out_path()
        try:
            add_vertical_cuts(self.pdf_path, positions, out, inset_mm=inset_mm)
            messagebox.showinfo('完成', f'已保存：\n{out}')
        except Exception as e:
            messagebox.showerror('生成失败', str(e))

    def _generate_circle(self):
        if not self.pdf_path:
            messagebox.showwarning('提示', '请先选择文件')
            return
        try:
            bleed_mm = self._parse_strict(self.entry_circle_bleed, '出血')
        except ValueError as e:
            messagebox.showerror('参数错误', str(e))
            return
        out = self._make_out_path()
        try:
            add_circle_cut(self.pdf_path, bleed_mm, out)
            messagebox.showinfo('完成', f'已保存：\n{out}')
        except Exception as e:
            messagebox.showerror('生成失败', str(e))

    def _make_out_path(self):
        base, _ = os.path.splitext(self.src_path)
        return base + '_cut.pdf'


if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
