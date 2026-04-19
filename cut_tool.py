import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pikepdf
from pikepdf import Array, Dictionary, Name

PT_PER_MM = 2.83465


def ensure_cut_cs(page):
    if '/Resources' not in page:
        page['/Resources'] = Dictionary()
    res = page['/Resources']
    if '/ColorSpace' not in res:
        res['/ColorSpace'] = Dictionary()
    cs_dict = res['/ColorSpace']

    for key in cs_dict.keys():
        cs = cs_dict[key]
        if isinstance(cs, Array) and len(cs) >= 2:
            if str(cs[0]) == '/Separation' and str(cs[1]) == '/CUT':
                return key.lstrip('/'), False

    tint_fn = Dictionary(
        FunctionType=2,
        Domain=Array([0, 1]),
        C0=Array([0.0, 0.0, 0.0, 0.0]),
        C1=Array([0.0, 0.0, 0.0, 1.0]),
        N=1.0,
        Range=Array([0.0, 1.0, 0.0, 1.0, 0.0, 1.0, 0.0, 1.0]),
    )
    cs_dict['/CS_CUT'] = Array([Name('/Separation'), Name('/CUT'), Name('/DeviceCMYK'), tint_fn])
    return 'CS_CUT', True


def _add_cut_layer_prop(page):
    """在 Properties 里注册 CUT 图层，返回 key（不含 /）。"""
    res = page['/Resources']
    if '/Properties' not in res:
        res['/Properties'] = Dictionary()
    props = res['/Properties']
    i = 1
    while f'/MC{i}' in props:
        i += 1
    mc_key = f'MC{i}'
    props[f'/{mc_key}'] = Dictionary(
        Color=Array([65535, 0, 0]),
        Dimmed=False,
        Editable=True,
        Preview=True,
        Printed=True,
        Title='CUT',
        Visible=True,
    )
    return mc_key


def build_cut_stream(mc_key, cs_name, positions_pt, x0, y0, x1, y1,
                     add_border=False, inset_pt=0.0):
    """生成完整的 CUT 图层流（独立 BDC/EMC，不依赖外部图形状态）。"""
    lines = [f'/Layer /{mc_key} BDC'.encode()]
    lines.append(b'q')
    lines.append(f'/{cs_name} CS 1 SCN'.encode())
    lines.append(b'0.5 w')
    if add_border:
        bx0, by0 = x0 + inset_pt, y0 + inset_pt
        bx1, by1 = x1 - inset_pt, y1 - inset_pt
        lines.append(f'{bx0:.4f} {by0:.4f} {bx1 - bx0:.4f} {by1 - by0:.4f} re'.encode())
        lines.append(b'S')
    for x in positions_pt:
        lines.append(f'{x:.4f} {y0:.4f} m'.encode())
        lines.append(f'{x:.4f} {y1:.4f} l'.encode())
        lines.append(b'S')
    lines.append(b'Q')
    lines.append(b'EMC')
    return b'\n'.join(lines)


def add_vertical_cuts(pdf_path, positions_mm, out_path, inset_mm=0.3):
    with pikepdf.open(pdf_path) as pdf:
        for page in pdf.pages:
            mb = page.mediabox
            x0, y0 = float(mb[0]), float(mb[1])
            x1, y1 = float(mb[2]), float(mb[3])

            cs_name, is_new = ensure_cut_cs(page)
            mc_key = _add_cut_layer_prop(page)
            positions_pt = [x0 + p * PT_PER_MM for p in positions_mm]
            inset_pt = inset_mm * PT_PER_MM
            cut_data = build_cut_stream(mc_key, cs_name, positions_pt, x0, y0, x1, y1,
                                        add_border=is_new, inset_pt=inset_pt)
            new_stream = pdf.make_stream(cut_data)

            # 去掉 AI 私有元数据，否则 Illustrator 会忽略我们写入的内容流
            if '/PieceInfo' in page:
                del page['/PieceInfo']

            contents = page.get('/Contents')
            if contents is None:
                page['/Contents'] = new_stream
            elif isinstance(contents, pikepdf.Array):
                contents.append(new_stream)
            else:
                # 单流转数组，原始流完全不动
                page['/Contents'] = pikepdf.Array([contents, new_stream])

        pdf.save(out_path)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title('PDF CUT 切割线工具')
        self.root.resizable(False, False)
        self.pdf_path = None
        self.page_w_mm = 0.0
        self._build()

    def _build(self):
        # ── 文件选择 ──────────────────────────────
        f0 = ttk.Frame(self.root)
        f0.pack(fill='x', padx=10, pady=8)
        ttk.Button(f0, text='选择 PDF', command=self._open).pack(side='left')
        self.lbl_file = ttk.Label(f0, text='未选择文件', foreground='gray')
        self.lbl_file.pack(side='left', padx=8)
        self.lbl_size = ttk.Label(f0, text='')
        self.lbl_size.pack(side='left')

        ttk.Separator(self.root).pack(fill='x', padx=10)

        # ── 模式 ─────────────────────────────────
        f1 = ttk.Frame(self.root)
        f1.pack(fill='x', padx=10, pady=6)
        self.mode = tk.StringVar(value='equal')
        ttk.Radiobutton(f1, text='等分', variable=self.mode,
                        value='equal', command=self._on_mode).pack(side='left')
        ttk.Radiobutton(f1, text='自定义 (mm)', variable=self.mode,
                        value='custom', command=self._on_mode).pack(side='left', padx=12)

        # ── 输入容器 ──────────────────────────────
        self.f_input = ttk.Frame(self.root)
        self.f_input.pack(fill='x', padx=10)

        self.f_equal = ttk.Frame(self.f_input)
        ttk.Label(self.f_equal, text='份数:').pack(side='left')
        self.entry_n = ttk.Entry(self.f_equal, width=6)
        self.entry_n.insert(0, '3')
        self.entry_n.pack(side='left', padx=5)
        self.entry_n.bind('<KeyRelease>', lambda _: self._refresh())

        self.f_custom = ttk.Frame(self.f_input)
        ttk.Label(self.f_custom, text='段宽 (mm，逗号分隔):').pack(side='left')
        self.entry_custom = ttk.Entry(self.f_custom, width=28)
        self.entry_custom.pack(side='left', padx=5)
        self.entry_custom.bind('<KeyRelease>', lambda _: self._refresh())

        self.f_equal.pack(fill='x')

        ttk.Separator(self.root).pack(fill='x', padx=10, pady=6)

        # ── 出血收缩 ──────────────────────────────
        f_bleed = ttk.Frame(self.root)
        f_bleed.pack(fill='x', padx=10, pady=(0, 4))
        ttk.Label(f_bleed, text='外框出血收缩:').pack(side='left')
        self.entry_inset = ttk.Entry(f_bleed, width=5)
        self.entry_inset.insert(0, '0.3')
        self.entry_inset.pack(side='left', padx=4)
        ttk.Label(f_bleed, text='mm / 边').pack(side='left')

        # ── 预览 ──────────────────────────────────
        f4 = ttk.Frame(self.root)
        f4.pack(fill='both', expand=True, padx=10)
        ttk.Label(f4, text='切割位置预览:').pack(anchor='w')
        self.txt = tk.Text(f4, height=6, state='disabled',
                           bg='#f8f8f8', relief='flat', font=('Consolas', 10))
        self.txt.tag_config('warn', foreground='red')
        self.txt.pack(fill='both', expand=True)

        ttk.Button(self.root, text='生成 PDF', command=self._generate).pack(pady=10)

    def _on_mode(self):
        if self.mode.get() == 'equal':
            self.f_custom.pack_forget()
            self.f_equal.pack(fill='x')
        else:
            self.f_equal.pack_forget()
            self.f_custom.pack(fill='x')
        self._refresh()

    def _open(self):
        path = filedialog.askopenfilename(filetypes=[('PDF 文件', '*.pdf')])
        if not path:
            return
        try:
            with pikepdf.open(path) as pdf:
                mb = pdf.pages[0].mediabox
                self.page_w_mm = (float(mb[2]) - float(mb[0])) / PT_PER_MM
                page_h_mm = (float(mb[3]) - float(mb[1])) / PT_PER_MM
        except Exception as e:
            messagebox.showerror('打开失败', str(e))
            return
        self.pdf_path = path
        self.lbl_file.config(text=os.path.basename(path), foreground='black')
        self.lbl_size.config(text=f'({self.page_w_mm:.1f} × {page_h_mm:.1f} mm)')
        self._refresh()

    def _parse_custom_segments(self):
        """返回 (segments, cut_positions, remaining) 或 None。"""
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
        """返回实际切割 X 坐标列表（mm）。"""
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

    def _refresh(self):
        self.txt.config(state='normal')
        self.txt.delete('1.0', 'end')

        if not self.page_w_mm:
            self.txt.insert('1.0', '（请先选择 PDF）')
            self.txt.config(state='disabled')
            return

        if self.mode.get() == 'equal':
            positions = self._positions()
            if positions:
                lines = [f'  {i+1}.  x = {p:.2f} mm' for i, p in enumerate(positions)]
                self.txt.insert('1.0', '\n'.join(lines))
            else:
                self.txt.insert('1.0', '（请输入份数）')
        else:
            result = self._parse_custom_segments()
            if result:
                segs, cuts, remaining = result
                warn_remaining = 0 < remaining < 100
                seg_lines = [f'  段 {i+1}:  {s:.2f} mm' for i, s in enumerate(segs)]
                self.txt.insert('end', '\n'.join(seg_lines))
                if remaining != 0.0:
                    rem_text = f'\n  剩余:  {remaining:.2f} mm'
                    if warn_remaining:
                        rem_text += '  ⚠ 尾料不足 10cm'
                    self.txt.insert('end', rem_text, 'warn' if warn_remaining else '')
                if cuts:
                    cut_lines = '\n\n  切割线位置:\n' + '\n'.join(f'    x = {x:.2f} mm' for x in cuts)
                    self.txt.insert('end', cut_lines)
            else:
                self.txt.insert('1.0', '（输入段宽，如: 50, 50）')

        self.txt.config(state='disabled')

    def _generate(self):
        if not self.pdf_path:
            messagebox.showwarning('提示', '请先选择 PDF 文件')
            return
        positions = self._positions()
        if not positions:
            messagebox.showwarning('提示', '请输入有效参数')
            return
        try:
            inset_mm = float(self.entry_inset.get())
        except ValueError:
            inset_mm = 0.3
        out = os.path.splitext(self.pdf_path)[0] + '_cut.pdf'
        try:
            add_vertical_cuts(self.pdf_path, positions, out, inset_mm=inset_mm)
            messagebox.showinfo('完成', f'已保存：\n{out}')
        except Exception as e:
            messagebox.showerror('生成失败', str(e))


if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
