import os
import tempfile

import pikepdf
from pikepdf import Array, Dictionary, Name

PT_PER_MM = 2.83465

IMAGE_EXTS = {'.png', '.jpg', '.jpeg', '.tif', '.tiff'}


def is_image(path):
    return os.path.splitext(path)[1].lower() in IMAGE_EXTS


def read_image_size_mm(image_path):
    """从图片 DPI 元数据算出物理尺寸（mm）。依赖员工在 PS 里正确设置 DPI。"""
    from PIL import Image
    with Image.open(image_path) as img:
        w_px, h_px = img.size
        dpi = img.info.get('dpi')
    if not dpi or dpi[0] == 0 or dpi[1] == 0:
        raise ValueError(
            f'图片无 DPI 信息（或为 0）：{os.path.basename(image_path)}\n'
            '请在 PS 里用"图像 → 图像大小"设置物理尺寸后再导入。'
        )
    dpi_x, dpi_y = float(dpi[0]), float(dpi[1])
    w_mm = w_px / dpi_x * 25.4
    h_mm = h_px / dpi_y * 25.4
    return w_mm, h_mm


def image_to_pdf(image_path, w_mm, h_mm):
    """把图片转成单页 PDF（临时文件），页面尺寸 = w_mm × h_mm。"""
    import pymupdf
    wp = w_mm * PT_PER_MM
    hp = h_mm * PT_PER_MM
    doc = pymupdf.open()
    page = doc.new_page(width=wp, height=hp)
    page.insert_image(pymupdf.Rect(0, 0, wp, hp), filename=image_path)
    fd, tmp = tempfile.mkstemp(suffix='.pdf', prefix='cut_img_')
    os.close(fd)
    doc.save(tmp)
    doc.close()
    return tmp


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
        C1=Array([0.0, 1.0, 1.0, 0.0]),  # CMYK 红色，便于肉眼区分（机器靠专色名识别，tint 不影响）
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


def build_circle_cut_stream(mc_key, cs_name, cx, cy, r):
    """四段三次贝塞尔逼近圆。cx/cy/r 均为 pt。"""
    k = 0.5522847498 * r
    lines = [f'/Layer /{mc_key} BDC'.encode()]
    lines.append(b'q')
    lines.append(f'/{cs_name} CS 1 SCN'.encode())
    lines.append(b'0.5 w')
    lines.append(f'{cx - r:.4f} {cy:.4f} m'.encode())
    lines.append(f'{cx - r:.4f} {cy + k:.4f} {cx - k:.4f} {cy + r:.4f} {cx:.4f} {cy + r:.4f} c'.encode())
    lines.append(f'{cx + k:.4f} {cy + r:.4f} {cx + r:.4f} {cy + k:.4f} {cx + r:.4f} {cy:.4f} c'.encode())
    lines.append(f'{cx + r:.4f} {cy - k:.4f} {cx + k:.4f} {cy - r:.4f} {cx:.4f} {cy - r:.4f} c'.encode())
    lines.append(f'{cx - k:.4f} {cy - r:.4f} {cx - r:.4f} {cy - k:.4f} {cx - r:.4f} {cy:.4f} c'.encode())
    lines.append(b'h S')
    lines.append(b'Q')
    lines.append(b'EMC')
    return b'\n'.join(lines)


def add_circle_cut(pdf_path, bleed_mm, out_path):
    """画圆切线。圆以页面中心为圆心，半径 = min(宽, 高)/2 - bleed_mm。"""
    with pikepdf.open(pdf_path) as pdf:
        for page in pdf.pages:
            mb = page.mediabox
            x0, y0 = float(mb[0]), float(mb[1])
            x1, y1 = float(mb[2]), float(mb[3])
            w, h = x1 - x0, y1 - y0
            cx = (x0 + x1) / 2
            cy = (y0 + y1) / 2
            r = min(w, h) / 2 - bleed_mm * PT_PER_MM
            if r <= 0:
                raise ValueError(f'出血 {bleed_mm}mm 过大，圆半径 ≤ 0')

            cs_name, _ = ensure_cut_cs(page)
            mc_key = _add_cut_layer_prop(page)
            cut_data = build_circle_cut_stream(mc_key, cs_name, cx, cy, r)
            new_stream = pdf.make_stream(cut_data)

            if '/PieceInfo' in page:
                del page['/PieceInfo']

            contents = page.get('/Contents')
            if contents is None:
                page['/Contents'] = new_stream
            elif isinstance(contents, pikepdf.Array):
                contents.append(new_stream)
            else:
                page['/Contents'] = pikepdf.Array([contents, new_stream])

        pdf.save(out_path)
