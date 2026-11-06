import os
import math
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
        if img.format == 'JPEG':
            unit = img.info.get('jfif_unit')
            density = img.info.get('jfif_density')
            dpi = None
            if unit in (1, 2) and density:
                factor = 2.54 if unit == 2 else 1
                dpi = tuple(float(value) * factor for value in density)
            else:
                exif = img.getexif()
                unit = exif.get(296)
                if unit in (2, 3) and exif.get(282) and exif.get(283):
                    factor = 2.54 if unit == 3 else 1
                    dpi = (float(exif[282]) * factor, float(exif[283]) * factor)
        elif img.format == 'TIFF':
            tags = img.tag_v2
            dpi = None
            unit = tags.get(296, 2)
            if unit in (2, 3) and tags.get(282) and tags.get(283):
                factor = 2.54 if unit == 3 else 1
                dpi = (float(tags[282]) * factor, float(tags[283]) * factor)
    if not dpi or any(not math.isfinite(float(value)) or float(value) <= 0 for value in dpi):
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
    page.insert_image(pymupdf.Rect(0, 0, wp, hp), filename=image_path, keep_proportion=False)
    fd, tmp = tempfile.mkstemp(suffix='.pdf', prefix='cut_img_')
    os.close(fd)
    doc.save(tmp)
    doc.close()
    return tmp


def find_cut_cs(page):
    resources = page.get('/Resources', Dictionary())
    colorspaces = resources.get('/ColorSpace', Dictionary())
    for key, cs in colorspaces.items():
        if isinstance(cs, Array) and len(cs) >= 2:
            if str(cs[0]) == '/Separation' and str(cs[1]) == '/CUT':
                return key.lstrip('/')
    return None


def ensure_cut_cs(page):
    existing = find_cut_cs(page)
    if existing is not None:
        return existing, False

    if '/Resources' not in page:
        page['/Resources'] = Dictionary()
    res = page['/Resources']
    if '/ColorSpace' not in res:
        res['/ColorSpace'] = Dictionary()
    cs_dict = res['/ColorSpace']

    tint_fn = Dictionary(
        FunctionType=2,
        Domain=Array([0, 1]),
        C0=Array([0.0, 0.0, 0.0, 0.0]),
        C1=Array([0.0, 1.0, 1.0, 0.0]),  # CMYK 红色，便于肉眼区分（机器靠专色名识别，tint 不影响）
        N=1.0,
        Range=Array([0.0, 1.0, 0.0, 1.0, 0.0, 1.0, 0.0, 1.0]),
    )
    key = 'CS_CUT'
    index = 1
    while f'/{key}' in cs_dict:
        key = f'CS_CUT_{index}'
        index += 1
    cs_dict[f'/{key}'] = Array([Name('/Separation'), Name('/CUT'), Name('/DeviceCMYK'), tint_fn])
    return key, True


def page_geometry(page):
    """显示方向的物理尺寸（pt），以及到原 PDF 坐标的变换。"""
    x0, y0, x1, y1 = map(float, page.mediabox)
    unit = float(page.get('/UserUnit', 1))
    rotation = int(page.get('/Rotate', 0)) % 360
    if not math.isfinite(unit) or unit <= 0 or rotation not in (0, 90, 180, 270):
        raise ValueError('PDF 页面缩放或旋转参数无效')
    width, height = (x1 - x0) * unit, (y1 - y0) * unit
    if not all(math.isfinite(value) and value > 0 for value in (width, height)):
        raise ValueError('PDF 页面尺寸无效')
    scale = 1 / unit
    matrices = {
        0: (scale, 0, 0, scale, x0, y0),
        90: (0, scale, -scale, 0, x1, y0),
        180: (-scale, 0, 0, -scale, x1, y1),
        270: (0, -scale, scale, 0, x0, y1),
    }
    if rotation in (90, 270):
        width, height = height, width
    return width, height, matrices[rotation]


def _prepare_cut_resources(page):
    # 每页独立修改资源表，避免原 PDF 的共享资源影响其他页。
    resources = Dictionary(page.get('/Resources', Dictionary()))
    for key in ('/ColorSpace', '/Properties'):
        resources[key] = Dictionary(resources.get(key, Dictionary()))
    page.Resources = resources
    cs_name, is_new = ensure_cut_cs(page)
    return cs_name, is_new, _add_cut_layer_prop(page)


def _append_cut_stream(pdf, page, data, matrix):
    depth = 0
    for _, operator in pikepdf.parse_content_stream(page):
        if str(operator) == 'q':
            depth += 1
        elif str(operator) == 'Q':
            depth -= 1
            if depth < 0:
                raise ValueError('PDF 原稿的绘图状态不完整，请重新导出 PDF')
    contents = page.get('/Contents')
    original = list(contents) if isinstance(contents, Array) else []
    if contents is not None and not isinstance(contents, Array):
        original = [contents]
    restore = b'Q\n' * (depth + 1) + b'n\n'
    transform = ' '.join(format(value, '.12g') for value in matrix).encode() + b' cm\n'
    page.Contents = Array([
        pdf.make_stream(b'q\n'), *original,
        pdf.make_stream(restore + b'q\n' + transform + data + b'\nQ\n'),
    ])
    if '/PieceInfo' in page:
        del page['/PieceInfo']


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


def add_vertical_cuts(pdf_path, positions_mm, out_path, inset_mm=0.3, *, add_border=True, equal_parts=None):
    with pikepdf.open(pdf_path) as pdf:
        for page in pdf.pages:
            width, height, matrix = page_geometry(page)
            if equal_parts is not None:
                if equal_parts < 2:
                    raise ValueError('请输入份数（至少 2）')
                positions_pt = [width * i / equal_parts for i in range(1, equal_parts)]
            else:
                positions_pt = [p * PT_PER_MM for p in positions_mm]
            if any(not math.isfinite(p) or p < 0 or p > width + 1e-7 for p in positions_pt):
                raise ValueError('切割位置超出页面宽度')
            positions_pt = [p for p in positions_pt if 1e-7 < p < width - 1e-7]
            inset_pt = inset_mm * PT_PER_MM
            if add_border and (not math.isfinite(inset_pt) or inset_pt < 0 or 2 * inset_pt >= min(width, height)):
                raise ValueError('外框出血须大于等于 0，且小于页面短边的一半')
            cs_name, is_new, mc_key = _prepare_cut_resources(page)
            cut_data = build_cut_stream(mc_key, cs_name, positions_pt, 0, 0, width, height,
                                        add_border=add_border and is_new, inset_pt=inset_pt)
            _append_cut_stream(pdf, page, cut_data, matrix)

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


def add_crop_marks(pdf_path, out_path,
                   margin_mm=10.0, mark_len_mm=3.0, gap_mm=3.0, lw_pt=0.25):
    """
    在每页 4 角加标准裁切标记（人眼对齐用，黑色），
    并在底部 margin 区域打印 文件名·页码·路径。
    把 MediaBox 向外各扩 margin_mm，原内容位置不变，在原 MediaBox 四角画 L 形标记。
    """
    margin_pt = margin_mm * PT_PER_MM
    mlen = mark_len_mm * PT_PER_MM
    gap = gap_mm * PT_PER_MM

    # Phase 1: pikepdf 扩 MediaBox + 画 L 标记（原封不动保留 /CUT 专色等）
    with pikepdf.open(pdf_path) as pdf:
        for page in pdf.pages:
            mb = page.mediabox
            x0, y0 = float(mb[0]), float(mb[1])
            x1, y1 = float(mb[2]), float(mb[3])

            lines = [b'q', b'0 0 0 1 K', f'{lw_pt:.3f} w'.encode()]
            for cx, cy, sx, sy in [
                (x0, y0, -1, -1), (x1, y0, +1, -1),
                (x0, y1, -1, +1), (x1, y1, +1, +1),
            ]:
                lines.append(f'{cx + sx * gap:.4f} {cy:.4f} m'.encode())
                lines.append(f'{cx + sx * (gap + mlen):.4f} {cy:.4f} l'.encode())
                lines.append(f'{cx:.4f} {cy + sy * gap:.4f} m'.encode())
                lines.append(f'{cx:.4f} {cy + sy * (gap + mlen):.4f} l'.encode())
            lines.append(b'S')
            lines.append(b'Q')
            new_stream = pdf.make_stream(b'\n'.join(lines))

            page.mediabox = [x0 - margin_pt, y0 - margin_pt,
                             x1 + margin_pt, y1 + margin_pt]

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

    # Phase 2: pymupdf 追加页脚文字（支持中文，增量保存不破坏原资源）
    import pymupdf
    src_name = os.path.basename(pdf_path)
    src_abspath = os.path.abspath(pdf_path)
    doc = pymupdf.open(out_path)
    total = doc.page_count
    try:
        font_size = 10
        for i in range(total):
            page = doc[i]
            rect = page.rect  # 扩展后的新页面（display 坐标, y-down）
            info = f'{src_name}   Page {i + 1}/{total}   {src_abspath}'
            page.insert_font(fontname='china-s')
            text_w = pymupdf.get_text_length(info, fontname='china-s', fontsize=font_size)
            # 水平居中；垂直居中于底部 margin 带（贴到内容边也无妨，出血会裁掉）
            x = (rect.x0 + rect.x1) / 2 - text_w / 2
            y = rect.y1 - margin_pt / 2 + font_size / 2
            page.insert_text(pymupdf.Point(x, y), info,
                             fontname='china-s', fontsize=font_size,
                             color=(0, 0, 0))
        doc.save(out_path, incremental=True,
                 encryption=pymupdf.PDF_ENCRYPT_KEEP)
    finally:
        doc.close()


def split_combine_pdf(entries, out_path):
    """
    把多个 PDF 的页面按左半/右半/整页任意组合，输出新 PDF。
    entries: [(src_pdf_path, page_index, 'L'|'R'|'F'), ...]
    只改 MediaBox，不重编内容流，矢量/字体/ICC/专色全部保留。
    """
    src_cache = {}
    out = pikepdf.Pdf.new()
    try:
        for src, idx, half in entries:
            if src not in src_cache:
                src_cache[src] = pikepdf.open(src)
            src_pdf = src_cache[src]
            page = src_pdf.pages[idx]
            out.pages.append(page)
            new_page = out.pages[-1]
            mb = new_page.mediabox
            x0, y0 = float(mb[0]), float(mb[1])
            x1, y1 = float(mb[2]), float(mb[3])
            mid = (x0 + x1) / 2
            if half == 'L':
                new_page.mediabox = [x0, y0, mid, y1]
            elif half == 'R':
                new_page.mediabox = [mid, y0, x1, y1]
        out.save(out_path)
    finally:
        for pdf in src_cache.values():
            pdf.close()


def add_circle_cut(pdf_path, bleed_mm, out_path):
    """画圆切线。圆以页面中心为圆心，半径 = min(宽, 高)/2 - bleed_mm。"""
    if not math.isfinite(bleed_mm) or bleed_mm < 0:
        raise ValueError('圆切出血必须是大于等于 0 的有效数值')
    with pikepdf.open(pdf_path) as pdf:
        for page in pdf.pages:
            w, h, matrix = page_geometry(page)
            cx = w / 2
            cy = h / 2
            r = min(w, h) / 2 - bleed_mm * PT_PER_MM
            if r <= 0:
                raise ValueError(f'出血 {bleed_mm}mm 过大，圆半径 ≤ 0')

            cs_name, _, mc_key = _prepare_cut_resources(page)
            cut_data = build_circle_cut_stream(mc_key, cs_name, cx, cy, r)
            _append_cut_stream(pdf, page, cut_data, matrix)

        pdf.save(out_path)
