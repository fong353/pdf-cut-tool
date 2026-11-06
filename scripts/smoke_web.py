"""本机冒烟：核对竖切 PDF 的 CUT 专色、分段线和可选外框。"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pikepdf
from pikepdf import Array

from pdf_ops import PT_PER_MM, add_vertical_cuts
from server.main import equal_positions, parse_custom_segments


def _has_cut(page) -> bool:
    res = page.get("/Resources")
    if res is None or "/ColorSpace" not in res:
        return False
    for key in res["/ColorSpace"].keys():
        cs = res["/ColorSpace"][key]
        if isinstance(cs, Array) and len(cs) >= 2:
            if str(cs[0]) == "/Separation" and str(cs[1]) == "/CUT":
                return True
    return False


def main() -> None:
    assert equal_positions(300.0, 3) == [100.0, 200.0]
    segs, cuts, rem = parse_custom_segments("200, 200, 479.9", 879.9)
    assert segs == [200.0, 200.0, 479.9]
    assert abs(cuts[-1] - 879.9) < 1e-6
    assert abs(rem) < 1e-6

    w = 200 * PT_PER_MM
    h = 100 * PT_PER_MM
    with tempfile.TemporaryDirectory() as td:
        src = Path(td) / "in.pdf"
        out = Path(td) / "in_cut.pdf"
        pdf = pikepdf.Pdf.new()
        pdf.add_blank_page(page_size=(w, h))
        pdf.save(src)
        pdf.close()
        for add_border in (False, True):
            add_vertical_cuts(str(src), [50.0, 100.0, 150.0], str(out), add_border=add_border)
            with pikepdf.open(out) as cut:
                page = cut.pages[0]
                assert _has_cut(page), "输出缺少 /Separation /CUT"
                ops = [str(op) for _, op in pikepdf.parse_content_stream(page)]
                assert ops.count("m") == ops.count("l") == 3
                assert ops.count("re") == int(add_border)
    print("ok")


if __name__ == "__main__":
    main()
