"""以生成结果验证坐标、原稿保留与输入边界。只使用临时文件。"""
import asyncio
from io import BytesIO
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pikepdf
import pymupdf
from PIL import Image, ImageCms
from fastapi import UploadFile, HTTPException
from pdf_ops import PT_PER_MM, read_image_size_mm
from server import main


class CutGeometryTests(unittest.TestCase):
    def setUp(self):
        self.tmp = TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.old_data = main.DATA_DIR
        main.DATA_DIR = self.root

    def tearDown(self):
        main.DATA_DIR = self.old_data
        main._sessions.clear()
        self.tmp.cleanup()

    def source(self, widths=(160,), height=80, rotation=0, unit=1, stream=None, collision=False):
        path = self.root / "source.pdf"
        with pikepdf.Pdf.new() as pdf:
            for width in widths:
                page = pdf.add_blank_page(page_size=(width * PT_PER_MM, height * PT_PER_MM))
                page.Rotate = rotation
                page.UserUnit = unit
                if stream:
                    page.Contents = pdf.make_stream(stream)
                if collision:
                    page.Resources = pikepdf.Dictionary(
                        ColorSpace=pikepdf.Dictionary(CS_CUT=pikepdf.Name("/DeviceRGB"))
                    )
            pdf.save(path)
        return path

    def upload(self, path):
        return asyncio.run(main.create_session(UploadFile(file=BytesIO(path.read_bytes()), filename=path.name)))

    def generate(self, session, **body):
        return main.session_cut(session["id"], main.CutBody(**body)).path

    def test_original_transform_and_clip_are_isolated(self):
        original = b"1 0 0 1 60 0 cm 0 0 20 20 re W n 0 1 0 rg 0 0 20 20 re f"
        session = self.upload(self.source(stream=original))
        output = self.generate(session, mode="equal", n=4)
        with pikepdf.open(output) as pdf:
            self.assertEqual(pdf.pages[0].Contents[1].read_bytes(), original)
        with pymupdf.open(output) as pdf:
            paths = [d for d in pdf[0].get_drawings() if d["type"] == "s"]
            self.assertEqual(len(paths), 3)
            for drawing, expected in zip(paths, (40, 80, 120)):
                self.assertAlmostEqual(drawing["items"][0][1].x / PT_PER_MM, expected, places=3)
            pix = pdf[0].get_pixmap()
            for x in (40, 80, 120):
                colors = [pix.pixel(int(x * PT_PER_MM) + offset, pix.height // 2) for offset in (0, 1)]
                self.assertTrue(any(color[0] > color[1] + 30 for color in colors))

    def test_resource_name_collision_preserves_art(self):
        session = self.upload(self.source(stream=b"/CS_CUT cs 0 1 0 scn 0 0 40 40 re f", collision=True))
        output = self.generate(session, mode="equal", n=4)
        with pymupdf.open(output) as pdf:
            self.assertEqual(pdf[0].get_drawings()[0]["fill"], (0, 1, 0))

    def test_custom_exact_fill_has_three_internal_lines(self):
        session = self.upload(self.source())
        output = self.generate(session, mode="custom", segments="40,40,40,40")
        with pymupdf.open(output) as pdf:
            self.assertEqual(len(pdf[0].get_drawings()), 3)

    def test_each_page_is_divided_equally(self):
        session = self.upload(self.source(widths=(300, 200)))
        output = self.generate(session, mode="equal", n=4)
        with pymupdf.open(output) as pdf:
            for page, width in zip(pdf, (300, 200)):
                paths = page.get_drawings()
                self.assertEqual(len(paths), 3)
                for index, path in enumerate(paths, 1):
                    self.assertAlmostEqual(path["items"][0][1].x / PT_PER_MM, width * index / 4, places=3)

    def test_rotations_and_user_units_match_rendered_positions(self):
        for rotation in (0, 90, 180, 270):
            for unit in (1, 2):
                with self.subTest(rotation=rotation, unit=unit):
                    session = self.upload(self.source(rotation=rotation, unit=unit))
                    output = self.generate(session, mode="equal", n=4)
                    expected_width = (80 if rotation in (90, 270) else 160) * unit
                    self.assertAlmostEqual(session["w_mm"], expected_width)
                    with pymupdf.open(output) as pdf:
                        pix = pdf[0].get_pixmap()
                        xs = [x for x in range(pix.width) if pix.pixel(x, pix.height // 2)[0] >
                              pix.pixel(x, pix.height // 2)[1] + 30]
                        for index in (1, 2, 3):
                            expected = expected_width * PT_PER_MM * index / 4
                            self.assertTrue(any(abs(x - expected) < 1.5 for x in xs), (rotation, unit, xs))
                        self.assertTrue(all(min(abs(x - expected_width * PT_PER_MM * i / 4)
                                                for i in (1, 2, 3)) < 1.5 for x in xs))

    def test_invalid_bleeds_and_custom_overflow_are_rejected(self):
        session = self.upload(self.source(widths=(300, 160)))
        for body in (
            dict(mode="circle", bleed_mm=-1),
            dict(mode="circle", bleed_mm=float("nan")),
            dict(mode="circle", bleed_mm=40),
            dict(mode="equal", n=4, add_border=True, inset_mm=100),
            dict(mode="custom", segments="100,100"),
        ):
            with self.subTest(body=body), self.assertRaises(HTTPException) as raised:
                self.generate(session, **body)
            self.assertEqual(raised.exception.status_code, 400)

    def test_real_resolution_is_required(self):
        for fmt in ("JPEG", "TIFF", "PNG"):
            path = self.root / ("no-dpi." + fmt.lower())
            kwargs = {}
            if fmt == "JPEG":
                exif = Image.Exif()
                exif[274] = 6
                kwargs["exif"] = exif
            Image.new("RGB", (80, 40)).save(path, format=fmt, **kwargs)
            with self.subTest(fmt=fmt), self.assertRaises(ValueError):
                read_image_size_mm(path)

    def test_axis_dpi_and_icc_survive_image_conversion(self):
        path = self.root / "sample.jpg"
        icc = ImageCms.ImageCmsProfile(ImageCms.createProfile("sRGB")).tobytes()
        Image.new("RGB", (800, 400), (220, 80, 120)).save(path, dpi=(200, 100), icc_profile=icc)
        session = self.upload(path)
        self.assertAlmostEqual(session["w_mm"], 101.6, places=3)
        self.assertAlmostEqual(session["h_mm"], 101.6, places=3)
        for body in (dict(mode="equal", n=4), dict(mode="circle", bleed_mm=2)):
            output = self.generate(session, **body)
            with pikepdf.open(output) as pdf:
                image = next(iter(pdf.pages[0].get_images().values()))
                self.assertEqual(image.read_raw_bytes(), path.read_bytes())
                self.assertEqual(image.ColorSpace[1].read_bytes(), icc)
            with pymupdf.open(output) as pdf:
                self.assertEqual(tuple(pdf[0].rect), pdf[0].get_image_info()[0]["bbox"])
                self.assertEqual(pdf[0].get_text(), "")


if __name__ == "__main__":
    unittest.main()
