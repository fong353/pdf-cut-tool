"""折页机 PDF CUT 切割线 Web 服务(内网,端口 8841)。

单进程 FastAPI。切割逻辑全部委托 pdf_ops(add_vertical_cuts / add_circle_cut),
web 层不重写内容流。会话文件落盘 data/jobs/<日期>/<id>/,保留 1 天。

CUT 注入是毫秒～数秒,生成接口同步返回文件,不抄 GPU 队列。
"""

from __future__ import annotations

import math
import shutil
import sys
import threading
import time
import unicodedata
import uuid
from datetime import datetime, timedelta
from pathlib import Path

import pikepdf
import pymupdf
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from pydantic import BaseModel, Field

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from pdf_ops import (
    PT_PER_MM,
    add_circle_cut,
    add_vertical_cuts,
    find_cut_cs,
    image_to_pdf,
    is_image,
    read_image_size_mm,
    page_geometry,
)

DATA_DIR = REPO_ROOT / "data" / "jobs"
STATIC = Path(__file__).resolve().parent / "static"

MAX_UPLOAD = 80 * 1024 * 1024
KEEP_DAYS = 1
CLEAN_INTERVAL = 3600
PREVIEW_MAX_PX = 1600
WIDTH_TOLERANCE_MM = 1e-9
ALLOWED_EXT = {".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff"}

app = FastAPI(title="PDF切割线", docs_url=None, redoc_url=None)

_sessions: dict[str, dict] = {}
_lock = threading.Lock()


# ---------- 参数计算(从 cut_tool.py 抄,不改 pdf_ops) ----------

def equal_positions(width_mm: float, n: int) -> list[float]:
    if n < 2 or width_mm <= 0:
        return []
    step = width_mm / n
    return [step * i for i in range(1, n)]


def parse_custom_segments(raw: str, width_mm: float) -> tuple[list[float], list[float], float]:
    segs = [float(p.strip()) for p in raw.split(",") if p.strip()]
    if not segs:
        raise ValueError("请输入段宽，如: 200, 200, 479.9")
    cuts: list[float] = []
    total = 0.0
    for s in segs:
        if not math.isfinite(s) or s <= 0:
            raise ValueError("段宽必须大于 0")
        total += s
        cuts.append(total)
    remaining = width_mm - total
    if not math.isfinite(total):
        raise ValueError("段宽合计超出允许宽度")
    if remaining < -WIDTH_TOLERANCE_MM:
        excess = -remaining
        amount = f"{excess:.2f} mm" if excess >= 0.01 else "不到 0.01 mm"
        raise ValueError(f"段宽合计超出允许宽度 {amount}")
    if remaining < 0:
        cuts[-1] = width_mm
        remaining = 0.0
    return segs, cuts, remaining


# ---------- 清理 ----------

def _prune_sessions() -> None:
    cutoff_ts = time.time() - KEEP_DAYS * 86400
    with _lock:
        stale = [sid for sid, s in _sessions.items() if s["created_at"] < cutoff_ts]
        for sid in stale:
            del _sessions[sid]


def _cleaner() -> None:
    while True:
        cutoff_dt = datetime.now() - timedelta(days=KEEP_DAYS)
        if DATA_DIR.is_dir():
            for day_dir in DATA_DIR.iterdir():
                try:
                    if datetime.strptime(day_dir.name, "%Y-%m-%d") < cutoff_dt:
                        shutil.rmtree(day_dir, ignore_errors=True)
                except ValueError:
                    continue
        _prune_sessions()
        time.sleep(CLEAN_INTERVAL)


DATA_DIR.mkdir(parents=True, exist_ok=True)
threading.Thread(target=_cleaner, daemon=True).start()


# ---------- helpers ----------

def _safe_stem(name: str) -> str:
    stem = Path(unicodedata.normalize("NFC", name)).stem
    cleaned = "".join(
        c for c in stem if c.isprintable() and c not in '/\\:*?"<>|'
    ).strip()
    return cleaned or "file"


def _session_or_404(session_id: str) -> dict:
    with _lock:
        session = _sessions.get(session_id)
    if not session:
        raise HTTPException(404, "会话不存在或已过期,请重新上传")
    pdf_path = Path(session["pdf_path"])
    if not pdf_path.is_file():
        raise HTTPException(404, "会话不存在或已过期,请重新上传")
    return session


def _page_info(pdf_path: Path) -> tuple[float, float, int, bool, float, float]:
    with pikepdf.open(pdf_path) as pdf:
        page = pdf.pages[0]
        sizes = [page_geometry(item)[:2] for item in pdf.pages]
        w_mm, h_mm = (value / PT_PER_MM for value in sizes[0])
        min_width = min(width for width, _ in sizes) / PT_PER_MM
        min_side = min(min(size) for size in sizes) / PT_PER_MM
        return w_mm, h_mm, len(pdf.pages), find_cut_cs(page) is not None, min_width, min_side


def _render_preview(pdf_path: Path, out_png: Path) -> None:
    doc = pymupdf.open(pdf_path)
    try:
        page = doc[0]
        rect = page.rect
        scale = min(PREVIEW_MAX_PX / rect.width, PREVIEW_MAX_PX / rect.height, 2.0)
        pix = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
        pix.save(out_png)
    finally:
        doc.close()


def _cut_filename(src_name: str) -> str:
    return f"{_safe_stem(src_name)}_cut.pdf"


class CutBody(BaseModel):
    mode: str = Field(..., description="equal | custom | circle")
    n: int | None = None
    segments: str | None = None
    add_border: bool = False
    inset_mm: float | None = None
    bleed_mm: float | None = None


# ---------- routes ----------

@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return (STATIC / "index.html").read_text(encoding="utf-8")


@app.get("/api/health")
def health() -> dict:
    return {"ok": True}


@app.post("/api/sessions")
async def create_session(file: UploadFile = File(...)) -> dict:
    orig = unicodedata.normalize("NFC", file.filename or "file.pdf")
    ext = Path(orig).suffix.lower()
    if ext not in ALLOWED_EXT:
        raise HTTPException(400, "只接受 PDF / PNG / JPG / TIFF")

    raw = await file.read()
    if not raw:
        raise HTTPException(400, "空文件")
    if len(raw) > MAX_UPLOAD:
        raise HTTPException(400, "文件超过 80MB")

    sid = uuid.uuid4().hex[:12]
    day = datetime.now().strftime("%Y-%m-%d")
    sdir = DATA_DIR / day / sid
    sdir.mkdir(parents=True, exist_ok=True)
    upload_path = sdir / f"upload{ext}"
    upload_path.write_bytes(raw)

    pdf_path = sdir / "input.pdf"
    try:
        if is_image(str(upload_path)):
            w_mm, h_mm = read_image_size_mm(str(upload_path))
            tmp = image_to_pdf(str(upload_path), w_mm, h_mm)
            shutil.move(tmp, pdf_path)
        else:
            shutil.move(str(upload_path), pdf_path)
        w_mm, h_mm, page_count, first_page_has_cut, min_page_w_mm, min_page_side_mm = _page_info(pdf_path)
        preview_path = sdir / "preview.png"
        _render_preview(pdf_path, preview_path)
    except HTTPException:
        shutil.rmtree(sdir, ignore_errors=True)
        raise
    except Exception as exc:  # noqa: BLE001
        shutil.rmtree(sdir, ignore_errors=True)
        raise HTTPException(400, str(exc)) from exc

    session = {
        "id": sid,
        "src_name": orig,
        "pdf_path": str(pdf_path),
        "preview_path": str(preview_path),
        "w_mm": w_mm,
        "h_mm": h_mm,
        "min_page_w_mm": min_page_w_mm,
        "min_page_side_mm": min_page_side_mm,
        "page_count": page_count,
        "created_at": time.time(),
    }
    with _lock:
        _sessions[sid] = session

    return {
        "id": sid,
        "name": orig,
        "w_mm": w_mm,
        "h_mm": h_mm,
        "min_page_w_mm": min_page_w_mm,
        "min_page_side_mm": min_page_side_mm,
        "page_count": page_count,
        "first_page_has_cut": first_page_has_cut,
        "suggest": "circle" if abs(w_mm - h_mm) < 1.0 else "vertical",
        "preview_url": f"/api/sessions/{sid}/preview",
    }


@app.get("/api/sessions/{session_id}/preview")
def session_preview(session_id: str) -> FileResponse:
    session = _session_or_404(session_id)
    path = Path(session["preview_path"])
    if not path.is_file():
        raise HTTPException(404, "预览不存在,请重新上传")
    return FileResponse(path, media_type="image/png")


@app.post("/api/sessions/{session_id}/cut")
def session_cut(session_id: str, body: CutBody) -> FileResponse:
    session = _session_or_404(session_id)
    pdf_path = Path(session["pdf_path"])
    w_mm = session["w_mm"]
    h_mm = session["h_mm"]
    out_name = _cut_filename(session["src_name"])
    out_path = pdf_path.parent / f"cut-{uuid.uuid4().hex}.pdf"

    try:
        if body.mode in ("equal", "custom"):
            if body.mode == "equal":
                positions = equal_positions(w_mm, body.n or 0)
                if not positions:
                    raise ValueError("请输入份数(至少 2)")
            else:
                _, positions, _ = parse_custom_segments(body.segments or "", session["min_page_w_mm"])
            inset = body.inset_mm if body.add_border else 0.0
            if inset is None or not math.isfinite(inset) or inset < 0:
                raise ValueError("请输入有效的外框出血（至少 0 mm）")
            add_vertical_cuts(
                str(pdf_path), positions, str(out_path), inset_mm=inset, add_border=body.add_border,
                equal_parts=body.n if body.mode == "equal" else None,
            )
        elif body.mode == "circle":
            bleed = body.bleed_mm
            if bleed is None:
                raise ValueError("出血未填写")
            r = min(w_mm, h_mm) / 2 - bleed
            if r <= 0:
                raise ValueError(f"出血 {bleed}mm 过大，圆半径 ≤ 0")
            add_circle_cut(str(pdf_path), bleed, str(out_path))
        else:
            raise ValueError("未知模式")
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from exc
    except HTTPException:
        raise
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(500, f"生成失败: {exc}") from exc

    return FileResponse(
        out_path,
        media_type="application/pdf",
        filename=out_name,
        content_disposition_type="attachment",
    )
