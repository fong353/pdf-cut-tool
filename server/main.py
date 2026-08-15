"""折页机 PDF CUT 切割线 Web 服务(内网,端口 8841)。

单进程 FastAPI。切割逻辑全部委托 pdf_ops(add_vertical_cuts / add_circle_cut),
web 层不重写内容流。会话文件落盘 data/jobs/<日期>/<id>/,保留 1 天。

CUT 注入是毫秒～数秒,生成接口同步返回文件,不抄 GPU 队列。
"""

from __future__ import annotations

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
    image_to_pdf,
    is_image,
    read_image_size_mm,
)

DATA_DIR = REPO_ROOT / "data" / "jobs"
STATIC = Path(__file__).resolve().parent / "static"

MAX_UPLOAD = 80 * 1024 * 1024
KEEP_DAYS = 1
CLEAN_INTERVAL = 3600
PREVIEW_MAX_PX = 1600
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
        if s <= 0:
            raise ValueError("段宽必须大于 0")
        total += s
        cuts.append(total)
    remaining = width_mm - total
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


def _page_size_mm(pdf_path: Path) -> tuple[float, float, int]:
    with pikepdf.open(pdf_path) as pdf:
        n = len(pdf.pages)
        mb = pdf.pages[0].mediabox
        w_mm = (float(mb[2]) - float(mb[0])) / PT_PER_MM
        h_mm = (float(mb[3]) - float(mb[1])) / PT_PER_MM
    return w_mm, h_mm, n


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
            w_mm, h_mm, _ = _page_size_mm(pdf_path)
        w_mm, h_mm, page_count = _page_size_mm(pdf_path)
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
        "page_count": page_count,
        "created_at": time.time(),
    }
    with _lock:
        _sessions[sid] = session

    return {
        "id": sid,
        "name": orig,
        "w_mm": round(w_mm, 3),
        "h_mm": round(h_mm, 3),
        "page_count": page_count,
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
    out_path = Path(session["pdf_path"]).parent / out_name

    try:
        if body.mode == "equal":
            n = int(body.n or 0)
            positions = equal_positions(w_mm, n)
            if not positions:
                raise ValueError("请输入份数(至少 2)")
            inset = body.inset_mm
            if inset is None:
                raise ValueError("外框出血未填写")
            add_vertical_cuts(str(pdf_path), positions, str(out_path), inset_mm=inset)
        elif body.mode == "custom":
            if not body.segments or not body.segments.strip():
                raise ValueError("请输入段宽，如: 200, 200, 479.9")
            _, cuts, _ = parse_custom_segments(body.segments, w_mm)
            inset = body.inset_mm
            if inset is None:
                raise ValueError("外框出血未填写")
            add_vertical_cuts(str(pdf_path), cuts, str(out_path), inset_mm=inset)
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
