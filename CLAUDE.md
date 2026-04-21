# PDF CUT 切割线工具

为折页印刷切割机生成 PDF CUT 路径的桌面工具。

## 项目结构

```
cut_tool.py          # 入口 + tkinter UI（App 类）
pdf_ops.py           # PDF 处理（切割线生成，机器识别验证过，字节级稳定，不要动）
preview.py           # PyMuPDF 渲染首页位图 + Canvas 叠加切割线/外框
PDF切割线工具.spec   # PyInstaller 打包配置
dist/PDF切割线工具.exe  # 打包后的可执行文件
.venv/                # 本地开发 venv（基于 brew python3.14 + python-tk@3.14）
```

## 项目约定

**PDF 页面尺寸 = MediaBox**：进入本工具的 PDF 只有 MediaBox，没有 TrimBox / BleedBox / ArtBox / CropBox。
员工在交付前会把文件改成这样。所有代码直接信任 `page.mediabox` 作为物理纸张尺寸，
不需要 fallback 读其他 box。

## 核心概念

切割机通过识别 PDF 中名为 `CUT` 的 **Separation 专色**路径来确定切割位置。专色定义为：

```
[/Separation /CUT /DeviceCMYK <<tint function: C1=[0,0,0,1]>>]
```

生成的 CUT 路径写入独立的 PDF 内容流，包裹在 `/Layer /MCn BDC ... EMC` 标记内，同时在页面 `Resources/Properties` 里注册 Illustrator 兼容的图层元数据（Title="CUT"）。

## 关键处理逻辑

### `ensure_cut_cs(page)`
检查页面是否已有 `/Separation /CUT` colorspace。有则复用（返回 `is_new=False`），无则新建 `/CS_CUT`（返回 `is_new=True`）。

### `add_vertical_cuts(pdf_path, positions_mm, out_path, inset_mm=0.3)`
主函数。流程：
1. 检测/创建 CUT 专色
2. 在 Properties 注册新 CUT 图层（`/MC1`, `/MC2`...，自动避开已有项）
3. 生成 CUT 内容流（外框 + 竖线）
4. 删除页面 `/PieceInfo`（Illustrator 私有元数据，若保留则 Illustrator 会忽略我们写入的内容流）
5. 将新流追加到 Contents 数组，**原始流完全不动**

### `inset_mm`（出血收缩）
外框矩形向内收缩，仅作用于 `is_new=True` 时新建的边框。竖切线始终延伸全页高度。

## UI 模式

**等分模式**：输入份数 N，自动在 `width/N * i`（i=1..N-1）处切割。

**自定义模式**：输入段宽（mm，逗号分隔），如 `200, 200, 479.9`。
- 计算累计切割位置
- 显示剩余量；剩余 < 100mm 红字警告

## 兼容性说明

- **RIP / Acrobat**：直接读 PDF 内容流，始终可见 CUT 路径
- **Illustrator**：若原文件为 Illustrator 保存（含 `/PieceInfo`），必须删除该字段，否则 Illustrator 读自有元数据而忽略内容流改动

## 依赖

```
pikepdf       # PDF 读写（切割线注入）
pymupdf       # 预览：渲染 PDF 首页为位图
Pillow        # 预览：位图转 tkinter 可显示格式
```

tkinter 为 Python 标准库，但 macOS 需要可用的 Tcl/Tk：
`brew install python@3.14 python-tk@3.14`（Apple 自带 python3.9 的 Tk 在 macOS 15+ 上崩溃）。

## 本地运行（macOS 开发）

```bash
/opt/homebrew/bin/python3.14 -m venv .venv
source .venv/bin/activate
pip install pikepdf pymupdf Pillow
python cut_tool.py
```

## 打包 EXE

```bash
python -m PyInstaller --onefile --windowed --name "PDF切割线工具" \
    --collect-all pikepdf --collect-all lxml --collect-all pymupdf cut_tool.py
```

输出：`dist/PDF切割线工具.exe`（加 pymupdf 后约 55MB，单文件，无需 Python 环境）
