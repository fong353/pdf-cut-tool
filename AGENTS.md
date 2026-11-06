# PDF CUT 切割线工具

为折页印刷切割机生成 PDF CUT 路径。员工入口是 **mac-company 上的 Web**；`pdf_ops.py` 是切割逻辑真相源。

## 员工入口

- 地址：`http://192.168.0.110:8841`（公司 LAN）/ `http://100.64.0.7:8841`（Tailnet）
- 独立进程，**不挂 ru-dan-web、不改入单系统**
- launchd `com.dayin.pdf-cut-tool`，KeepAlive
- 日常更新：`cd ~/projects/deploy-kit && ./deploy.sh pdf-cut-tool`

## 项目结构

```
server/main.py              # FastAPI 薄封装（上传 / 预览 / 生成）
server/static/index.html    # 单页 UI
pdf_ops.py                  # PDF 处理（切割线生成，机器识别验证过，字节级稳定，不要动）
cut_tool.py / preview.py    # 旧 tkinter 桌面，本机兜底，不是员工入口
deploy/com.dayin.pdf-cut-tool.plist
pyproject.toml              # uv 依赖（含 FastAPI）
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

（代码里 tint 实际是肉眼红 `C1=[0,1,1,0]`，机器认专色名。）

生成的 CUT 路径写入独立的 PDF 内容流，包裹在 `/Layer /MCn BDC ... EMC` 标记内，同时在页面 `Resources/Properties` 里注册 Illustrator 兼容的图层元数据（Title="CUT"）。

## 关键处理逻辑

### `ensure_cut_cs(page)`
检查页面是否已有 `/Separation /CUT` colorspace。有则复用（返回 `is_new=False`），无则新建 `/CS_CUT`（返回 `is_new=True`）。

### `add_vertical_cuts(pdf_path, positions_mm, out_path, inset_mm=0.3, *, add_border=True)`
主函数。流程：
1. 检测/创建 CUT 专色
2. 在 Properties 注册新 CUT 图层（`/MC1`, `/MC2`...，自动避开已有项）
3. 生成竖切线。仅在 `add_border=True` 且当前页原本没有 CUT 专色时新增外框。
4. 删除页面 `/PieceInfo`（Illustrator 私有元数据，若保留则 Illustrator 会忽略我们写入的内容流）
5. 将新流追加到 Contents 数组，**原始流完全不动**

### `inset_mm`（出血收缩）
外框矩形向内收缩。竖切线始终延伸全页高度。

## UI 模式

**竖切折页**默认只加竖切线。勾选「生成外轮廓」后显示「外框出血」，初始值为每边 0.3 mm。

**等分模式**：输入份数 N，自动在 `width/N * i`（i=1..N-1）处切割。

**自定义模式**：每段宽度（mm）填在独立输入框，使用「＋ 添加一段」和「−」增删。
- 计算累计切割位置
- 显示剩余量；剩余 < 100mm 红字警告

**圆切贴纸**：圆心=页中心，半径 = min(宽,高)/2 − 出血 mm。方图（|w−h|&lt;1mm）上传后自动切到此 tab。

## 兼容性说明

- **RIP / Acrobat**：直接读 PDF 内容流，始终可见 CUT 路径
- **Illustrator**：若原文件为 Illustrator 保存（含 `/PieceInfo`），必须删除该字段，否则 Illustrator 读自有元数据而忽略内容流改动

## 依赖

```
pikepdf / pymupdf / Pillow   # 切割 + 预览位图
fastapi / uvicorn            # Web
```

## 本地运行（macOS 开发）

```bash
cd ~/projects/pdf-cut-tool
uv sync
uv run uvicorn server.main:app --host 127.0.0.1 --port 18841
```

旧桌面：`uv run python cut_tool.py`（需要可用的 Tcl/Tk）。

## 部署

目标机代码路径 `~/pdf-cut-tool`（ASCII）。首装见 deploy-kit `docs/onboard-new-tool.md`。
推代码前必须先 commit（bundle 只含已提交内容）。

---

## Cursor 记忆迁移（2026-08-06）

已放弃 Claude Code。原 `~/.claude/projects/*/memory` 已备份到 `~/Documents/claude-memory-backup-20260806/`，并迁入：

- `.cursor/rules/claude-memory-*.mdc` — 常驻/按需规则
- `.cursor/skills/claude-memory/` — 完整记忆正文（`references/`；凭据在 `private/` 且 gitignore）

本 `AGENTS.md` 仍由 Cursor 读取；细节踩坑优先查 skill。
