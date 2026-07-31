# slide2pptx — AI 幻灯片图像到可编辑 PPTX（MVP）

把一张 AI 生成的 16:9 幻灯片图像（PNG/JPG）反向重建为可在 PowerPoint 中二次编辑的 `.pptx`，并附带一份原图 vs. 重建效果的对比报告。

## 这是什么、不是什么

**能交付的**

- 接近 1:1 的视觉保真度：复杂视觉保留在清理后的背景层，识别出的文本以原生对象覆盖；简单页面通常能取得较高相似度。
- **可编辑的 OCR 文本**：所有识别出的文字块以原生 `TextBox` 写入，字体、字号、颜色尽量逼近原图，可直接在 PowerPoint 中修改。
- **复杂视觉（图表、插画、艺术字、纹理背景）以光栅形式保留在背景层**：图像被作为底图嵌入，不会被错误地"猜"成可编辑形状。
- 渲染对比报告：HTML 报告含原图、重建图、差异热图、SSIM、像素差、可编辑元素清单。

**诚实说明**

- **不是魔术般的原始 PPT 还原**：原 PPT 的母版、主题、占位符逻辑、动画、过渡效果、嵌入对象在栅格化时已经丢失，本工具无法恢复。
- **不是 100% 字符无差错**：艺术字、低分辨率小字、特殊符号可能识别错误；置信度低的文本块会标注 `low_confidence`，最坏情况以光栅回退保留。
- **不是像素级一致**：在不做全矢量重建的前提下，整体 SSIM 目标 0.85+，复杂版式可能在 0.7~0.85 之间。
- **不是商用级字体还原**：仅能猜测 family（如微软雅黑、Calibri），字重、字距、连字无法恢复；不嵌入商业字体以避免授权风险。
- **不做**：多页批量、SmartArt、动画、超链接、母版继承、嵌入 Excel/公式、视频/音频。

## 架构一览

```
[ 输入图像 PNG/JPG ]
       │
       ▼
(1) detect (Python)            ── OCR + 文本样式估计 + 背景清理 → detected.json
       │
       ▼
(2) build  (Node, artifact-tool) ── 按 detected.json 写 PPTX (native + 光栅背景)
       │
       ▼
(3) render (PowerPoint COM)    ── 把 PPTX 渲成 PNG
       │
       ▼
(4) report (Python)            ── SSIM/像素差/可编辑清单 → HTML 报告
```

- **detect 模块**：Python，依赖 RapidOCR（ONNX 推理）和 OpenCV；当前 MVP 原生重建文本，复杂图形保留在背景中。
- **build 模块**：Node 端，使用 `@oai/artifact-tool` 把 detected.json 编译成 PPTX，复杂视觉以背景图形式嵌入。
- **render 模块**：Python 通过 PowerShell 调用本地 PowerPoint COM，不依赖 `pywin32`。
- **report 模块**：Python 使用 Pillow + NumPy 计算 MAE、RMSE、全局近似 SSIM，并用 Jinja2 渲染 HTML。

详细分层与决策门控见 [research/architecture.md](research/architecture.md)。

## 先决条件

| 项目 | 要求 | 说明 |
|---|---|---|
| 操作系统 | **Windows 10/11** | render 步骤依赖 PowerPoint COM，仅 Windows 可用 |
| 渲染端 | **Microsoft PowerPoint（Office 2016+）** | 已安装并能在当前用户下打开文件 |
| Python | **3.10+** | detect 与 report 模块 |
| Node.js | Codex 捆绑运行时或兼容 Node 20+ | build 模块还需要可用的 `@oai/artifact-tool` |
| 磁盘 | 至少 500 MB 可用空间 | OCR、OpenCV 与运行产物 |

**Python 依赖**（已分文件管理，按需安装）：

```text
# detect
Pillow>=10.0
numpy>=1.26
rapidocr_onnxruntime>=1.3   # OCR；缺失时跳过文本识别
opencv-python>=4.8          # 背景清理；缺失时使用原始背景

# report
Pillow>=10.0
numpy>=1.26
Jinja2>=3.1

# test
pytest>=7.4
```

**Node 依赖**：当前工作区已经由 Codex 演示文稿运行时初始化，`artifact-runtime/node_modules/@oai/artifact-tool` 指向本机捆绑包。它不是普通公开 npm 依赖；复制到另一台机器前，需要用 Codex 的演示文稿工作流重新初始化，或提供有权限的等价运行时。

## 安装

所有命令在 PowerShell 中、**项目根目录**下执行：

```powershell
# 1. 克隆 / 进入项目目录
cd "C:\Users\tangvx\Documents\AI ppt"

# 2. 创建并激活 Python 虚拟环境
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 3. 安装 detect 模块依赖
pip install -r python\requirements-detect.txt

# 4. 安装 report 模块依赖
pip install -r python\requirements-report.txt

# 5. 当前 Codex 工作区已经初始化 artifact-runtime；无需 npm install
```

> 单独安装 Node 并不等于拥有 `@oai/artifact-tool`。本 MVP 当前优先作为 Codex 本机原型运行。

## 一键端到端使用

把一张 16:9 的 PNG/JPG 一键转换为可编辑 PPTX + 渲染对比报告：

```powershell
$env:PYTHONPATH = "$PWD\python"
python -m slide2pptx.pipeline_cli samples\source.png --out outputs\sample
```

完成后打开 `outputs\sample\report\report.html` 查看对比报告，在 PowerPoint 中打开 `outputs\sample\build\reconstructed.pptx` 进行二次编辑。只需要 PPTX、不需要 PowerPoint 渲染和报告时，加 `--skip-report`。

## 分步命令参考

### detect：图像分析

```powershell
python python\slide2pptx\detect_cli.py <INPUT_IMAGE> --out <OUT_DIR>
python python\slide2pptx\detect_cli.py --self-test   # 内置合成图自检
```

输入校验失败时返回非零退出码（10 = 输入不存在 / 非图像；20 = OCR 引擎初始化失败）。

### build：JSON → PPTX

```powershell
cd artifact-runtime
node src\convert.mjs --spec <DETECTED_JSON> --out <OUTPUT_PPTX> --preview <PREVIEW_PNG>
node src\convert.mjs --self-test
```

### report：渲染 / 差异 / 报告

```powershell
python -m slide2pptx.report_cli render --pptx slide.pptx --out slide_rendered.png [--slide-index 1] [--timeout 90]
python -m slide2pptx.report_cli diff   --source source.png --rendered rendered.png --out <OUT_DIR> [--threshold 30]
python -m slide2pptx.report_cli full   --source source.png --rendered rendered.png `
                                       --pptx slide.pptx --detected detected.json `
                                       --out <OUT_DIR>
```

退出码约定：0 成功；10 输入缺失；40 渲染失败（仅报告受限，PPTX 仍可用）；99 未捕获异常。

## 输出目录结构

一键命令以 `outputs\sample\` 为例：

```
outputs\sample\
├── detect\
│   ├── detected.json
│   ├── original-background.png
│   └── cleaned-background.png
├── build\
│   ├── reconstructed.pptx
│   └── artifact-preview.png
└── report\
    ├── powerPoint-render.png
    ├── diff.png
    └── report.html
```

`detected.json` 顶层字段（v1.0，以 [spec/detected.schema.json](spec/detected.schema.json) 为准）：

- `source`：原图路径与像素尺寸
- `slide`：目标画布尺寸，当前统一为 1280×720 像素坐标
- `background`：`original` / `cleaned` / `solid` 及背景图路径
- `elements[]`：`text` / `shape` / `image` 元素；当前检测器主要产出原生文本
- `render_strategy`：`native` / `image` / `background`

## 验收指标

| 类别 | 指标 | MVP 目标 | 样本（自检） |
|---|---|---|---|
| 文本 | OCR 字符准确率 | ≥ 85% | — |
| 文本 | 文本块召回率 | ≥ 75% | — |
| 形状 | 拟合 IoU | 后续阶段 | — |
| 整体视觉 | **SSIM** | ≥ 0.85 | **0.959** |
| 整体视觉 | **像素差比例** | ≤ 5% | **1.57%** |
| 可编辑率 | 可编辑元素占比 | ≥ 60% | — |
| 端到端 | 单页处理时延（CPU） | ≤ 30 s | — |

> 样本指标来自 `samples/source.png` 的完整检测、构建、PowerPoint 渲染和对比流程，不代表真实业务图像的极限水平。SSIM 在简单版式（少装饰 + 清晰文字）可达 0.95+；复杂版式、渐变背景、艺术字可能回落到 0.7~0.85。

## 故障排除

| 症状 | 可能原因 | 解决 |
|---|---|---|
| `detect_cli` 报告"OCR 引擎初始化失败" | `rapidocr_onnxruntime` 未装或首次下载失败 | `pip install rapidocr_onnxruntime`；检查网络/代理；模型缓存于 `%USERPROFILE%\.cache\rapidocr` |
| 报告里没有 `text` 元素 | OCR 依赖缺失 | 装回 `rapidocr_onnxruntime` 后重跑 detect |
| `background` 始终是原始图（没有 `cleaned-background.png`） | `opencv-python` 未装 | `pip install opencv-python`，或接受"原始背景"输出 |
| `render` 子命令失败 | 非 Windows、PowerShell 不可用或 PowerPoint 未安装 | 确认 Windows PowerShell 与 Microsoft PowerPoint 均可正常启动 |
| `render` 报 `PowerPoint 无法启动 COM` | Office 安装异常 / 用户权限 / WPS 替代 | 用 `Get-Process POWERPNT` 验证；卸载重装 Office；不要用 WPS 替代 |
| `node` 命令不存在 | 未装 Node 20 LTS | 安装 Node 20 LTS 或 `fnm install 20` |
| PPTX 中字体显示为默认字体而非原图字体 | 目标机器未装相应字体 | PPTX 只记录字体名；用 OFL 字体（思源黑体等）或确认 PowerPoint 字体替换正常 |
| 渲染对比图差异大 | LibreOffice 与 PowerPoint 渲染差异（不在本 MVP 中使用） / 输入是非 16:9 | 仅使用 PowerPoint COM 渲染；将输入裁剪到 16:9 |
| `report_cli full` 退出码 40 | 渲染失败，但 PPTX 仍可独立使用 | 检查 `powerpoint-render.png` 是否存在；用 `--rendered` 指向一个已知好的 PNG 单独跑报告 |

## 测试

```powershell
# Python 全部测试
pytest tests -v

# 仅报告模块
pytest tests\report -v

# Node 构建器冒烟
cd artifact-runtime
node src\convert.mjs --self-test
```

## 路线图（shape / icon / vector 重建）

当前 MVP 尚未自动拆解形状/图标；复杂图标、艺术字、矢量路径统一保留在背景层。后续路线图：

1. **图标库匹配**：CLIP 零样本分类 top-K → 命中 Nerd Font / Material Symbols / emoji 时直接以字符形式嵌入。
2. **几何形状拟合**：从 `findContours` + `approxPolyDP` 升级到最小二乘椭圆/直线 + 头部三角检测，专门处理箭头与多段折线。
3. **自由矢量路径**：尝试 `potrace` 把位图反光栅化为 SVG 路径，再嵌入 PPTX（复杂度高，**仅对图标级细节启用**）。
4. **图表识别**：检测轴、刻度、柱形边界，反推数值表 → 真正的 PowerPoint 图表对象（非光栅）。
5. **母版/主题恢复**：识别标题占位符区域，复用主题色板与字体配对。
6. **多页批处理**：当前仅单页；后续串联多张幻灯片并保持全局样式一致。

每个阶段的触发条件与替代方案见 [research/toolchain.md](research/toolchain.md)。

## 参考文档

- [research/architecture.md](research/architecture.md) — 端到端管道总览、决策门控、置信度评分
- [research/toolchain.md](research/toolchain.md) — Windows 平台工具链对比与推荐栈
- [research/prototype-blueprint.md](research/prototype-blueprint.md) — MVP 模块边界、CLI 合约、`detected.json` Schema

## 许可与免责

- 本 MVP 仅做"单页 16:9 幻灯片图像 → 可编辑 PPTX"的方向性验证，不承诺商用级保真度。
- 输出 PPTX 不嵌入任何商业字体；用户需自行确保 PowerPoint 中已安装所需字体或接受字体替换。
- 不要把受版权保护的截图/设计稿未经授权地转换或分发。
