# 本地 MVP 实现蓝图：单张 16:9 幻灯片图片 → 可编辑 PPTX + 渲染对比报告

> 文档版本：v0.1（原型设计阶段）
> 目标读者：后续实现工程师、可独立委派的子任务负责人
> 工作目录：`C:\Users\tangvx\Documents\AI ppt`
> 平台假设：Windows 11 + PowerShell 5.1；Python 3.10+；所有命令均为非破坏性（不删除 `C:\` 系统目录、不修改注册表、不联网安装未授权服务）

---

## 1. 目标与边界

### 1.1 目标
输入一张 **16:9** 的 PNG/JPG 幻灯片图片（用户制作的成品截图或导出版），自动重建为一个**可二次编辑**的 PPTX，并产出一份**渲染对比报告**（原图 vs. 重建后幻灯片的差异并排图 + 指标）。

### 1.2 非目标（MVP 暂不做）
- 多张幻灯片批量处理（仅 1 张）
- 复杂动画 / 母版 / SmartArt / 嵌入的 Excel 对象 / 公式
- 复杂图表（柱状图、折线图）→ 形状近似即可，不做数据反推
- 视频、音频
- 任何形式的云端调用（全部本地）
- 中文以外语种字体的精确匹配（拉丁字母覆盖；CJK 字体回退到系统默认）

### 1.3 成功标准（MVP 级）
- 单张 1920×1080 图像，端到端处理时间 < 30 秒（CPU，无 GPU 加速）
- 输出 PPTX 在 Microsoft PowerPoint 2016+ 中可打开，文本可编辑、形状可选中移动
- 文本召回率（Recall@word）≥ 75%，准确率（Precision@word）≥ 85%（基于自带 6 张固件）
- 渲染对比报告：HTML 形式，含原图、再构图、像素差异热力图、可编辑性 checklist

---

## 2. 技术选型

| 关注点 | 选型 | 理由 |
|---|---|---|
| 语言 | Python 3.10+ | 库生态最完整；Windows 友好；脚本化便于委派 |
| 图像读取 | Pillow | 读取 PNG/JPG，EXIF 忽略 |
| 文本检测 / OCR | PaddleOCR（ONNX 推理版） | 中文识别强；可纯 CPU |
| 形状 / 矩形检测 | OpenCV 4.x | `findContours` + 多边形近似 |
| 颜色采样 | Pillow + scikit-learn（KMeans 仅在需要时） | 取区域代表色 |
| 字体估算 | Pillow 测量 + 启发式字号映射 | 不依赖深度模型 |
| PPTX 生成 | python-pptx 0.6.21+ | 唯一成熟稳定的本地 PPTX 库 |
| 报告渲染 | Jinja2 + WeasyPrint（PDF）/ Playwright（HTML） | 跨平台；不依赖 Office |
| 测试 | pytest + pytest-cov | 标配 |
| CLI | Click 8.x | 简单清晰；自动生成 help |
| 配置 | PyYAML | `config.yaml` 存阈值、字体映射等 |

> **所有依赖都通过 `pip install -r requirements.txt` 安装；脚本不调用 `os.remove` 以外的系统命令。**

---

## 3. 项目目录结构

```
AI ppt/
├── README.md                          # 快速开始（≤ 1 页）
├── requirements.txt                   # 锁定版本
├── pyproject.toml                     # 包元数据（可选）
├── config.yaml                        # 阈值、字体映射、颜色表
├── .gitignore                         # 忽略 __pycache__、outputs/、venv/
│
├── src/
│   └── slide2pptx/
│       ├── __init__.py
│       ├── cli.py                     # 入口：Click 命令
│       ├── pipeline.py                # 顶层编排：串联 detect → assemble → render
│       │
│       ├── detect/                    # ★ 子模块 1：图像理解
│       │   ├── __init__.py
│       │   ├── text_detector.py       # OCR：文字块 + bbox + 置信度
│       │   ├── shape_detector.py      # 矩形/圆角矩形/椭圆/线段
│       │   ├── image_region.py        # 图片型区域（连通域+边缘密度）
│       │   ├── background.py          # 背景色 / 渐变 / 单色判定
│       │   └── layout_infer.py        # 从元素推断占位语义（标题/正文/页眉...）
│       │
│       ├── assemble/                  # ★ 子模块 2：PPT 组装
│       │   ├── __init__.py
│       │   ├── pptx_builder.py        # python-pptx 封装
│       │   ├── style_resolver.py      # 颜色→RGB、字体→名称、字号
│       │   ├── element_writer.py      # 文本框 / 形状 / 图片 的写入
│       │   └── fallback.py            # ★ 可编辑回退策略
│       │
│       ├── report/                    # ★ 子模块 3：对比报告
│       │   ├── __init__.py
│       │   ├── renderer.py            # PPTX → PNG（libreoffice headless 或 fallback）
│       │   ├── diff.py                # 原图 vs 渲染图：像素差 + SSIM
│       │   ├── checklist.py           # 可编辑性自动检查
│       │   └── html_builder.py        # Jinja2 → HTML
│       │
│       └── common/                    # 共享工具
│           ├── __init__.py
│           ├── io.py                  # 路径、JSON 读写
│           ├── geometry.py            # bbox 工具
│           ├── color.py               # 颜色命名 / 转换
│           └── logging.py             # 统一日志
│
├── tests/
│   ├── conftest.py                    # pytest 固件
│   ├── fixtures/                      # 测试样本
│   │   ├── slides/                    # 6 张 1920x1080 固件
│   │   │   ├── 01_text_only.png
│   │   │   ├── 02_title_body.png
│   │   │   ├── 03_two_column.png
│   │   │   ├── 04_with_shape.png
│   │   │   ├── 05_with_image.png
│   │   │   └── 06_complex.png
│   │   └── expected/                  # 黄金 JSON / PPTX 校验
│   ├── test_detect/
│   │   ├── test_text_detector.py
│   │   ├── test_shape_detector.py
│   │   └── test_layout_infer.py
│   ├── test_assemble/
│   │   ├── test_pptx_builder.py
│   │   └── test_fallback.py
│   ├── test_report/
│   │   └── test_diff.py
│   └── test_pipeline.py                # 端到端冒烟
│
├── outputs/                           # 运行时产物（不入 Git）
│   └── <job_id>/
│       ├── detected.json
│       ├── slide.pptx
│       ├── slide_rendered.png
│       ├── diff.png
│       └── report.html
│
└── research/                          # ★ 本文档所在目录
    └── prototype-blueprint.md
```

---

## 4. 模块边界与接口契约

> 原则：**每个子模块只依赖 `common/` 和标准库**；子模块之间不互相 import；编排逻辑集中在 `pipeline.py`。

### 4.1 `detect/` 子模块

| 模块 | 输入 | 输出 | 关键接口 | 依赖 |
|---|---|---|---|---|
| `text_detector.TextDetector.detect(image_path) -> list[TextBlock]` | 图像路径 | `TextBlock` 列表 | `TextBlock(bbox, text, confidence, language)` | PaddleOCR |
| `shape_detector.ShapeDetector.detect(image_path) -> list[Shape]` | 图像路径 | `Shape` 列表 | `Shape(bbox, type, fill_color, stroke_color, stroke_width)` | OpenCV |
| `image_region.ImageRegionDetector.detect(image_path) -> list[ImageRegion]` | 图像路径 | `ImageRegion` 列表 | `ImageRegion(bbox, source_image_bytes)` | OpenCV、Pillow |
| `background.BackgroundAnalyzer.analyze(image_path) -> Background` | 图像路径 | `Background` | `Background(type, primary_color, gradient_stops?)` | Pillow |
| `layout_infer.LayoutInferer.infer(elements) -> list[Element]` | 全部检测元素 | 语义元素列表 | `Element(role, bbox, payload, editable_score)` | 仅 `common/` |

### 4.2 `assemble/` 子模块

| 模块 | 输入 | 输出 | 关键接口 | 依赖 |
|---|---|---|---|---|
| `pptx_builder.PptxBuilder.build(elements, slide_size) -> bytes` | 元素列表 | PPTX 字节流 | 同上 | python-pptx |
| `style_resolver.StyleResolver.resolve(visual_hint) -> Style` | 视觉提示 | `Style` | `Style(font_name, font_size, color_rgb, bold, italic, align)` | `common/color` |
| `element_writer.ElementWriter` | 元素 + Style | 直接写 PPTX 内部 API | 一组 `_write_textbox / _write_shape / _write_image` 私有方法 | python-pptx |
| `fallback.FallbackDecider.decide(element) -> RenderPlan` | 元素 | 渲染计划 | `RenderPlan(strategy, params)` | 仅 `common/` |

### 4.3 `report/` 子模块

| 模块 | 输入 | 输出 | 关键接口 | 依赖 |
|---|---|---|---|---|
| `renderer.SlideRenderer.render(pptx_bytes, out_png) -> Path` | PPTX 字节 | PNG 路径 | 同名 | LibreOffice headless（可选）/ Pillow fallback |
| `diff.DiffCalculator.compute(orig_png, rendered_png) -> DiffResult` | 两张 PNG | `DiffResult(heatmap_path, ssim_score, pixel_diff_ratio)` | 同名 | scikit-image |
| `checklist.EditabilityChecker.check(pptx_path) -> list[CheckItem]` | PPTX 路径 | 检查项列表 | `CheckItem(name, pass_, detail)` | python-pptx |
| `html_builder.HtmlBuilder.build(context, out_html) -> Path` | 上下文 | HTML 路径 | 同名 | Jinja2 |

### 4.4 `pipeline.py` 编排

```
run(image_path: Path, output_dir: Path, config: Config) -> JobResult
```

只做：调用各子模块、记录日志、聚合 `JobResult`。**不写任何检测 / 重建 / 报告逻辑。**

---

## 5. CLI 合约

### 5.1 命令

```powershell
python -m slide2pptx.cli convert <IMAGE> --out <DIR> [--config config.yaml] [--no-report] [--verbose]
python -m slide2pptx.cli detect   <IMAGE> --out <DIR>           # 仅跑检测，输出 detected.json
python -m slide2pptx.cli build    <DETECTED_JSON> --out <DIR>   # 从 JSON 重建 PPTX
python -m slide2pptx.cli report   <PPTX> <ORIG_IMAGE> --out <DIR>
python -m slide2pptx.cli --version
python -m slide2pptx.cli --help
```

### 5.2 参数

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `IMAGE` | 路径 | 是（convert/detect） | — | 16:9 的 PNG/JPG；建议 1920×1080 |
| `--out` | 目录 | 是 | — | 产物输出目录；不存在则自动创建 |
| `--config` | 文件 | 否 | `config.yaml` | 阈值、字体映射 |
| `--no-report` | 标志 | 否 | False | 仅生成 PPTX，不渲染报告 |
| `--verbose` | 标志 | 否 | False | 打印 DEBUG 日志 |

### 5.3 退出码

| 码 | 含义 |
|---|---|
| 0 | 成功 |
| 10 | 输入文件不存在或不是 PNG/JPG |
| 11 | 图像宽高比不是 16:9（容差 ±1.5%） |
| 12 | 图像分辨率过低（< 800px 宽） |
| 20 | OCR 引擎初始化失败 |
| 30 | PPTX 写入失败（磁盘满 / 路径只读） |
| 40 | 报告渲染失败（非致命：仅退出码非 0，PPTX 仍可用） |
| 99 | 未捕获异常 |

### 5.4 输入校验

```python
def validate_image(path: Path) -> None:
    """检查文件存在、是 PNG/JPG、宽高比 16:9（容差 0.015）、最小宽度 800。"""
```

> 校验失败抛出 `InputError`，CLI 捕获后输出人类可读消息并以对应退出码退出。

---

## 6. 中间 JSON Schema（`detected.json`）

> **这是整个系统的核心契约**。所有子模块围绕它解耦；委派实现时按此交付。

```json
{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "$id": "slide2pptx.detected.v1",
  "title": "Slide2PPTX Detected Elements",
  "type": "object",
  "required": ["version", "slide_size", "background", "elements"],
  "properties": {
    "version": { "const": "1.0" },
    "slide_size": {
      "type": "object",
      "required": ["width_px", "height_px", "width_emu", "height_emu"],
      "properties": {
        "width_px":  { "type": "integer", "minimum": 800 },
        "height_px": { "type": "integer", "minimum": 450 },
        "width_emu": { "type": "integer", "description": "1 px = 9525 EMU at 96 DPI" },
        "height_emu":{ "type": "integer" }
      }
    },
    "background": {
      "oneOf": [
        { "type": "object", "required": ["type", "color_hex"],
          "properties": { "type": { "const": "solid" },
                          "color_hex": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" } } },
        { "type": "object", "required": ["type", "stops"],
          "properties": { "type": { "const": "gradient" },
                          "stops": { "type": "array", "minItems": 2, "maxItems": 3,
                                     "items": { "type": "object",
                                                "required": ["position", "color_hex"],
                                                "properties": {
                                                  "position": { "type": "number", "minimum": 0, "maximum": 1 },
                                                  "color_hex": { "type": "string" } } } } } }
      ]
    },
    "elements": {
      "type": "array",
      "items": { "$ref": "#/$defs/Element" }
    }
  },
  "$defs": {
    "BBox": {
      "type": "object",
      "required": ["x", "y", "w", "h"],
      "properties": {
        "x": { "type": "integer", "minimum": 0 },
        "y": { "type": "integer", "minimum": 0 },
        "w": { "type": "integer", "minimum": 1 },
        "h": { "type": "integer", "minimum": 1 }
      }
    },
    "Element": {
      "type": "object",
      "required": ["id", "role", "bbox", "editable_score", "render_plan", "payload"],
      "properties": {
        "id":        { "type": "string", "pattern": "^el_[0-9]{4}$" },
        "role":      { "enum": ["title", "subtitle", "body", "caption", "footer",
                                "shape_rect", "shape_round_rect", "shape_ellipse",
                                "shape_line", "image", "chart_placeholder", "unknown"] },
        "bbox":      { "$ref": "#/$defs/BBox" },
        "z_order":   { "type": "integer", "minimum": 0, "default": 0 },
        "editable_score": { "type": "number", "minimum": 0, "maximum": 1,
                            "description": "1=完全可编辑，0=完全用位图占位" },
        "render_plan": {
          "type": "object",
          "required": ["strategy"],
          "properties": {
            "strategy": { "enum": ["native_shape", "native_textbox",
                                   "native_image", "bitmap_fallback"] },
            "params":   { "type": "object" }
          }
        },
        "payload": {
          "oneOf": [
            { "$ref": "#/$defs/TextPayload" },
            { "$ref": "#/$defs/ShapePayload" },
            { "$ref": "#/$defs/ImagePayload" }
          ]
        }
      }
    },
    "TextPayload": {
      "type": "object",
      "required": ["text", "style"],
      "properties": {
        "text":  { "type": "string" },
        "style": {
          "type": "object",
          "required": ["font_name", "font_size_pt", "color_hex"],
          "properties": {
            "font_name":    { "type": "string" },
            "font_size_pt": { "type": "number", "minimum": 6, "maximum": 200 },
            "color_hex":    { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" },
            "bold":         { "type": "boolean" },
            "italic":       { "type": "boolean" },
            "align":        { "enum": ["left", "center", "right", "justify"] },
            "line_spacing": { "type": "number", "minimum": 0.5, "maximum": 5 }
          }
        },
        "language":  { "enum": ["zh", "en", "mixed", "other"] },
        "confidence":{ "type": "number", "minimum": 0, "maximum": 1 }
      }
    },
    "ShapePayload": {
      "type": "object",
      "required": ["shape_type"],
      "properties": {
        "shape_type": { "enum": ["rect", "round_rect", "ellipse", "line", "arrow"] },
        "fill":   { "type": ["string", "null"], "description": "十六进制颜色或 null=无填充" },
        "stroke": { "type": ["string", "null"] },
        "stroke_width_px": { "type": "number", "minimum": 0 },
        "corner_radius_px":{ "type": "number", "minimum": 0 }
      }
    },
    "ImagePayload": {
      "type": "object",
      "required": ["source"],
      "properties": {
        "source": {
          "type": "object",
          "required": ["kind", "data"],
          "properties": {
            "kind": { "enum": ["extracted", "base64"] },
            "data": { "type": "string",
                      "description": "若 kind=extracted，指向 outputs/<job>/assets/ 下的相对路径" }
          }
        }
      }
    }
  }
}
```

### 6.1 关键约定
- `bbox` 全部使用**像素**；PPTX 写入时再换算 EMU（`1 px = 9525 EMU`）。
- `z_order` 越大越在上层；同 `z_order` 按数组顺序。
- `editable_score` 用来驱动回退：< 0.4 触发 `bitmap_fallback`（见 §8）。
- `id` 全局唯一，便于日志引用。

---

## 7. 重建算法

### 7.1 顶层流程

```
1. validate_image(image_path)
2. background = BackgroundAnalyzer.analyze(image_path)
3. text_blocks    = TextDetector.detect(image_path)
4. shapes         = ShapeDetector.detect(image_path)
5. image_regions  = ImageRegionDetector.detect(image_path)
6. elements = LayoutInferer.infer(text_blocks, shapes, image_regions, background)
7. for el in elements: el.render_plan = FallbackDecider.decide(el)
8. pptx_bytes = PptxBuilder.build(elements, slide_size)
9. if not --no-report:
     rendered_png = SlideRenderer.render(pptx_bytes)
     diff         = DiffCalculator.compute(image_path, rendered_png)
     checks       = EditabilityChecker.check(pptx_path)
     HtmlBuilder.build({...}, out_html)
10. write detected.json
```

### 7.2 各检测器算法要点

**TextDetector**
- 调 PaddleOCR（`det=True, rec=True, use_angle_cls=True`）
- 过滤 `confidence < 0.5` 的结果
- 相邻文字块若垂直距离 < `0.6 × avg_height` 且水平重叠 > 50% → 合并为同一行
- 同一行内 `gap < 0.3 × char_width` → 合并段（处理中文/英文混排）
- 输出最终 `TextBlock`，记录 `language`（用 Unicode 范围判断 zh/en/mixed）

**ShapeDetector**
- 灰度化 → 自适应阈值（Canny 边缘或 OTSU）
- `findContours` + `approxPolyDP`（epsilon=0.02×perimeter）
- 分类规则：
  - 顶点数=4 且边接近正交 → `rect`
  - 顶点数=4 且圆角明显（凸包 vs 多边形面积比 < 0.92）→ `round_rect`
  - 顶点数≈12 → `ellipse`
  - 长宽比 > 10 且面积小 → `line`
- 颜色采样：取轮廓内 5×5 区域均值，KMeans 量化到调色板（白/黑/红/蓝/绿/黄/灰/自定义 8 色）

**ImageRegionDetector**
- 与 Shape 区域做差集：减去已被形状覆盖的部分
- 剩余区域用**边缘密度**判定是否为图片（自然图像的 Laplacian 方差远高于纯色 / 渐变 / 文字）
- 阈值：`laplacian_var > 80` 且 `unique_color_count > 256` → 视为图片
- 从原图裁切该 bbox 并保存为 `assets/el_XXXX.png`

**BackgroundAnalyzer**
- 采样四角 + 中心共 5 个 20×20 块的代表色
- 若 5 块颜色 max-distance < 10（Lab 空间）→ solid
- 否则沿水平 / 垂直中线采样 11 个点，色相平滑变化 → gradient，记录 stops

**LayoutInferer**
- 启发式：
  - bbox 在上 25% 且字号估计 > 28pt → `title`
  - bbox 在下 85%~95%、字号 < 14pt → `footer`
  - 剩余文本块按从左到右、从上到下编号
- 字号估计：bbox 高度 × 0.75（粗略经验值，后续按字体度量校准）
- 元素之间重叠检测：若 A 完全包含 B，B 是装饰子元素（圆点/图标），标记 `z_order += 1`

### 7.3 PptxBuilder 算法

```
slide = pres.slides.add_slide(pres.slide_layouts[6])  # 空白版式
set_background(slide, background)
sorted_elements = sorted(elements, key=lambda e: -e.z_order)
for el in sorted_elements:
    plan = el.render_plan
    if plan.strategy == "native_textbox": ElementWriter._write_textbox(slide, el)
    elif plan.strategy == "native_shape":   ElementWriter._write_shape(slide, el)
    elif plan.strategy == "native_image":   ElementWriter._write_image(slide, el)
    elif plan.strategy == "bitmap_fallback":ElementWriter._write_bitmap(slide, el, el.bbox)
```

**坐标换算**
- `left_emu   = bbox.x * 9525`
- `top_emu    = bbox.y * 9525`
- `width_emu  = bbox.w * 9525`
- `height_emu = bbox.h * 9525`
- 幻灯片尺寸用 16:9 标准 13.333" × 7.5"（= 12192000 × 6858000 EMU）

**样式解析**（`StyleResolver`）
- `font_name` 来自 `config.yaml` 的 `font_map`（例：黑体 → SimHei；思源黑体 → Microsoft YaHei UI）
- `font_size_pt` = `bbox.h * 0.75 / 96 * 72`（像素→磅）
- 找不到的字体回退到 `Calibri`（英文）或 `Microsoft YaHei`（中文）

---

## 8. 可编辑回退规则

> 目标：**宁可多保留可编辑性，也不为了完美还原牺牲可用性。**

### 8.1 FallbackDecider 决策表

| 条件 | strategy | 备注 |
|---|---|---|
| 文本，OCR confidence ≥ 0.75，bbox 高度 ≥ 18px | `native_textbox` | 正常 |
| 文本，OCR confidence < 0.5 | `bitmap_fallback` | 整块按图片插入（可拖动/缩放，但文字不能改） |
| 文本，OCR confidence 0.5~0.75 | `native_textbox` + 标注 `low_confidence=true` | 仍生成文本框，附 `custom_property` |
| 形状，顶点分类置信度 ≥ 0.8 | `native_shape` | 正常 |
| 形状，置信度 < 0.8 或形状怪异（> 8 顶点） | `bitmap_fallback` | 整块按图片插入 |
| 图片区域，laplacian_var > 80 | `native_image` | 正常 |
| 装饰性元素（图标、logo、点） | `native_shape` 简化 | 单色椭圆或矩形近似 |
| 文字 + 形状叠合（按钮 = 圆角矩形 + 文字） | 合并为单个 `native_shape` + 内嵌文字 | 减少元素数 |
| 元素与背景渐变重叠 | `bitmap_fallback` | 渐变上叠文字很难独立编辑 |

### 8.2 `bitmap_fallback` 实现

```python
def _write_bitmap(self, slide, el, bbox):
    cropped = original_image.crop((bbox.x, bbox.y, bbox.x + bbox.w, bbox.y + bbox.h))
    cropped.save(tmp_path)  # 临时文件
    slide.shapes.add_picture(tmp_path, left_emu, top_emu, width_emu, height_emu)
    # 在 PPTX 备注中记录：el_XXXX = bitmap_fallback
```

> **回退不等于失败**：报告里清楚标注哪些元素是位图占位即可。

### 8.3 编辑性 checklist 自动检查

`EditabilityChecker.check` 返回：

- [x] 文本元素 ≥ 1 且均可编辑 → 通过
- [x] 所有 `bitmap_fallback` 元素都标记到 PPT 备注 → 通过
- [x] 形状可单独选中（未合并到背景位图）→ 通过
- [x] 字体名均为系统可用字体 → 通过
- [x] 文件能在 PowerPoint 中无错误打开 → 通过（用 python-pptx 反序列化校验）

---

## 9. 渲染对比报告

### 9.1 渲染管线

```
PPTX bytes
   ↓
[可选] libreoffice --headless --convert-to png slide.pptx
   ↓
slide_rendered.png  (目标 1920×1080)
   ↓
[若 LibreOffice 不可用] 降级：直接对 detected.json 用 Pillow 复刻一次（仅供视觉参考，不计入 SSIM）
```

> LibreOffice 在 Windows 上是**可选**依赖；安装则用，未安装则报告写明"渲染受限"。

### 9.2 差异计算

```python
from skimage.metrics import structural_similarity as ssim
import numpy as np

a = np.array(orig.convert("RGB"))
b = np.array(rendered.convert("RGB").resize(orig.size))
score, diff_map = ssim(a, b, channel_axis=2, full=True)
heatmap = (1 - diff_map) * 255  # 差异越大越亮
pixel_diff_ratio = np.mean(np.abs(a.astype(int) - b.astype(int)) > 30)
```

### 9.3 HTML 报告结构（Jinja2 模板）

```
┌──────────────────────────────────────────────┐
│  Slide2PPTX Report          Job: <job_id>    │
├──────────────────────────────────────────────┤
│  元信息：源图尺寸、PPTX 字节数、生成耗时     │
├──────────────────────────────────────────────┤
│  ┌────────────┐  ┌────────────┐  ┌─────────┐ │
│  │  原图       │  │  重建渲染   │  │ 差异热图│ │
│  │             │  │             │  │         │ │
│  └────────────┘  └────────────┘  └─────────┘ │
├──────────────────────────────────────────────┤
│  指标：SSIM = 0.83  像素差 = 6.2%            │
├──────────────────────────────────────────────┤
│  元素清单（表格）：                          │
│  id | role | strategy | confidence | bbox    │
├──────────────────────────────────────────────┤
│  编辑性 checklist：✅ / ❌ 列表                │
├──────────────────────────────────────────────┤
│  备注：低置信度元素 / bitmap_fallback 列表    │
└──────────────────────────────────────────────┘
```

输出单一 HTML 文件（内嵌 base64 图片），用户双击即可在浏览器打开。

---

## 10. 测试固件

### 10.1 6 张固件（`tests/fixtures/slides/`）

| 文件 | 内容 | 验证重点 |
|---|---|---|
| `01_text_only.png` | 居中一段英文 + 标题 | OCR + 字号估计 |
| `02_title_body.png` | 大标题 + 短正文 | 角色推断（title/body） |
| `03_two_column.png` | 左右两栏文本 | 多元素排版、z_order |
| `04_with_shape.png` | 文本 + 圆角矩形按钮 | 形状 + 文字合并 |
| `05_with_image.png` | 全屏背景图 + 文字 | 图片区域检测 + 渐变背景 |
| `06_complex.png` | 标题 + 列表 + 形状 + 图片 + 页脚 | 全链路压力 |

> 固件生成方式（**仅开发期一次性**，不进入生产流程）：
> 1. 用 PowerPoint 手画 6 张
> 2. 导出为 1920×1080 PNG
> 3. 提交到 `tests/fixtures/slides/`
> 4. **任何对固件的修改必须更新 `tests/fixtures/expected/` 的黄金 JSON**

### 10.2 验收测试矩阵

| 层级 | 测试 | 工具 | 通过条件 |
|---|---|---|---|
| 单元 | `test_text_detector` | pytest | OCR 召回 ≥ 80% (按字符) |
| 单元 | `test_shape_detector` | pytest | 形状数量与位置 IoU ≥ 0.7 |
| 单元 | `test_layout_infer` | pytest | role 分类准确 ≥ 85% |
| 单元 | `test_pptx_builder` | pytest | 生成文件可被 python-pptx 反读 + 元素数量匹配 |
| 单元 | `test_fallback` | pytest | 给定合成元素，决策符合决策表 |
| 单元 | `test_diff` | pytest | 同一张图 diff ratio < 1%；不同图 > 5% |
| 集成 | `test_pipeline` | pytest | 6 张固件全部跑通，PPTX 字节非空 |
| 端到端 | `test_e2e_report` | pytest | 报告 HTML 包含必备区块 |
| 验收 | 手工 | PowerPoint | 6 张 PPTX 打开无报错，文字可改 |

### 10.3 黄金 JSON

`tests/fixtures/expected/<fixture_name>.json` 存储人工标注的 `detected.json` 预期结构。  
对比规则：
- 元素数量 ±20% 容差
- 文本字段允许字符串编辑距离 ≤ 2
- 形状 / 图片的 bbox IoU ≥ 0.7

---

## 11. 分阶段里程碑

> 每完成一个里程碑产出可演示产物；委派实现时按里程碑拆任务最自然。

### M0 — 项目骨架（0.5 天）
- 创建目录结构
- `requirements.txt`、`.gitignore`、`README.md`（运行方式）
- `cli.py` 仅有 `--help` 和 `--version`
- 验收：`python -m slide2pptx.cli --help` 正常输出

### M1 — 图像输入与校验（0.5 天）
- `common/io.py` + 图像校验
- `cli convert` 接受空实现
- 验收：错误输入返回退出码 10/11/12

### M2 — 检测层（3 天）
- M2.1 文本检测（1 天）
- M2.2 形状检测（1 天）
- M2.3 布局推断 + 背景（1 天）
- 验收：6 张固件全部产出 `detected.json`，黄金对比通过

### M3 — 组装层（2 天）
- M3.1 PptxBuilder 基础（1 天）
- M3.2 Fallback + StyleResolver（1 天）
- 验收：6 张 PPTX 全部生成，python-pptx 反读无异常

### M4 — 报告层（1.5 天）
- M4.1 渲染管线 + diff（1 天）
- M4.2 HTML 报告 + checklist（0.5 天）
- 验收：报告 HTML 在浏览器渲染正确

### M5 — 端到端 + 验收（1 天）
- 流水线串联
- 6 张固件完整跑通
- PowerPoint 手工打开验证可编辑
- 性能基线 < 30s/张
- 验收：完成 §1.3 全部成功标准

**总计：约 8.5 个工作日（单人）**

---

## 12. 委派指南（可拆分给多人的子任务）

| 子任务 | 模块 | 输入文件 | 输出契约 | 估时 |
|---|---|---|---|---|
| T1 | `common/` | 无 | `io.py / geometry.py / color.py / logging.py` + 单测 | 0.5 d |
| T2 | `detect/text_detector` | T1 | `detect(image) -> list[TextBlock]` + 单测 | 1.0 d |
| T3 | `detect/shape_detector` | T1 | `detect(image) -> list[Shape]` + 单测 | 1.0 d |
| T4 | `detect/image_region` + `background` | T1 | 两个 `detect` 函数 + 单测 | 1.0 d |
| T5 | `detect/layout_infer` | T2/T3/T4 产物 | `infer(...) -> list[Element]` + 单测 | 1.0 d |
| T6 | `assemble/pptx_builder` + `style_resolver` | T5 JSON schema | `build(...) -> bytes` + 单测 | 1.0 d |
| T7 | `assemble/fallback` | T6 | `decide(...) -> RenderPlan` + 单测 | 0.5 d |
| T8 | `report/renderer` + `report/diff` | T6 | `render` + `compute` + 单测 | 1.0 d |
| T9 | `report/checklist` + `report/html_builder` | T6/T8 | `check` + `build` + 单测 | 0.5 d |
| T10 | `pipeline` + `cli` | 全部 | 端到端可跑 | 1.0 d |
| T11 | 6 张固件 + 黄金 JSON | 手工 | `tests/fixtures/slides/*.png` + `expected/*.json` | 0.5 d |

> T2~T9 可完全并行（接口已用 §6 JSON schema 锁定）。T10/T11 串行收尾。

---

## 13. 风险与缓解

| 风险 | 概率 | 影响 | 缓解 |
|---|---|---|---|
| PaddleOCR 模型在 Windows 首次运行下载慢 | 中 | M2 延期 | 预下载到 `models/`；提供离线模式 |
| 复杂背景渐变上叠文字检测失败 | 高 | 元素被回退 | 已有 `bitmap_fallback` 兜底；报告标注 |
| LibreOffice headless 不可用 | 中 | 报告无渲染图 | 检测缺失则降级为 Pillow 复刻 + 提示 |
| 字体在用户机器缺失 | 中 | 文字外观变化 | `config.yaml` 的 `font_map` + 自动回退到 `Microsoft YaHei` |
| OCR 召回偏低（< 75%） | 中 | MVP 不达标 | 调低 confidence 阈值（0.5 → 0.4）+ 引入行合并 |
| 检测到的形状并非真实形状（误把文字当成形状） | 中 | 多余元素 | 形状检测后与文字 bbox 做 IoU 过滤（IoU > 0.6 视为文字） |

---

## 14. 非破坏性保证（Windows 友好）

- 所有写操作限定在 `--out` 指定目录及其子目录
- 不修改 `PATH`、不写注册表、不动用户配置
- 临时文件统一放 `%TEMP%\slide2pptx\`，任务结束**可选**清理（默认保留，幂等）
- 不联网拉取任何模型/字体（依赖打包到 `models/` 和 `fonts/`）
- 进程单实例，不开后台服务

---

## 15. 模块接口摘要（供快速查阅）

```text
# detect
TextDetector.detect(image_path) -> list[TextBlock]
ShapeDetector.detect(image_path) -> list[Shape]
ImageRegionDetector.detect(image_path) -> list[ImageRegion]
BackgroundAnalyzer.analyze(image_path) -> Background
LayoutInferer.infer(text_blocks, shapes, image_regions, background) -> list[Element]

# assemble
StyleResolver.resolve(hint: dict) -> Style
FallbackDecider.decide(element: Element) -> RenderPlan
PptxBuilder.build(elements: list[Element], slide_size: SlideSize) -> bytes
ElementWriter._write_textbox / _write_shape / _write_image / _write_bitmap (internal)

# report
SlideRenderer.render(pptx_bytes: bytes, out_png: Path) -> Path
DiffCalculator.compute(orig_png: Path, rendered_png: Path) -> DiffResult
EditabilityChecker.check(pptx_path: Path) -> list[CheckItem]
HtmlBuilder.build(context: dict, out_html: Path) -> Path

# pipeline
Pipeline.run(image_path: Path, output_dir: Path, config: Config) -> JobResult

# common
validate_image(path: Path) -> None
BBox.to_emu() -> tuple[int, int, int, int]
Color.nearest_named(hex: str) -> str
read_json / write_json / ensure_dir
```

---

## 16. 后续可能扩展（MVP 之外）

- 多张幻灯片批处理
- 母版（Master）识别
- 表格识别（行 / 列分割）
- 嵌入图表数据反推
- 浏览器拖拽上传界面
- VS Code 插件入口

---

## 17. 当前文档状态

- [x] 模块边界清晰
- [x] CLI 合约冻结
- [x] JSON Schema 锁定 v1.0
- [x] 里程碑 M0~M5 拆解
- [x] 测试固件列表
- [x] 委派子任务清单
- [ ] 各子任务具体实现（M1 开始）

> 任何模块实现开始前，先核对 §6 JSON schema 与 §4 接口契约；如发现需要扩展，先回到本文档修改 `version` 字段到 `1.1` 并附录变更说明。
