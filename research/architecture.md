# AI 幻灯片图像 → 可编辑 PPTX 端到端架构设计

> 版本：v0.1 草案
> 范围：单张 AI 生成幻灯片图像（PNG/JPG/WEBP）→ 视觉近 1:1 的可编辑 `.pptx`
> 非目标：批量流水线、动画、过渡效果、母版主题复用、跨幻灯片一致性

---

## 1. 摘要与核心判断

把"一张位图"反向重建为"可编辑 PPTX"，本质是一个**不可逆信息恢复问题**：原图的字号、字体、矢量路径、渐变 stops、混合模式在栅格化时已经丢失最多信息。下表给出本设计的核心判断：

| 维度 | 现实可达 | 不应承诺 |
|---|---|---|
| 文本可编辑 | 高（OCR + 重建） | 100% 字符无差错（艺术字/低分辨率必损） |
| 形状可编辑 | 中（基本几何可识别） | 任意矢量路径、布尔运算、3D |
| 颜色 | 高（采样即可） | 复杂渐变 stops、混合模式 |
| 字体 | 中（猜测 family） | 完美字重/字距/连字 |
| 图表/插画 | 低-中（回退为图片） | 还原数据系列、矢量节点 |
| 动画/过渡 | 不支持 | — |
| 整体 1:1 视觉 | 70%~92% SSIM | 像素级一致 |

**核心策略**：分层 + 置信度门控 + 可编辑优先 / 光栅回退 + 视觉闭环迭代。把"什么可编辑"当作一个**带置信度的决策**，而不是全有或全无。

---

## 2. 端到端管道总览

```
[ 输入图像 ]
   │
   ▼
(1) 预处理 ── 去噪 / 颜色归一化 / 尺寸归一 / 旋转校正
   │
   ▼
(2) 元素分解 ── 背景层 │ 装饰层 │ 文本区域 │ 形状/图标 │ 图像/插画
   │
   ▼
(3) OCR ── 文本区域 → 文本段（字符 + bbox + 置信度）
   │
   ▼
(4) 形状识别 ── 几何/图标分类（矩形、圆、线、箭头、图标）
   │
   ▼
(5) 布局重建 ── 读序、对齐、分组、z 序
   │
   ▼
(6) 决策层 ── 每元素：可编辑 │ 局部光栅 │ 全层光栅
   │
   ▼
(7) PPTX 合成 ── 写出形状/文本/图像，按 z 序
   │
   ▼
(8) 渲染回归 ── LibreOffice/POI 把 PPTX 转 PNG
   │
   ▼
(9) 视觉 diff ── SSIM + LPIPS + 锚点距离 → 不达标则回流 (3)-(7)
   │
   ▼
[ 输出 PPTX + 元素清单 + 置信度报告 ]
```

每个环节都是**显式有产出**的（不只是中间张量），便于回放、调参与失败定位。

---

## 3. 模块设计

### 3.1 预处理

- 目标分辨率：长边 1920px，超过则缩放（减少显存、加速后续）
- 颜色空间：保留 RGB；对纯装饰背景层单独处理
- 旋转检测：Hough 直线 + 文本行角度，>1° 时旋转校正
- 去噪：轻量 bilateral filter，仅对将要 OCR 的区域

### 3.2 元素分解（Decomposition）

**目标**：将图像分解为 5 类语义层 — 背景、装饰、文本区域、几何/图标、图像内容。

**推荐方案**（按成本递增）：

| 方案 | 工具 | 精度 | 成本 | 适用 |
|---|---|---|---|---|
| A. 传统 CV | OpenCV 边缘 + 连通域 + GrabCut | 低-中 | CPU、毫秒 | 简单版式 |
| B. SAM 通用分割 | Meta SAM ViT-B/L | 高（轮廓） | GPU、秒级 | 推荐基线 |
| C. 专用检测器 | YOLOv8/v11 训练于 slide 数据集 | 高（语义） | GPU、需训练数据 | 高保真 |
| D. 视觉 LLM | GPT-4o / Qwen2-VL | 高（语义） | API 成本 | 复杂版式兜底 |

**推荐组合**：`SAM ViT-B` 抠出所有前景 mask → 用颜色/纹理/边缘特征把 mask 分类到 5 类 → 文本区域送 OCR，图像内容单独走图像管线，几何/图标再细化。

**输出**：每元素 `{ id, type, mask, bbox, area, layer_z, confidence }`。

### 3.3 OCR

| 引擎 | 中英混排 | 旋转 | 艺术字 | bbox 粒度 | 离线 | 成本 |
|---|---|---|---|---|---|---|
| PaddleOCR 3.x | ★★★★★ | ★★★★ | ★★★ | 多边形/矩形 | ✓ | 免费 |
| PaddleOCR-VL | ★★★★★ | ★★★★ | ★★★★ | 块/行/词 | ✓ | 免费 |
| Tesseract 5 | ★★★ | ★★ | ★ | 词/行 | ✓ | 免费 |
| EasyOCR | ★★★★ | ★★★ | ★★ | 矩形 | ✓ | 免费 |
| Google Vision | ★★★★★ | ★★★★ | ★★★★ | 词 | ✗ | 按量 |
| GPT-4o vision | ★★★★★ | ★★★★ | ★★★★ | 文本（粗 bbox） | ✗ | 高 |
| Azure Read | ★★★★★ | ★★★★ | ★★★★ | 词+行 | ✗ | 按量 |

**推荐**：基线用 **PaddleOCR-VL**（开源、中文友好、带版面分析），对低置信度区域用云端 OCR 二次校验。

**OCR 输出**：
```json
{
  "blocks": [
    { "text": "AI 时代的产品设计", "bbox": [x1,y1,x2,y2], "poly": [...],
      "conf": 0.96, "lang": "zh", "angle": 0, "lines": [
        { "text": "AI 时代的产品设计", "conf": 0.97, "words": [...] }
      ]
    }
  ]
}
```

### 3.4 形状/图标识别

- **几何形状**：用 OpenCV `minAreaRect` + `fitEllipse` 判定矩形/圆/椭圆；箭头通过轮廓凸包 + 头部三角检测；线段通过 HoughLinesP。
- **图标**：用 `microsoft/phi-3-vision` 或 CLIP 零样本分类 top-K 候选；匹配到 emoji 或开源图标库（Nerd Font、Material Symbols）时优先匹配。
- **复杂插画**：识别失败直接归类为"图像内容"，整块光栅嵌入。

### 3.5 布局重建

**子任务**：
1. **读序**：自上而下、自左而右；标题优先于正文；同区域内按行优先。
2. **对齐**：聚类所有元素中心点和左右边，得到参考网格线（k-means / RANSAC）。
3. **分组**：距离阈值（如 < 24px）且语义相关（icon + label、title + subtitle）合并为 group。
4. **z 序**：从背景到前景按层写入；同层按"先生成后置顶"原则处理（在 PPTX 里后写 = 上层）。

**算法骨架**（伪代码）：
```python
def reconstruct_layout(elements):
    grid = detect_grid_lines(elements, axis='x|y', tol=8)
    elements = snap_to_grid(elements, grid)
    groups = cluster(elements, distance=24, semantic=True)
    order = topo_sort(groups, by=('y_top', 'x_left'))
    return groups, order, grid
```

### 3.6 决策层：可编辑 vs 光栅回退

每元素按下表判决定最终输出形态：

| 条件 | 输出 |
|---|---|
| 文本 & OCR conf ≥ 0.85 | 原生 TextBox（可编辑） |
| 文本 & 0.6 ≤ conf < 0.85 | 原生 TextBox + 标记 `low_conf` |
| 文本 & conf < 0.6 或高度 < 12px | 该块替换为光栅图像 |
| 几何形状（矩形/圆/线/箭头）& 拟合误差 < 3% | 原生 MSO_SHAPE |
| 几何形状 & 拟合误差 ≥ 3% | 形状近似 + 局部光栅（用剪贴遮罩） |
| 图标 & 匹配到 Nerd Font / emoji | 原生 TextBox（用对应字符） |
| 图标 & 不匹配 | 整块光栅 |
| 图像内容 | 始终光栅（图片） |
| 装饰背景（覆盖 > 40% 页面） | 单张全屏图作为底图 |
| 复杂渐变 / 混合模式 | 光栅（无法矢量重建） |
| 整页 conf 平均 < 0.5 | **全页光栅**（兜底） |

**关键原则**：宁可承认"这一块不可编辑"，也别输出错误的可编辑结果（用户会更难受）。

### 3.7 坐标映射（像素 → EMU）

python-pptx 用 EMU（English Metric Units）：

```
1 inch = 914400 EMU
1 pt   = 12700 EMU
1 cm   = 360000 EMU
16:9 幻灯片 = 13.333" × 7.5" = 12192000 × 6858000 EMU
```

**映射函数**：
```python
SLIDE_W_EMU = 12192000
SLIDE_H_EMU = 6858000

def px_to_emu(px_xy, img_w, img_h):
    x_px, y_px = px_xy
    x_emu = int(x_px / img_w * SLIDE_W_EMU)
    y_emu = int(y_px / img_h * SLIDE_H_EMU)
    return x_emu, y_emu

def pt_to_emu(pt): return int(pt * 12700)
```

**注意陷阱**：
- 字体大小在 PPTX 里是 `pt`，不是 px；OCR 通常不给字号，需要从字符高度反推：经验值 `font_pt ≈ pixel_height / 1.5`（需按字体微调）
- python-pptx 的 `add_textbox(left, top, width, height)` 全部单位是 EMU

### 3.8 PPTX 合成

**库选择矩阵**：

| 库 | 能力 | 限制 | 推荐度 |
|---|---|---|---|
| python-pptx | 文本/形状/图像/表格/图表 | 渐变 stops、3D、动画 | ★★★★ |
| pptxgenjs (Node) | 较丰富 | Node 依赖 | ★★★ |
| Aspose.Slides | 接近完整 OOXML | 商业许可 | ★★★★★（预算允许时） |
| Spire.Presentation | 较完整 | 商业、有限免费 | ★★★★ |
| Apache POI | Java 完整 | JVM 依赖 | ★★★ |

**MVP 推荐**：python-pptx。足够覆盖 80% 场景，必要时把个别高保真需求外包给 LibreOffice / Aspose。

**写入骨架**：
```python
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

prs = Presentation()
prs.slide_width  = Emu(SLIDE_W_EMU)
prs.slide_height = Emu(SLIDE_H_EMU)
slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

for el in ordered_elements:
    if el.decision == 'native_shape':
        shp = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Emu(el.x), Emu(el.y), Emu(el.w), Emu(el.h)
        )
        shp.fill.solid(); shp.fill.fore_color.rgb = RGBColor(*el.color)
    elif el.decision == 'native_text':
        tb = slide.shapes.add_textbox(Emu(el.x), Emu(el.y), Emu(el.w), Emu(el.h))
        tf = tb.text_frame
        for line in el.lines:
            p = tf.add_paragraph()
            run = p.add_run()
            run.text = line.text
            run.font.size = Pt(el.font_pt)
            run.font.name = el.font_name or 'Microsoft YaHei'
            run.font.color.rgb = RGBColor(*el.color)
    elif el.decision == 'raster':
        slide.shapes.add_picture(el.png_path, Emu(el.x), Emu(el.y), Emu(el.w), Emu(el.h))
```

### 3.9 置信度评分

**5 维评分**（每维 0~1）：

| 维度 | 含义 | 测量方式 |
|---|---|---|
| T — Text | 文本识别准确性 | OCR engine conf 平均 |
| S — Shape | 形状拟合质量 | mask IoU vs 拟合几何 |
| C — Color | 颜色还原度 | ΔE (CIEDE2000) in 关键像素 |
| P — Position | 位置准确度 | bbox 中心点 L2 距离 / 图像对角线 |
| Z — z-order | 层叠正确性 | 重投影后从上到下遮挡关系一致性 |

**聚合**：
```python
score = 0.30*T + 0.25*S + 0.15*C + 0.20*P + 0.10*Z
```

**门控阈值**：

| score | 行为 |
|---|---|
| ≥ 0.85 | 直接通过 |
| 0.65 ~ 0.85 | 走迭代校正（最多 3 轮） |
| 0.40 ~ 0.65 | 标记可疑区域，建议人工复核 |
| < 0.40 | 该元素 / 整页回退为光栅 |

### 3.10 迭代视觉校正

```
渲染: PPTX → PNG (LibreOffice headless soffice --convert-to png)
对齐: 用 ORB/SIFT 找仿射矩阵，校正后再比
diff:  SSIM 全图 + LPIPS 关键块 + 锚点距离
诊断: 把 diff heatmap 叠加在原图，定位失效区域
回流: 针对失效区域重新走 OCR / 形状识别 / 颜色采样
```

**实现要点**：
- 渲染用 LibreOffice：`soffice --headless --convert-to png slide.pptx`
- 锚点：取 5~10 个高对比度角点（标题左上角、段落首行、形状中心），重投影误差 > 8px 视为布局漂移
- 限制：最多 3 轮，3 轮不达标则停止（避免无限循环）
- 显式记录每轮 diff，便于解释为何停

### 3.11 元素清单与可解释性

输出一个 sidecar JSON：
```json
{
  "elements": [
    {"id":"e1","type":"text","decision":"native_text","conf":0.93,"text":"标题"},
    {"id":"e2","type":"shape","decision":"native_shape","conf":0.88,"shape":"rect"},
    {"id":"e3","type":"icon","decision":"raster","conf":0.45,"reason":"low_match"}
  ],
  "page_score": 0.82,
  "warnings": ["e3 不可编辑，已嵌入图片"],
  "rounds": 2
}
```

让用户清楚知道哪一块可改、哪一块是图。

---

## 4. MVP 管道（建议的最小可行实现）

```python
# pipeline.py  —— 单文件骨架（生产请拆模块）
from pathlib import Path
from PIL import Image
import numpy as np
from paddleocr import PaddleOCR
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

SLIDE_W_EMU, SLIDE_H_EMU = 12192000, 6858000

def step1_preprocess(img_path):
    img = Image.open(img_path).convert('RGB')
    w, h = img.size
    if max(w, h) > 1920:
        scale = 1920 / max(w, h)
        img = img.resize((int(w*scale), int(h*scale)))
    return np.array(img), img.size

def step2_decompose(img_arr):
    # TODO: 接入 SAM / YOLO；MVP 用颜色聚类 + 边缘
    # 返回 [{mask, bbox, type}]
    return []

def step3_ocr(img_arr, regions):
    ocr = PaddleOCR(use_angle_cls=True, lang='ch')
    blocks = []
    for r in regions:
        x1,y1,x2,y2 = r['bbox']
        crop = img_arr[y1:y2, x1:x2]
        result = ocr.ocr(crop, cls=True)
        for line in result:
            if not line: continue
            for box, (text, conf) in line:
                blocks.append({'text': text, 'conf': conf,
                               'bbox': [int(c) for pt in box for c in pt]})
    return blocks

def step4_emit_pptx(img_arr, img_size, elements, blocks, out_path):
    prs = Presentation()
    prs.slide_width  = Emu(SLIDE_W_EMU)
    prs.slide_height = Emu(SLIDE_H_EMU)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    W, H = img_size

    # 底图占位（可省略）
    # slide.shapes.add_picture('original.png', 0, 0,
    #                          width=prs.slide_width, height=prs.slide_height)

    for b in sorted(blocks, key=lambda b: (b['bbox'][1], b['bbox'][0])):
        x1,y1,x2,y2 = b['bbox']
        x = Emu(int(x1 / W * SLIDE_W_EMU))
        y = Emu(int(y1 / H * SLIDE_H_EMU))
        w = Emu(int((x2-x1) / W * SLIDE_W_EMU))
        h = Emu(int((y2-y1) / H * SLIDE_H_EMU))
        tb = slide.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]
        run = p.add_run(); run.text = b['text']
        # 粗估字号：px高 / 1.5
        run.font.size = Pt(max(8, (y2-y1) / 1.5))
        run.font.name = 'Microsoft YaHei'

    prs.save(out_path)
    return out_path

def main(img_in, pptx_out):
    img_arr, size = step1_preprocess(img_in)
    regions = step2_decompose(img_arr)   # MVP 暂用全图
    blocks  = step3_ocr(img_arr, regions)
    return step4_emit_pptx(img_arr, size, [], blocks, pptx_out)
```

**MVP 边界**：
- 只处理单页
- 不接 SAM（先用全图 OCR）
- 不做颜色/形状还原
- 不做迭代校正

跑通后按 3.x 节增量引入。

---

## 5. 验收指标

| 类别 | 指标 | 目标 | 备注 |
|---|---|---|---|
| 文本 | OCR 字符准确率 | ≥ 95% | 人工标注 50 张 |
| 文本 | 文本块召回率 | ≥ 90% | 漏检 = 关键缺陷 |
| 形状 | 形状类型分类准确率 | ≥ 85% | 矩形/圆/线/箭头 |
| 几何 | 拟合 IoU | ≥ 0.85 | 与 mask 比 |
| 颜色 | ΔE2000 平均 | ≤ 6 | 关键采样点 |
| 位置 | 中心点平均偏移 | ≤ 8px @1920 | |
| 整体 | SSIM | ≥ 0.85 | 全图 |
| 整体 | LPIPS | ≤ 0.20 | |
| 可编辑率 | 可编辑元素占比 | ≥ 60% | 行业典型目标 |
| 端到端 | 单页处理时延 | ≤ 8s @GPU | 不含视觉 diff |
| 端到端 | 失败回退成功率 | ≥ 99% | 即任何输入都能出 PPTX |

**测试集**：50 张 AI 生成样张 + 50 张真实设计师样张，分布涵盖：纯文字、图文混排、图标密集、复杂版式、低对比度、艺术字。

---

## 6. 现实局限性（必须坦白）

1. **字体**：OCR 不知道原图用的什么字体，只能猜 family（黑体/宋体/无衬线）。最终用户需手动确认；最坏情况需要把字体文件打包随 PPTX 一起交付。
2. **字号**：从像素高度反推是经验公式，不同字体差异 ±20%。无法精确还原。
3. **字重/字距/连字**：基本无法从位图恢复。
4. **渐变 stops**：python-pptx 写出线性渐变可行，但 stop 位置/颜色是采样插值，不是原参数。
5. **混合模式（glassmorphism、模糊、阴影）**：仅可"近似"，效果差异肉眼可见。
6. **矢量路径**：自由曲线/手绘形状基本只能光栅。
7. **图表**：图像中的图表基本只能光栅；从图像反推数据是另一个研究问题。
8. **z 序冲突**：两个半透明元素交叠时，原图遮挡关系可能在反向重建时丢信息。
9. **多语言脚本**：阿拉伯文、印地语、竖排中文支持依赖 OCR 引擎。
10. **母版/主题**：单页方案不感知原 PPT 的母版、配色变量、占位符逻辑。
11. **动画/过渡/超链接**：本方案不处理。
12. **像素级 1:1**：在不引入矢量重建大模型的前提下，**不可能**实现。SSIM 0.85+ 是更现实的目标。

---

## 7. 决策矩阵速查

| 场景 | 推荐路径 |
|---|---|
| 简单商务汇报页 | PaddleOCR + OpenCV 几何 + python-pptx |
| 复杂版式 + 艺术字 | PaddleOCR-VL + SAM + GPT-4o 兜底 + python-pptx |
| 必须像素级一致 | 全页光栅 + 元素清单（不做"可编辑"承诺） |
| 必须可二次编辑 | 限制输入风格（简单版式），输出 MVP 管道 |
| 预算充足 + 高保真 | Aspose.Slides 主路 + python-pptx 旁路 |

---

## 8. 假设（Assumptions）

- 输入图像分辨率 ≥ 1280×720，主体未被裁切
- 图像为横版 16:9（或可自动 letterbox/pillarbox）
- 文本主要为简体中文 + 英文
- 运行环境具备 GPU（用于 SAM/PaddleOCR 推理）
- 视觉 diff 用 LibreOffice 渲染 PPTX，可接受 ≥ 200ms 单页开销
- 不要求保留原 PPT 母版/主题/动画

---

## 9. 未决风险（Open Risks）

| 风险 | 影响 | 缓解 |
|---|---|---|
| SAM 在装饰/纹理背景上过度分割 | 元素爆炸、噪声 | 加面积+边缘强度过滤 |
| 视觉 LLM 误判版式层级 | z 序错乱 | 显式规则兜底 + 人工标记 |
| LibreOffice 与 PowerPoint 渲染差异 | diff 永远不收敛 | 用 PowerPoint Online API 替代或仅做主色+位置级 diff |
| 复杂渐变/混合模式无法还原 | 视觉保真度低 | 允许降级为光栅 |
| 字体版权 | 商业合规问题 | 仅嵌入开源/已授权字体 |
| 单页 > 10s 时延 | 用户体验差 | 异步 + 进度反馈 |
| 置信度阈值难调优 | 漏检/误检 | 提供配置中心 + 影子模式 |
| GPT-4o 等闭源 API 不可用 | 兜底失效 | 准备 Qwen2-VL/InternVL 开源替代 |
| OCR 对低分辨率小字失效 | 文本丢失 | 局部超分（Real-ESRGAN）后二次 OCR |

---

## 10. 路线图建议

| 阶段 | 范围 | 周期估 |
|---|---|---|
| P0 MVP | 预处理 + PaddleOCR + python-pptx + 全图 OCR | 1~2 周 |
| P1 元素分解 | 接入 SAM 抠前景 + 简单几何识别 | 2~3 周 |
| P2 决策层 | 5 维置信度 + 可编辑/光栅分流 | 1~2 周 |
| P3 视觉闭环 | LibreOffice 渲染 + SSIM/LPIPS 迭代 | 1~2 周 |
| P4 图标匹配 | CLIP 零样本 + 字体图标库命中 | 1 周 |
| P5 兜底 | 闭源 VLM 兜底 + 整页光栅回退 | 1 周 |

---

## 11. 一次性结论

- **走"分层 + 置信度门控 + 视觉闭环"路线**，不追求"完美可编辑"
- **基线栈**：PaddleOCR-VL + SAM ViT-B + OpenCV 几何 + python-pptx + LibreOffice 渲染回环
- **诚实目标**：SSIM ≥ 0.85，可编辑元素占比 ≥ 60%，字符准确率 ≥ 95%
- **必须保留"光栅回退"开关**，对低置信度区域主动降级而不是硬上
- **第一次跑通**以 MVP（仅文本）为锚，逐步引入分解/形状/颜色/闭环

---

> 备注：本报告未引用在线资源；所有库与阈值基于通用领域知识。落地前应对所选版本的 API（特别是 python-pptx 渐变 stop 操作、PaddleOCR-VL 版面分析输出 schema）做一次实机验证。
