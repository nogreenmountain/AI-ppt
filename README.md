# slide2pptx

把一张 16:9 幻灯片图片（PNG/JPG）重建成可在 PowerPoint 中继续编辑的 `.pptx`。

项目现在按“一键即用”整理：Windows 用户 clone 后可以直接运行 `convert_image_to_ppt.bat`。脚本会自动准备 Python 虚拟环境、安装 Python 依赖、安装 Node 依赖，并用 `samples/source.png` 跑出一个示例 PPTX。

## 一键使用

```powershell
git clone https://github.com/nogreenmountain/AI-ppt.git
cd AI-ppt
.\convert_image_to_ppt.bat
```

默认输出：

```text
outputs\one-click-<timestamp>\build\reconstructed.pptx
outputs\one-click-<timestamp>\detect\detected.json
```

转换你自己的图片：

```powershell
.\convert_image_to_ppt.bat "C:\path\to\slide.png"
.\convert_image_to_ppt.bat "C:\path\to\slide.png" "C:\path\to\out"
```

如果本机安装了 Microsoft PowerPoint，并且想生成 HTML 对比报告：

```powershell
.\convert_image_to_ppt.bat "C:\path\to\slide.png" "C:\path\to\out" -Report
```

## 自检

```powershell
.\setup_and_test.bat
```

这个脚本会执行：

- Python 报告模块测试
- JavaScript builder 测试
- PPTX 生成 smoke test

## 脚本会自动做什么

- 创建 `.venv`
- 安装 `requirements.txt`
- 安装 `artifact-runtime/package.json` 里的 npm 依赖
- 找不到 Python 或 Node.js 时，优先尝试用 `winget` 安装：
  - `Python.Python.3.12`
  - `OpenJS.NodeJS.LTS`

如果你的环境禁用了 `winget`，请先安装 Python 3.10+ 和 Node.js 20+，再重新运行脚本。

## 能力边界

当前版本是 MVP：

- 复杂背景、图表、插画会作为背景或图片层保留。
- OCR 可用时会输出可编辑文本框；OCR 不可用时会降级，只保留视觉重建。
- HTML 报告依赖 Windows + Microsoft PowerPoint COM 渲染；核心 PPTX 生成不依赖 PowerPoint。
- 不是原始 PPT 的完美反编译，无法恢复母版、动画、主题、SmartArt、嵌入对象等信息。

## 手动命令

检测图片：

```powershell
$env:PYTHONPATH = "$PWD\python"
.\.venv\Scripts\python.exe -m slide2pptx.detect_cli samples\source.png --out outputs\detect-only
```

从 `detected.json` 生成 PPTX：

```powershell
cd artifact-runtime
npm install
node src\convert.mjs --spec ..\outputs\detect-only\detected.json --out ..\outputs\detect-only\reconstructed.pptx
```

端到端 pipeline：

```powershell
$env:PYTHONPATH = "$PWD\python"
.\.venv\Scripts\python.exe -m slide2pptx.pipeline_cli samples\source.png --out outputs\manual-run --skip-report
```

## 项目结构

```text
artifact-runtime/      Node.js PPTX builder, powered by pptxgenjs
python/slide2pptx/     Python detector, pipeline, report renderer
samples/               示例输入图片
spec/                  detected.json schema
tests/                 Python tests
```

## License

MIT
