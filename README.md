# slide2pptx / AI PPT 拆页器

本地离线的幻灯片拆解工具：输入一张幻灯片图片，或输入一个 PPT/PPTX 文件，把页面逐页拆成可继续编辑的 `.pptx`。

核心流程不需要联网。第一次安装或构建时需要下载依赖；安装完成后，桌面软件可以在本地离线处理图片。PPT/PPTX 输入需要 Windows 上安装 Microsoft PowerPoint，用它把每一页本地导出成图片后再逐页拆解。

## 桌面软件

构建可安装的本地软件：

```powershell
git clone https://github.com/nogreenmountain/AI-ppt.git
cd AI-ppt
.\build_windows_app.bat
```

构建并安装到本机，同时创建桌面和开始菜单快捷方式：

```powershell
.\install_desktop_app.bat
```

安装后运行 `AI PPT 拆页器`，选择图片、PPT 或 PPTX，再选择输出目录即可。

## 一键脚本

直接处理示例图片：

```powershell
.\convert_image_to_ppt.bat
```

处理自己的图片：

```powershell
.\convert_image_to_ppt.bat "C:\path\to\slide.png"
.\convert_image_to_ppt.bat "C:\path\to\slide.png" "C:\path\to\out"
```

处理 PPT/PPTX，逐页拆解：

```powershell
.\convert_image_to_ppt.bat "C:\path\to\deck.pptx" "C:\path\to\out"
```

PPT/PPTX 输出结构：

```text
out\
  source-slides\
    slide-001.png
    slide-002.png
  slide-001\
    build\reconstructed.pptx
    detect\detected.json
  slide-002\
    build\reconstructed.pptx
    detect\detected.json
```

如果本机安装了 Microsoft PowerPoint，并且想生成 HTML 对比报告：

```powershell
.\convert_image_to_ppt.bat "C:\path\to\slide.png" "C:\path\to\out" -Report
```

## 自检

```powershell
.\setup_and_test.bat
```

会执行 Python 测试、JavaScript builder 测试和 PPTX 生成 smoke test。

## 自动准备内容

脚本会自动：

- 创建 `.venv`
- 安装 `requirements.txt`
- 安装 `artifact-runtime/package.json` 中的 Node 依赖
- 找不到 Python 或 Node.js 时，优先尝试用 `winget` 安装：
  - `Python.Python.3.12`
  - `OpenJS.NodeJS.LTS`

如果环境禁用了 `winget`，请先安装 Python 3.10+ 和 Node.js 20+。

## 能力边界

- 图片输入：核心转换不需要 PowerPoint。
- PPT/PPTX 输入：需要本机 PowerPoint 导出每页图片。
- OCR 可用时会输出可编辑文本框；OCR 不可用时会降级为视觉重建。
- 复杂图表、插画、纹理背景会作为图片层保留。
- 这不是原始 PPT 的完美反编译，无法恢复母版、动画、主题、SmartArt、嵌入对象等信息。

## 手动命令

端到端 pipeline：

```powershell
$env:PYTHONPATH = "$PWD\python"
$env:SLIDE2PPTX_NODE = "C:\path\to\node.exe"
.\.venv\Scripts\python.exe -m slide2pptx.pipeline_cli samples\source.png --out outputs\manual-run --skip-report
```

从 `detected.json` 生成 PPTX：

```powershell
cd artifact-runtime
npm install
node src\convert.mjs --spec ..\outputs\manual-run\detect\detected.json --out ..\outputs\manual-run\reconstructed.pptx
```

启动桌面应用源码版：

```powershell
$env:PYTHONPATH = "$PWD\python"
.\.venv\Scripts\python.exe -m slide2pptx.gui_app
```

## 项目结构

```text
artifact-runtime/      Node.js PPTX builder, powered by pptxgenjs
python/slide2pptx/     Python detector, PPT input, desktop GUI, pipeline, report renderer
samples/               示例输入图片
scripts/               一键运行、打包、安装脚本
spec/                  detected.json schema
tests/                 Python tests
```

## License

MIT
