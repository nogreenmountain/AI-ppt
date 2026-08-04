"""Tkinter desktop app for slide2pptx."""

from __future__ import annotations

import contextlib
import io
import json
import os
import queue
import shutil
import sys
import threading
import time
from pathlib import Path
from tkinter import BooleanVar, StringVar, Tk, filedialog, messagebox
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

from slide2pptx import pipeline_cli


APP_NAME = "AI PPT 拆页器"


def resource_root() -> Path:
    """Return the repo/resource root in source and PyInstaller modes."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parents[2]


def bundled_node(root: Path) -> str | None:
    candidates = [
        root / "runtime" / "node" / "node.exe",
        root / "runtime" / "node.exe",
    ]
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)
    env_node = os.environ.get("SLIDE2PPTX_NODE", "").strip()
    if env_node and Path(env_node).is_file():
        return env_node
    return shutil.which("node")


class Slide2PptxApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.repo_root = resource_root()
        self.events: queue.Queue[tuple[str, str]] = queue.Queue()
        self.worker: threading.Thread | None = None

        self.input_path = StringVar()
        default_out = Path.home() / "Documents" / "AI PPT 拆页器" / "输出" / f"任务-{time.strftime('%Y%m%d-%H%M%S')}"
        self.output_dir = StringVar(value=str(default_out))
        self.with_report = BooleanVar(value=False)
        self.status = StringVar(value="准备就绪")

        self._build_ui()
        self._poll_events()

    def _build_ui(self) -> None:
        self.root.title(APP_NAME)
        self.root.geometry("920x620")
        self.root.minsize(760, 520)
        self.root.configure(bg="#111318")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", font=("Segoe UI", 10))
        style.configure("Root.TFrame", background="#111318")
        style.configure("Panel.TFrame", background="#181B22")
        style.configure("Title.TLabel", background="#111318", foreground="#F5F7FA", font=("Segoe UI Semibold", 22))
        style.configure("Sub.TLabel", background="#111318", foreground="#AAB2C0", font=("Segoe UI", 10))
        style.configure("Panel.TLabel", background="#181B22", foreground="#E7EAF0")
        style.configure("Hint.TLabel", background="#181B22", foreground="#8D98A8")
        style.configure("TEntry", fieldbackground="#0F1117", foreground="#F5F7FA", bordercolor="#303747")
        style.configure("Primary.TButton", background="#4E9A8A", foreground="#07100E", font=("Segoe UI Semibold", 10), padding=(16, 8))
        style.map("Primary.TButton", background=[("active", "#69C2AE"), ("disabled", "#313841")])
        style.configure("TButton", background="#252B36", foreground="#F5F7FA", padding=(12, 7))
        style.map("TButton", background=[("active", "#303848")])
        style.configure("TCheckbutton", background="#181B22", foreground="#E7EAF0")
        style.configure("Horizontal.TProgressbar", troughcolor="#0F1117", background="#4E9A8A")

        outer = ttk.Frame(self.root, style="Root.TFrame", padding=24)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="AI PPT 拆页器", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="本地处理图片、PPT 和 PPTX，安装完成后核心拆解可离线运行。",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(2, 18))

        panel = ttk.Frame(outer, style="Panel.TFrame", padding=18)
        panel.pack(fill="x")
        panel.columnconfigure(1, weight=1)

        ttk.Label(panel, text="输入文件", style="Panel.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 12), pady=6)
        ttk.Entry(panel, textvariable=self.input_path).grid(row=0, column=1, sticky="ew", pady=6)
        ttk.Button(panel, text="选择文件", command=self.choose_input).grid(row=0, column=2, padx=(10, 0), pady=6)

        ttk.Label(panel, text="输出目录", style="Panel.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 12), pady=6)
        ttk.Entry(panel, textvariable=self.output_dir).grid(row=1, column=1, sticky="ew", pady=6)
        ttk.Button(panel, text="选择目录", command=self.choose_output).grid(row=1, column=2, padx=(10, 0), pady=6)

        ttk.Checkbutton(
            panel,
            text="同时生成 HTML 对比报告（需要本机安装 PowerPoint）",
            variable=self.with_report,
        ).grid(row=2, column=1, sticky="w", pady=(8, 2))
        ttk.Label(
            panel,
            text="输入 PPT/PPTX 时，会先调用本机 PowerPoint 把每一页导出为图片，再逐页拆解。",
            style="Hint.TLabel",
        ).grid(row=3, column=1, sticky="w", pady=(0, 4))

        actions = ttk.Frame(panel, style="Panel.TFrame")
        actions.grid(row=4, column=1, sticky="w", pady=(12, 0))
        self.run_button = ttk.Button(actions, text="开始拆解", style="Primary.TButton", command=self.start)
        self.run_button.pack(side="left")
        ttk.Button(actions, text="打开输出目录", command=self.open_output).pack(side="left", padx=(10, 0))

        self.progress = ttk.Progressbar(outer, mode="indeterminate")
        self.progress.pack(fill="x", pady=(18, 8))
        ttk.Label(outer, textvariable=self.status, style="Sub.TLabel").pack(anchor="w")

        self.log = ScrolledText(
            outer,
            height=16,
            bg="#0B0D12",
            fg="#DCE2EC",
            insertbackground="#DCE2EC",
            relief="flat",
            font=("Consolas", 10),
        )
        self.log.pack(fill="both", expand=True, pady=(10, 0))
        self.log.insert("end", "请选择一张图片，或一个 PPT/PPTX 文件，然后点击“开始拆解”。\n")
        self.log.configure(state="disabled")

    def choose_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择输入文件",
            filetypes=[
                ("幻灯片和图片", "*.png *.jpg *.jpeg *.webp *.bmp *.ppt *.pptx"),
                ("PowerPoint 文件", "*.ppt *.pptx"),
                ("图片文件", "*.png *.jpg *.jpeg *.webp *.bmp"),
                ("所有文件", "*.*"),
            ],
        )
        if path:
            self.input_path.set(path)

    def choose_output(self) -> None:
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir.set(path)

    def open_output(self) -> None:
        path = Path(self.output_dir.get())
        path.mkdir(parents=True, exist_ok=True)
        os.startfile(path)  # type: ignore[attr-defined]

    def append_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text)
        if not text.endswith("\n"):
            self.log.insert("end", "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def start(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        input_path = Path(self.input_path.get().strip())
        if not input_path.is_file():
            messagebox.showerror(APP_NAME, "请选择一个存在的图片、PPT 或 PPTX 文件。")
            return
        out_dir = Path(self.output_dir.get().strip())
        if not out_dir:
            messagebox.showerror(APP_NAME, "请选择输出目录。")
            return

        node = bundled_node(self.repo_root)
        if not node:
            messagebox.showerror(APP_NAME, "没有找到 Node.js。请使用安装版，或先安装 Node.js 20 及以上版本。")
            return

        self.run_button.configure(state="disabled")
        self.progress.start(12)
        self.status.set("正在拆解...")
        self.append_log(f"\n输入文件：{input_path}\n输出目录：{out_dir}\n")

        self.worker = threading.Thread(
            target=self._run_pipeline,
            args=(input_path, out_dir, node),
            daemon=True,
        )
        self.worker.start()

    def _run_pipeline(self, input_path: Path, out_dir: Path, node: str) -> None:
        os.environ["SLIDE2PPTX_NODE"] = node
        os.environ["SLIDE2PPTX_REPO_ROOT"] = str(self.repo_root)
        os.environ["PYTHONPATH"] = str(self.repo_root / "python")

        args = [
            str(input_path),
            "--out",
            str(out_dir),
            "--visual-passes",
            "2",
            "--second-pass-max-components",
            "96",
        ]
        if not self.with_report.get():
            args.append("--skip-report")

        stdout = io.StringIO()
        stderr = io.StringIO()
        try:
            with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
                code = pipeline_cli.main(args)
        except Exception as exc:  # noqa: BLE001 - GUI boundary
            self.events.put(("log", f"错误：{type(exc).__name__}: {exc}\n"))
            self.events.put(("done", "失败"))
            return

        summary = self._format_result(stdout.getvalue())
        if summary:
            self.events.put(("log", summary))
        elif stdout.getvalue():
            self.events.put(("log", stdout.getvalue()))
        if stderr.getvalue() and code != 0:
            self.events.put(("log", stderr.getvalue()))
        self.events.put(("done", "完成" if code == 0 else f"失败（退出码 {code}）"))

    def _format_result(self, raw: str) -> str:
        raw = raw.strip()
        if not raw:
            return ""
        start = raw.find("{")
        end = raw.rfind("}")
        if start == -1 or end == -1 or end <= start:
            return ""
        try:
            data = json.loads(raw[start:end + 1])
        except json.JSONDecodeError:
            return ""

        lines = []
        if data.get("input_presentation"):
            lines.append(f"输入 PPT：{data.get('input_presentation')}")
            lines.append(f"共拆解 {data.get('slide_count', 0)} 页。")
            failed = data.get("failed_slides") or []
            if failed:
                failed_indexes = "、".join(str(item.get("index")) for item in failed)
                lines.append(f"有 {len(failed)} 页未完成：第 {failed_indexes} 页。其他页面已继续处理。")
            for slide in data.get("slides", []):
                if slide.get("ok") is False:
                    lines.append(f"第 {slide.get('index')} 页失败：{slide.get('error')}")
                    continue
                lines.append(f"第 {slide.get('index')} 页 PPTX：{slide.get('pptx')}")
                retry_count = int(slide.get("retry_count") or 0)
                if retry_count:
                    lines.append(f"第 {slide.get('index')} 页已自动重试 {retry_count} 次后成功。")
                warnings = slide.get("warnings") or []
                if warnings:
                    lines.append(f"第 {slide.get('index')} 页提示：{'；'.join(self._zh_warning(w) for w in warnings)}")
        else:
            lines.append(f"输入图片：{data.get('input_image')}")
            build = data.get("build") or {}
            detect = data.get("detect") or {}
            lines.append(f"生成 PPTX：{build.get('pptx')}")
            retry_count = int(build.get("retry_count") or 0)
            if retry_count:
                lines.append(f"已自动重试 {retry_count} 次后成功。")
            lines.append(f"检测结果：{detect.get('detected_json')}")
            warnings = detect.get("warnings") or []
            if warnings:
                lines.append(f"提示：{'；'.join(self._zh_warning(w) for w in warnings)}")
        lines.append("")
        return "\n".join(lines)

    def _zh_warning(self, warning: str) -> str:
        known = {
            "OCR component not installed; native text elements skipped.": "软件缺少文字识别组件，本次不会生成可编辑文字框。",
            "heuristic text-like regions detected for component separation; install the OCR component to emit editable text.": "已用启发式方法分离疑似文字区域；完整安装后可生成可编辑文字。",
        }
        return known.get(warning, warning)

    def _poll_events(self) -> None:
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "log":
                    self.append_log(payload)
                elif kind == "done":
                    self.progress.stop()
                    self.run_button.configure(state="normal")
                    if payload == "完成":
                        self.status.set("已完成")
                        self.append_log("[完成] 拆解完成。\n")
                    else:
                        self.status.set("失败")
                        self.append_log(f"[错误] {payload}\n")
        except queue.Empty:
            pass
        self.root.after(100, self._poll_events)


def main() -> int:
    root = Tk()
    Slide2PptxApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
