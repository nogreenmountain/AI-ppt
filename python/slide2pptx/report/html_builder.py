"""Self-contained HTML comparison report.

Given a :class:`slide2pptx.report.models.ReportInputs`, the builder
emits a single HTML file with all images inlined as base64 data URIs.
That way the file is fully self-contained and works in any browser on
a sealed-down Windows box.

We use :mod:`jinja2` for template rendering but explicitly avoid
shipping a template as an external file - the template lives in this
module so the report builder only depends on the source tree it was
installed in.

The CSS is also inlined and deliberately conservative: works in
Chrome/Edge/Firefox and degrades gracefully when JavaScript is
disabled.
"""

from __future__ import annotations

import base64
import datetime as _dt
import mimetypes
from html import escape
from pathlib import Path
from typing import Optional, Sequence

from jinja2 import Environment, StrictUndefined

from slide2pptx.report.models import ReportInputs


_TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Slide2PPTX Report — {{ inputs.job_id }}</title>
<style>
  :root {
    --bg: #0e1116;
    --panel: #161b22;
    --panel-2: #1c222c;
    --border: #2b3138;
    --text: #e6edf3;
    --muted: #9aa5b1;
    --accent: #f0883e;
    --good: #3fb950;
    --warn: #d29922;
    --bad: #f85149;
    --mono: ui-monospace, "Cascadia Code", Consolas, monospace;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    background: var(--bg);
    color: var(--text);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    line-height: 1.5;
  }
  header {
    padding: 24px 32px;
    border-bottom: 1px solid var(--border);
    background: linear-gradient(180deg, var(--panel) 0%, var(--bg) 100%);
  }
  header h1 { margin: 0 0 4px; font-size: 22px; }
  header .job {
    color: var(--muted);
    font-family: var(--mono);
    font-size: 13px;
  }
  .container {
    padding: 24px 32px;
    max-width: 1600px;
    margin: 0 auto;
  }
  section {
    background: var(--panel);
    border: 1px solid var(--border);
    border-radius: 8px;
    margin-bottom: 20px;
    overflow: hidden;
  }
  section h2 {
    margin: 0;
    padding: 12px 16px;
    border-bottom: 1px solid var(--border);
    background: var(--panel-2);
    font-size: 15px;
    letter-spacing: .3px;
  }
  .metrics {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 12px;
    padding: 16px;
  }
  .metric {
    background: var(--panel-2);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 12px;
  }
  .metric .label { color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: .4px; }
  .metric .value { font-size: 22px; font-family: var(--mono); margin-top: 6px; }
  .metric .value.good { color: var(--good); }
  .metric .value.warn { color: var(--warn); }
  .metric .value.bad  { color: var(--bad); }
  .three-up {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 16px;
    padding: 16px;
  }
  .three-up figure { margin: 0; }
  .three-up figcaption {
    color: var(--muted);
    font-size: 13px;
    margin-bottom: 8px;
    text-transform: uppercase;
    letter-spacing: .3px;
  }
  .three-up img {
    max-width: 100%;
    display: block;
    border: 1px solid var(--border);
    border-radius: 4px;
    background: #000;
  }
  table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
  }
  th, td {
    text-align: left;
    padding: 8px 12px;
    border-bottom: 1px solid var(--border);
  }
  th { background: var(--panel-2); color: var(--muted); text-transform: uppercase; font-size: 11px; letter-spacing: .4px; }
  td.id { font-family: var(--mono); color: var(--accent); }
  .pill {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 999px;
    border: 1px solid currentColor;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: .3px;
  }
  .pill.native { color: var(--good); }
  .pill.image, .pill.background { color: var(--warn); }
  .pill.unknown { color: var(--bad); }
  ul.checklist {
    list-style: none;
    margin: 0;
    padding: 16px 24px;
  }
  ul.checklist li {
    padding: 6px 0;
    font-family: var(--mono);
    font-size: 13px;
  }
  ul.checklist li.pass::before { content: "✓"; color: var(--good); margin-right: 8px; }
  ul.checklist li.fail::before { content: "✗"; color: var(--bad); margin-right: 8px; }
  ul.checklist li.muted::before { content: "–"; color: var(--muted); margin-right: 8px; }
  .warnings {
    background: var(--panel-2);
    border-left: 3px solid var(--warn);
    margin: 16px;
    padding: 12px 16px;
    border-radius: 4px;
  }
  .warnings li { color: var(--warn); font-family: var(--mono); font-size: 12px; }
  footer {
    color: var(--muted);
    font-size: 12px;
    padding: 16px 32px 32px;
    border-top: 1px solid var(--border);
  }
  .status-banner {
    display: inline-block;
    margin-left: 12px;
    padding: 2px 8px;
    border-radius: 4px;
    border: 1px solid currentColor;
    font-size: 12px;
  }
  .status-banner.full { color: var(--good); }
  .status-banner.partial { color: var(--warn); }
  .status-banner.fallback { color: var(--bad); }
</style>
</head>
<body>
<header>
  <h1>Slide2PPTX Comparison Report</h1>
  <div>
    <span class="job">job: {{ inputs.job_id }}</span>
    <span class="status-banner {{ banner_class }}">{{ status_label }}</span>
  </div>
</header>

<div class="container">

  <section>
    <h2>Inputs</h2>
    <div class="metrics">
      <div class="metric">
        <div class="label">Source</div>
        <div class="value">{{ source_dim }}</div>
      </div>
      <div class="metric">
        <div class="label">PPTX</div>
        <div class="value">{{ pptx_status }}</div>
      </div>
      <div class="metric">
        <div class="label">Total Elements</div>
        <div class="value">{{ editability.total_elements }}</div>
      </div>
      <div class="metric">
        <div class="label">Editable (Native)</div>
        <div class="value {{ 'good' if editability.editable_ratio >= 0.6 else 'warn' }}">
          {{ editability.editable_count }} / {{ editability.total_elements }}
        </div>
      </div>
    </div>
  </section>

  <section>
    <h2>Visual Comparison</h2>
    <div class="three-up">
      <figure>
        <figcaption>Source (PNG)</figcaption>
        <img src="{{ source_data_uri }}" alt="source image">
      </figure>
      <figure>
        <figcaption>Rendered (PowerPoint export)</figcaption>
        <img src="{{ rendered_data_uri }}" alt="rendered image">
      </figure>
      <figure>
        <figcaption>Diff Heatmap</figcaption>
        <img src="{{ heatmap_data_uri }}" alt="diff heatmap">
      </figure>
    </div>
  </section>

  <section>
    <h2>Pixel-Level Metrics</h2>
    <div class="metrics">
      <div class="metric">
        <div class="label">MAE</div>
        <div class="value">{{ "%.3f"|format(metrics.mae) }}</div>
      </div>
      <div class="metric">
        <div class="label">RMSE</div>
        <div class="value">{{ "%.3f"|format(metrics.rmse) }}</div>
      </div>
      <div class="metric">
        <div class="label">SSIM</div>
        <div class="value {{ ssim_class }}">{{ "%.3f"|format(metrics.ssim) }}</div>
      </div>
      <div class="metric">
        <div class="label">Pixel Diff Ratio</div>
        <div class="value {{ diff_class }}">{{ "%.3f"|format(metrics.pixel_diff_ratio) }}</div>
      </div>
    </div>
  </section>

  <section>
    <h2>Elements ({{ rows|length }})</h2>
    <table>
      <thead>
        <tr>
          <th>ID</th><th>Kind</th><th>Strategy</th><th>Editable Score</th>
          <th>Confidence</th><th>BBox (left, top, w, h)</th><th>Text</th>
        </tr>
      </thead>
      <tbody>
        {% for row in rows %}
        <tr>
          <td class="id">{{ row.id }}</td>
          <td>{{ row.kind }}</td>
          <td><span class="pill {{ row.strategy }}">{{ row.strategy }}</span></td>
          <td>{{ "%.2f"|format(row.editable_score) }}</td>
          <td>{{ row.confidence or '—' }}</td>
          <td>{{ row.bbox }}</td>
          <td>{{ row.text }}</td>
        </tr>
        {% endfor %}
      </tbody>
    </table>
  </section>

  <section>
    <h2>Editability Checklist</h2>
    <ul class="checklist">
      {% for item in checklist %}
        <li class="{{ item.state }}">{{ item.label }}{% if item.detail %} — <span style="color: var(--muted);">{{ item.detail }}</span>{% endif %}</li>
      {% endfor %}
    </ul>
    {% if warnings %}
    <ul class="warnings">
      {% for w in warnings %}<li>{{ w }}</li>{% endfor %}
    </ul>
    {% endif %}
  </section>

  <section>
    <h2>Timings</h2>
    <table>
      <thead><tr><th>Stage</th><th>Duration (ms)</th></tr></thead>
      <tbody>
        {% for k, v in timings %}
        <tr><td>{{ k }}</td><td>{{ "%.1f"|format(v) }}</td></tr>
        {% endfor %}
      </tbody>
    </table>
  </section>

</div>

<footer>
  Generated by slide2pptx.report on {{ generated_at }}{% if pptx_path %} &middot; PPTX: <code>{{ pptx_path }}</code>{% endif %}{% if detected_json_path %} &middot; Detected JSON: <code>{{ detected_json_path }}</code>{% endif %}
</footer>
</body>
</html>
"""


def _data_uri(path: Path) -> str:
    """Inline ``path`` as ``data:image/...;base64,...``."""
    p = Path(path)
    if not p.is_file():
        return ""
    mime, _ = mimetypes.guess_type(str(p))
    if mime is None:
        mime = "image/png"
    blob = p.read_bytes()
    encoded = base64.b64encode(blob).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def _image_dimensions(path: Path) -> str:
    """Best-effort ``WxH`` string for the source image; falls back to ``?``."""
    try:
        from PIL import Image  # local import to keep module importable
    except Exception:
        return "?"
    try:
        with Image.open(str(path)) as src:
            w, h = src.size
        return f"{w} x {h}"
    except Exception:
        return "?"


def _checklist_items(editability, pptx_path) -> list:
    """Translate :class:`EditabilitySummary` into flat checklist entries."""
    items = []
    total = editability.total_elements
    items.append({
        "label": "Pipeline produced a non-empty detected.json",
        "state": "pass" if total > 0 else "fail",
        "detail": f"{total} elements",
    })
    items.append({
        "label": "At least one editable element",
        "state": "pass" if editability.editable_count > 0 else "fail",
        "detail": f"{editability.editable_count} editable / {total} total",
    })
    items.append({
        "label": "Bitmap fallback rate below 50%",
        "state": "pass" if editability.bitmap_fallback_count * 2 <= total else "fail",
        "detail": f"{editability.bitmap_fallback_count} bitmap / {total} total",
    })
    items.append({
        "label": "Average editable score >= 0.5",
        "state": "pass" if editability.avg_editable_score >= 0.5 else "fail",
        "detail": f"avg={editability.avg_editable_score:.2f}",
    })
    items.append({
        "label": "PPTX file present",
        "state": "muted",
        "detail": str(pptx_path) if pptx_path else "not provided",
    })
    return items


def _ssim_class(ssim: float) -> str:
    if ssim >= 0.85:
        return "good"
    if ssim >= 0.6:
        return "warn"
    return "bad"


def _diff_class(ratio: float) -> str:
    if ratio <= 0.05:
        return "good"
    if ratio <= 0.15:
        return "warn"
    return "bad"


def _status_banner(status) -> tuple:
    mapping = {
        "fully_editable": ("full", "Fully Editable"),
        "bitmap_fallback": ("fallback", "Bitmap-Only"),
        "low_confidence": ("partial", "Low Confidence"),
        "unknown": ("partial", "No Data"),
    }
    return mapping.get(getattr(status, "value", status), ("partial", "Unknown"))


def _clamp_text(text: str, limit: int = 200) -> str:
    if len(text) > limit:
        return text[: limit - 1] + "\u2026"
    if len(text) <= limit:
        return text
    return text[: limit - 1] + "…"


def render_report(
    inputs: ReportInputs,
    detected_elements: Optional[Sequence] = None,
    *,
    out_path: Optional[Path] = None,
) -> str:
    """Render the HTML report and (optionally) write it to disk.

    Args:
        inputs: Bundled report inputs (see :class:`ReportInputs`).
        detected_elements: Optional precomputed element rows. When not
            provided the builder derives checklist-only data and skips
            the element table.
        out_path: Destination ``.html`` file. Its parent directory is
            created on demand. If not provided the HTML string is
            returned and nothing is written.

    Returns:
        The full HTML string.
    """
    env = Environment(autoescape=True, undefined=StrictUndefined)
    template = env.from_string(_TEMPLATE)

    elements = detected_elements or []
    rows = []
    for element in elements:
        bbox = element.get("bbox", {}) or {}
        bbox_str = (
            f"{bbox.get('left', 0):.0f}, {bbox.get('top', 0):.0f}, "
            f"{bbox.get('width', 0):.0f}, {bbox.get('height', 0):.0f}"
        )
        confidence = None
        conf = element.get("confidence")
        if isinstance(conf, dict):
            values = [v for v in conf.values() if isinstance(v, (int, float))]
            if values:
                confidence = f"{max(values):.2f}"
        rows.append({
            "id": element.get("id", ""),
            "kind": element.get("kind", element.get("__kind__", "?")),
            "strategy": element.get("render_strategy", element.get("__strategy__", "?")),
            "editable_score": float(element.get("editable_score", 0.0) or 0.0),
            "confidence": confidence,
            "bbox": bbox_str,
            "text": _clamp_text(element.get("text", "") or ""),
        })

    metrics = inputs.diff_metrics
    editability = inputs.editability
    banner_class, status_label = _status_banner(editability.status)

    context = {
        "inputs": inputs,
        "metrics": metrics,
        "editability": editability,
        "rows": rows,
        "source_data_uri": _data_uri(inputs.source_image),
        "rendered_data_uri": _data_uri(inputs.rendered_image),
        "heatmap_data_uri": _data_uri(inputs.heatmap),
        "ssim_class": _ssim_class(metrics.ssim),
        "diff_class": _diff_class(metrics.pixel_diff_ratio),
        "checklist": _checklist_items(editability, inputs.pptx_path),
        "warnings": list(inputs.warnings) + list(editability.warnings),
        "banner_class": banner_class,
        "status_label": status_label,
        "pptx_status": "present" if inputs.pptx_path and inputs.pptx_path.exists() else "n/a",
        "pptx_path": escape(str(inputs.pptx_path)) if inputs.pptx_path else None,
        "detected_json_path": escape(str(inputs.detected_json_path)) if inputs.detected_json_path else None,
        "source_dim": _image_dimensions(inputs.source_image),
        "timings": sorted(inputs.timings_ms.items()),
        "generated_at": _dt.datetime.now().isoformat(timespec="seconds"),
    }

    html = template.render(**context)

    if out_path is not None:
        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(html, encoding="utf-8")
    return html


__all__ = ["render_report"]
