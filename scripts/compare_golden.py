from __future__ import annotations

import argparse
import json
import re
import statistics
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


SLIDE_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def load_detected_texts(root: Path) -> tuple[dict, list[dict]]:
    detected_path = root / "detect" / "detected.json"
    data = json.loads(detected_path.read_text(encoding="utf-8"))
    texts = [el for el in data.get("elements", []) if el.get("kind") == "text"]
    return data, texts


def pptx_font_sizes(pptx_path: Path) -> list[float]:
    sizes: list[float] = []
    with zipfile.ZipFile(pptx_path) as package:
        for name in package.namelist():
            if not re.fullmatch(r"ppt/slides/slide\d+\.xml", name):
                continue
            root = ET.fromstring(package.read(name))
            for rpr in root.iter(f"{SLIDE_NS}rPr"):
                raw_size = rpr.attrib.get("sz")
                if raw_size:
                    sizes.append(int(raw_size) / 100)
    return sizes


def stat(values: list[float]) -> dict[str, float | int | None]:
    if not values:
        return {"count": 0, "avg": None, "median": None, "min": None, "max": None}
    return {
        "count": len(values),
        "avg": round(statistics.mean(values), 4),
        "median": round(statistics.median(values), 4),
        "min": round(min(values), 4),
        "max": round(max(values), 4),
    }


def diff_number(label: str, expected: float, actual: float, tolerance: float, errors: list[str]) -> None:
    if abs(actual - expected) > tolerance:
        errors.append(f"{label}: expected {expected}, got {actual}")


def compare_texts(baseline: list[dict], candidate: list[dict], tolerance: float) -> list[str]:
    errors: list[str] = []
    if len(candidate) != len(baseline):
        errors.append(f"text count: expected {len(baseline)}, got {len(candidate)}")
        return errors

    for index, (base_el, cand_el) in enumerate(zip(baseline, candidate), start=1):
        base_text = str(base_el.get("text", ""))
        cand_text = str(cand_el.get("text", ""))
        if cand_text != base_text:
            errors.append(f"text #{index}: expected {base_text!r}, got {cand_text!r}")
            continue
        for key in ("left", "top", "width", "height"):
            diff_number(
                f"text #{index} {base_text!r} bbox.{key}",
                float(base_el["bbox"][key]),
                float(cand_el["bbox"][key]),
                tolerance,
                errors,
            )
        diff_number(
            f"text #{index} {base_text!r} font_size",
            float(base_el.get("font_size", 0)),
            float(cand_el.get("font_size", 0)),
            tolerance,
            errors,
        )
    return errors


def compare_font_stats(baseline_root: Path, candidate_root: Path, tolerance: float) -> tuple[dict, dict, list[str]]:
    base_stats = stat(pptx_font_sizes(baseline_root / "build" / "reconstructed.pptx"))
    cand_stats = stat(pptx_font_sizes(candidate_root / "build" / "reconstructed.pptx"))
    errors: list[str] = []
    for key in ("count", "avg", "median", "min", "max"):
        if base_stats[key] is None or cand_stats[key] is None:
            if base_stats[key] != cand_stats[key]:
                errors.append(f"pptx font {key}: expected {base_stats[key]}, got {cand_stats[key]}")
        elif key == "count":
            if base_stats[key] != cand_stats[key]:
                errors.append(f"pptx font count: expected {base_stats[key]}, got {cand_stats[key]}")
        else:
            diff_number(f"pptx font {key}", float(base_stats[key]), float(cand_stats[key]), tolerance, errors)
    return base_stats, cand_stats, errors


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Compare a slide2pptx output against a golden baseline.")
    parser.add_argument(
        "candidate",
        type=Path,
        help="Output directory containing detect/detected.json and build/reconstructed.pptx.",
    )
    parser.add_argument(
        "--baseline",
        type=Path,
        default=Path("tests/golden/two-pass-original-6b94462e"),
        help="Golden output directory. Defaults to tests/golden/two-pass-original-6b94462e.",
    )
    parser.add_argument("--tolerance", type=float, default=0.01, help="Numeric tolerance for bbox/font comparisons.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    baseline_root = args.baseline.resolve()
    candidate_root = args.candidate.resolve()

    base_data, base_texts = load_detected_texts(baseline_root)
    cand_data, cand_texts = load_detected_texts(candidate_root)

    errors: list[str] = []
    for key in ("element_count", "text_element_count", "visual_component_count"):
        expected = base_data.get("metrics", {}).get(key)
        actual = cand_data.get("metrics", {}).get(key)
        if actual != expected:
            errors.append(f"metrics.{key}: expected {expected}, got {actual}")

    errors.extend(compare_texts(base_texts, cand_texts, args.tolerance))
    base_font_stats, cand_font_stats, font_errors = compare_font_stats(
        baseline_root,
        candidate_root,
        args.tolerance,
    )
    errors.extend(font_errors)

    summary = {
        "ok": not errors,
        "baseline": str(baseline_root),
        "candidate": str(candidate_root),
        "baseline_metrics": base_data.get("metrics", {}),
        "candidate_metrics": cand_data.get("metrics", {}),
        "baseline_pptx_font_stats": base_font_stats,
        "candidate_pptx_font_stats": cand_font_stats,
        "errors": errors[:20],
        "error_count": len(errors),
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if not errors else 1


if __name__ == "__main__":
    sys.exit(main())
