from pathlib import Path

import pytest
from PIL import Image, ImageDraw

from slide2pptx.detect import core


@pytest.mark.skipif(core.cv2 is None, reason="opencv-python is not installed")
def test_detect_exports_visual_components(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(core, "RapidOCR", None)
    img = Image.new("RGB", (1280, 720), "#F7FAFF")
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle((120, 140, 360, 300), radius=28, fill="#2563EB")
    draw.ellipse((760, 180, 860, 280), fill="#F97316")
    source = tmp_path / "slide.png"
    img.save(source)

    result = core.detect(source, tmp_path / "detect")

    elements = result.payload["elements"]
    visual = [
        el
        for el in elements
        if el.get("metadata", {}).get("detector") == "opencv_residual_components"
    ]
    assert result.payload["background"]["strategy"] == "cleaned"
    assert result.payload["metrics"]["visual_component_count"] >= 1
    assert len(visual) >= 1
    assert all(el["kind"] in {"shape", "image"} for el in visual)
    for el in visual:
        if el["kind"] == "image":
            assert Path(el["image_path"]).is_file()


@pytest.mark.skipif(core.cv2 is None, reason="opencv-python is not installed")
def test_text_like_regions_are_not_merged_into_visual_components(tmp_path: Path) -> None:
    img = Image.new("RGB", (900, 420), "#FFFFFF")
    draw = ImageDraw.Draw(img)
    draw.ellipse((70, 130, 170, 230), fill="#0B4EA2")
    draw.text((205, 140), "Customer", fill="#0B4EA2")
    draw.text((205, 172), "Needs and value", fill="#111827")
    source = tmp_path / "slide.png"
    img.save(source)

    text_boxes = core._detect_text_like_regions(img)
    assert text_boxes, "heuristic text boxes should be available as a fallback mask"

    visual, _masks, _warnings = core._extract_visual_components(
        img,
        tmp_path / "detect",
        text_boxes,
        max_components=8,
    )

    assert visual
    icon_region_right_px = 190
    scaled_limit = icon_region_right_px * (core.SLIDE_WIDTH / img.width)
    assert all(
        el["bbox"]["left"] + el["bbox"]["width"] < scaled_limit
        for el in visual
    ), "visual extraction should keep adjacent text out of icon/image components"


@pytest.mark.skipif(core.cv2 is None, reason="opencv-python is not installed")
def test_iterative_residual_components_are_tagged_as_second_pass(tmp_path: Path) -> None:
    img = Image.new("RGB", (900, 420), "#FAFCFF")
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle((60, 60, 840, 110), radius=18, fill="#0B4EA2")
    draw.rectangle((210, 260, 760, 264), fill="#9DB9E9")
    draw.ellipse((120, 220, 170, 270), fill="#0B4EA2")
    draw.polygon([(500, 215), (560, 245), (500, 275)], fill="#0B4EA2")

    elements, masks, warnings = core._extract_iterative_residual_components(
        img,
        tmp_path / "detect",
        max_components=16,
    )

    assert not warnings
    assert elements
    assert len(masks) == len(elements)
    assert all(
        el["metadata"]["detector"] == "iterative_residual_components"
        and el["metadata"]["pass"] == 2
        for el in elements
    )
    assert all(Path(el["image_path"]).is_file() for el in elements)


def test_unpack_current_rapidocr_output() -> None:
    class Output:
        boxes = [
            [[10.0, 20.0], [110.0, 20.0], [110.0, 50.0], [10.0, 50.0]]
        ]
        txts = ("中文标题",)
        scores = (0.98,)

    boxes, texts, scores = core._unpack_ocr_result(Output())

    assert len(boxes) == 1
    assert texts == ["中文标题"]
    assert scores == [0.98]


def test_unpack_legacy_rapidocr_output() -> None:
    box = [[10.0, 20.0], [110.0, 20.0], [110.0, 50.0], [10.0, 50.0]]
    boxes, texts, scores = core._unpack_ocr_result(([[box, "标题", 0.95]], 0.1))

    assert boxes == [box]
    assert texts == ["标题"]
    assert scores == [0.95]
