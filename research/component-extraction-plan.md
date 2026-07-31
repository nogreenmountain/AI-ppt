# Component Extraction Plan

## Goal

Improve the slide image to editable PPTX pipeline so non-text visual parts are not only baked into the background. The new detector should split out logos, icons, decorations, simple panels, and illustration fragments as independent PPT elements while preserving the existing OCR text reconstruction flow.

## Open-Source Research Summary

The most useful reference direction is not a single end-to-end "perfect image to PPT" repo. The best route is a hybrid, borrowing ideas from several open-source families:

- `ningzimu/image-to-editable-ppt-skill`: best product-level reference for image-to-editable-slide workflow. The useful idea is a layered reconstruction contract: preserve visual fidelity with an image background, then progressively lift trusted objects into editable layers.
- `LayoutParser`: useful as a future layout-analysis reference, but too document-centric and heavy for this MVP. It is better suited when we want semantic regions such as title, figure, table, and caption.
- `facebookresearch/segment-anything`: useful as an optional advanced segmentation path for complex icons or illustrations. It should not be mandatory for the MVP because model size and prompt selection can introduce unstable results.
- `visioncortex/vtracer` / `potrace`: useful for future vectorization of small monochrome icons. Not suitable as the first implementation because SVG-to-PPT editable path handling is more fragile than transparent PNG overlays.
- OpenCV contours / connected components: best MVP base. It is deterministic, local, already aligned with the current `cv2` optional dependency, and can reuse OCR boxes plus inpainting masks.

Selected approach for this project: OpenCV residual connected-component extraction, with conservative filtering and schema-compatible output.

## Implementation Chosen

The new detector lives inside `python/slide2pptx/detect/core.py`.

Data flow:

1. Load input image.
2. Run OCR as before, producing native text elements and text masks.
3. Build a residual visual mask with OpenCV:
   - Canny edges for outlines.
   - Local contrast for non-background visual blocks.
   - Saturation/contrast mask for colored logos and decorations.
   - Subtract inflated OCR boxes so text remains text, not image crops.
4. Run connected components on the residual mask.
5. Filter dangerous candidates:
   - too small,
   - too large,
   - near full-page panels,
   - extreme aspect ratios,
   - text-like long sparse components,
   - heavy overlap with OCR boxes.
6. For simple, uniform components, emit native PPT `shape` elements.
7. For complex components, export transparent PNG crops into `detect/assets/` and emit `image` elements.
8. Inpaint both OCR masks and visual component masks into `cleaned-background.png` so the exported elements do not duplicate background pixels.

Schema impact:

- No new top-level fields.
- Existing `elements[]` now may include more `shape` and `image` entries.
- Added metrics:
  - `text_element_count`
  - `visual_component_count`
- Added per-component metadata:
  - `metadata.detector = "opencv_residual_components"`
  - `metadata.source_bbox_px`
  - `metadata.source_area_px`

## Why This Avoids Conflict

- OCR remains authoritative for text.
- Visual extraction runs after OCR box collection and subtracts text masks.
- Background cleaning now removes both OCR and visual components, so PPT layers do not double-render.
- The build step already supported `shape` and `image`, so no new renderer contract was needed.
- If OpenCV is missing, the feature degrades gracefully and existing behavior is preserved.

## Current Limitations

- Without OCR installed, text-like regions can still be ambiguous. A conservative text-like component filter was added to avoid exporting large text lines as images.
- Complex chart reconstruction is not yet semantic. Charts are only movable image components, not native PowerPoint charts.
- SVG/vector editing is not implemented yet. Transparent PNG extraction is the reliable MVP.
- SAM or other segmentation models are intentionally not mandatory.

## Verification

Commands run:

```powershell
pytest tests -q
```

Result:

```text
82 passed, 1 skipped
```

```powershell
node src\convert.mjs --self-test
```

Result:

```text
[self-test] passed
```

```powershell
python -m slide2pptx.pipeline_cli samples\source.png --out outputs\component-extraction-pipeline --skip-report
```

Result:

- Detection succeeded.
- Build succeeded.
- Output PPTX: `outputs/component-extraction-pipeline/build/reconstructed.pptx`
- Output preview: `outputs/component-extraction-pipeline/build/artifact-preview.png`

## MiniMax Plugin Usage Notes

What worked:

- `headroom doctor` showed Headroom running and reusable.
- A MiniMax task that passed the router gate completed successfully.
- The run used Headroom:
  - `headroom.enabled = true`
  - `status = reused`
  - `baseUrl = http://127.0.0.1:19244`
- The JSON report included useful run metadata such as duration, token counts, and `reviewStatus`.

Issues observed:

- `route --task-file` on a `{"tasks":[...]}` run-many file misclassified the batch as a code-edit task with missing scope/test command.
- Read-only research tasks containing the word `image` were blocked by the safety gate as Codex-owned, even though they were safe bounded research tasks.
- The accepted MiniMax worker could not read files under `C:\Users\tangvx\Documents\AI ppt`; it reported permission errors even though Codex itself had access. The worker output was still useful as a generic implementation checklist, but not as a real code inspection.
- Headroom reported `tokensSaved = 0` for this tiny task. That is understandable, but a short explanation in router output would help users interpret it.
- The Node child process emitted a `[DEP0190]` warning about passing args with `shell: true`. This is not a task failure, but it is worth cleaning up in the plugin.

Suggested plugin improvements:

- Add first-class support for route-checking each item inside a `tasks[]` run-many file.
- Make the `image` keyword gate more nuanced: "generate/edit image" should stay Codex-owned, but "read-only research about image processing libraries" should be delegateable.
- Ensure Claude/MiniMax workers inherit the same workspace read permissions as Codex, or surface the missing permission before spending a full worker run.
- Include a short Headroom interpretation field, e.g. why compression saved zero tokens.
- Avoid `shell: true` arg concatenation warnings where possible.
