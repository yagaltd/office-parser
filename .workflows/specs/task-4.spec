spec: task
name: "page-geometry"
tags: [pdf, geometry, extraction]
verify: "cargo test --features pdf --test pdf_geometry && cargo test --features pdf"
max-rounds: 2
---

## Intent

Emit per-page geometric truth from the PDF content interpreter — every span's box, font name and
size, plus the page's MediaBox, CropBox (origin subtracted), native rotation and coordinate system —
so that reading order, ink coverage and citation all have a substrate. Strictly additive: no
existing block, field, or output byte changes.

## Diagrams

```mermaid
flowchart LR
    A[content stream] --> B[interpreter state: tm, font_size]
    B --> C[SpanBox x0 y0 x1 y1 font_name font_size]
    C --> D[PageGeometry]
    E[MediaBox / CropBox / Rotate] --> D
    D --> F[ParsedPdfDocument.pages_geometry]
    D --> G[coordinate_system: source | viewer_oriented]
```

## Decisions

- `PageGeometry` and `SpanBox` are defined **in `src/formats/pdf.rs`** (PDF-scoped types, `pub`), not
  in `src/document.rs` — that file belongs to TASK 5 and must not be touched here. Downstream tasks
  (`src/quality.rs`) import these types read-only.
- Boxes are in PDF points, origin top-left of the **visible** page: CropBox when present, MediaBox
  otherwise, with the crop origin subtracted (`norm_x = x - origin_x`, `norm_y = max_y - y`).
- `initial_rotation` is the page's native `/Rotate` value and is preserved even when
  `coordinate_system` already reports `viewer_oriented`.
- `coordinate_system` starts as `"source"` in this task; viewer-oriented normalisation is not in scope.
- `char_count` per span uses `estimate_text_advance()` for advance where the font metrics are
  unavailable — the same helper the current text path already uses (`src/formats/pdf.rs:471`).
- The existing `Block` output is untouched: the same `blocks` vector, the same text, the same order.

## Boundaries

### Allowed Changes
- `src/formats/pdf.rs`
- `tests/pdf_geometry.rs` (new)

### Forbidden
- `src/document.rs` — owned by TASK 5 (parallel group A)
- `src/quality.rs`, `src/lib.rs` — owned by TASK 5
- `src/formats/indra_native.rs`, `src/formats/pdf_ocr.rs`
- Do not change any existing `Block` variant, field, text, or ordering
- Do not add a dependency
- Do not add pdfium or hayro rendering

## Completion Criteria

Scenario: every span carries a box, font name and font size
  Test: test_spans_carry_bbox_font_name_and_size
  Given `tests/fixtures/pdf_mixed.pdf` parsed with `cargo test --features pdf`
  When the page geometry for page 1 is inspected
  Then it holds at least one `SpanBox`
  And every span has `x1 > x0`, `y1 > y0`, a non-empty `font_name` and `font_size > 0.0`

Scenario: crop-box origin is subtracted
  Test: test_cropbox_origin_is_subtracted
  Given a fixture whose CropBox origin is non-zero
  When the page geometry is built
  Then the minimum `x0` across spans is `>= 0.0`
  And the stored `crop_box` records the original origin so the translation is reversible

Scenario: native rotation is preserved
  Test: test_initial_rotation_preserved
  Given `tests/fixtures/pdf_rot90_scan.pdf`, whose page carries `/Rotate`
  When the page geometry is built
  Then `initial_rotation` is `90` (or `270` for a 270-page), not `0`
  And `coordinate_system` is `"source"`

Scenario: existing block output is byte-identical (negative)
  Test: test_existing_block_output_unchanged
  Given the same five fixtures parsed before and after this change
  When `render::to_markdown` and `render::to_json_value` output is compared
  Then it is identical for all five fixtures
  And `Document.blocks` count is unchanged

Scenario: an empty page does not panic (edge case)
  Test: test_page_without_spans_yields_empty_geometry
  Given a page that emits no text operators
  When the page geometry is built
  Then it returns an empty span vector rather than panicking or omitting the page
