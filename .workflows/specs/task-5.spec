spec: task
name: "quality-types"
tags: [pdf, quality, api, serde]
verify: "cargo test --test quality_types && cargo test"
max-rounds: 2
---

## Intent

Define the additive public surface that replaces a single document-wide enum with per-page detail:
`PageTextQuality` and `PageVerdict`, plus `DocumentMetadata.page_quality`. Types only — no wiring
into the PDF parser, no behaviour change. This is the shape the rest of the plan fills in.

## Diagrams

```mermaid
flowchart LR
    A[DocumentMetadata] --> B[pdf_text_quality: Option of PdfTextQuality<br/>UNCHANGED]
    A --> C[page_quality: Option of Vec of PageTextQuality<br/>NEW]
    C --> D[PageTextQuality<br/>page, text_richness,<br/>geometry_confidence,<br/>layout_complexity, verdict, columns]
    D --> E[PageVerdict]
    E --> F[Trust geometry]
    E --> G[Ocr reason]
    E --> H[Escalate complexity]
```

## Decisions

- `PdfTextQuality` is **unchanged** (variants, thresholds, doc comments). New surface is additive;
  existing 0.2.0 consumers keep compiling.
- `page_quality` is `Option<Vec<PageTextQuality>>` and is `None` for **every** non-PDF format,
  mirroring `pdf_text_quality` (`src/document.rs:35`).
- `PageVerdict` is internally tagged on `kind` with values `trust` | `ocr` | `escalate`, so the JSON
  shape is self-describing and extensible.
- `Escalate` and `Ocr` are distinct variants and must remain so: OCR cannot fix reading order
  (`.workflows/CONTEXT.md` §Domain Rules).
- `layout_complexity` and `geometry_confidence` are `f32` in `0.0..=1.0`; the type does not enforce
  the range (that is the scorer's job), but the field docs state it.
- `columns: u8` is the detected column count; `1` for single-column pages.
- Types live in a new `src/quality.rs`, re-exported from `src/lib.rs`. `src/formats/pdf.rs` is
  **forbidden** in this task — that file belongs to the parallel TASK 4.
- No new dependency.

## Boundaries

### Allowed Changes
- `src/quality.rs` (new)
- `src/document.rs` — add the `page_quality` field only
- `src/lib.rs` — `pub mod quality;` and the re-export
- `tests/quality_types.rs` (new)

### Forbidden
- `src/formats/pdf.rs` — owned by TASK 4 (parallel group A)
- `src/render.rs` — JSON shape is exercised through it, not changed by it
- Do not change `PdfTextQuality`, its variants, or its doc comments
- Do not populate `page_quality` anywhere in this task — it stays `None` everywhere
- Do not add a dependency

## Completion Criteria

Scenario: non-PDF formats report no page quality
  Test: test_page_quality_is_none_for_docx
  Given `tests/fixtures/docx_headings_lists_tables_images.docx` parsed
  When `doc.metadata.page_quality` is read
  Then it is `None`
  And `pdf_text_quality` is also `None` (parity with existing behaviour)

Scenario: the new surface round-trips through the existing JSON renderer
  Test: test_page_quality_serde_roundtrip
  Given a `Document` whose metadata carries two `PageTextQuality` values, one per verdict kind
  When `render::to_json_value(&doc)` is called
  Then both entries appear under `page_quality`
  And each verdict carries its `kind` discriminator
  And the value deserialises back into an equal `DocumentMetadata`

Scenario: verdict variants are distinguishable in JSON (negative)
  Test: test_escalate_and_ocr_are_distinct_in_json
  Given a `PageVerdict::Ocr { reason }` and a `PageVerdict::Escalate { complexity }`
  When both are serialised
  Then their `kind` values differ (`ocr` vs `escalate`)
  And an `Escalate` value contains no OCR reason field

Scenario: existing public API is untouched (regression)
  Test: test_pdf_text_quality_variants_unchanged
  Given the crate compiled with default features
  When `PdfTextQuality::{TextRich, TextSparse, ImageOnly}` are matched exhaustively
  Then all three variants still exist with their original meaning
  And `cargo test` passes without modifying any existing test
