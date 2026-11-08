spec: task
name: "xy-cut-reading-order"
tags: [pdf, layout, reading-order]
verify: "cargo test --features pdf --test pdf_reading_order && cargo test --features pdf"
max-rounds: 2
---

## Intent

Order born-digital PDF content by recursive XY-Cut over the TASK 4 span boxes, so multi-column pages
stop interleaving their columns into a single stream. Single-column pages must produce byte-identical
output — the fix must be invisible where the current behaviour is already correct.

## Diagrams

```mermaid
flowchart TD
    A[page spans] --> B{vertical gap splits page?}
    B -->|yes| C[split into vertical regions]
    B -->|no| D{horizontal gap splits page?}
    D -->|yes| E[split into horizontal bands]
    D -->|no| F[emit spans in this block, top-to-bottom then left-to-right]
    C --> G[recurse per region]
    E --> G
    G --> H[ordered spans]
    H --> I[existing heading level walk]
    I --> J[Blocks]
```

## Decisions

- The cut runs on `SpanBox` from `src/formats/pdf.rs` (TASK 4), not on `Block` — blocks are built
  after ordering.
- Column detection and ordering are pure geometry: no model, no pdfium, no new dependency.
- **Single-column pages keep the current output exactly.** If the XY-Cut result for a single-column
  page differs from the linear order, the linear order wins for that page — the regression guard is a
  byte-identical assertion, not a tolerance.
- The existing heading `level` walk (`src/formats/pdf.rs:803-875`) runs **after** ordering. If it
  cannot survive reordering, stop and escalate: that is the TASK 2p NO-GO trigger, not a local
  workaround.
- Page furniture handling is out of scope; this task changes order only.
- `--split-pages` must observe the same order as the combined path.

## Boundaries

### Allowed Changes
- `src/formats/pdf.rs`
- `tests/pdf_reading_order.rs` (new)

### Forbidden
- `src/quality.rs` — owned by TASK 7 (parallel group B)
- `src/document.rs`, `src/lib.rs`
- Do not change `PdfTextQuality`, `PageGeometry`, or `SpanBox` shapes (TASK 4's contract)
- Do not change heading text, `level` values, or block variants
- Do not reorder pages — only content within a page
- Do not add a dependency

## Completion Criteria

Scenario: single-column output is byte-identical
  Test: test_single_column_output_unchanged
  Given `tests/fixtures/pdf_table_scan.pdf`'s text-layer pages and any single-column fixture
  When parsed with `cargo test --features pdf`
  Then `render::to_markdown` output equals the pre-change output byte for byte
  And the block count and every block's text are unchanged

Scenario: multi-column content is no longer interleaved
  Test: test_mixed_fixture_columns_not_interleaved
  Given `tests/fixtures/pdf_mixed.pdf`, whose page holds more than one column
  When the page is parsed
  Then a sentence that begins in the left column is not followed by a sentence from the right column
    before its own continuation
  And the assertion is made on specific known words from the fixture, not on a snapshot hash

Scenario: heading levels survive reordering (negative)
  Test: test_heading_levels_preserved_after_reorder
  Given a fixture with headings
  When the page is parsed with XY-Cut ordering
  Then every `Block::Heading` that existed before still exists with the same `level` and text
  And no heading is demoted to `Block::Paragraph`

Scenario: empty page does not panic (edge case)
  Test: test_empty_page_ordering_returns_no_blocks
  Given a page with no spans
  When ordering runs
  Then it produces zero blocks rather than panicking
