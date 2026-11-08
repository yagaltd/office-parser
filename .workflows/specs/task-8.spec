spec: task
name: "page-verdicts"
tags: [pdf, quality, routing]
verify: "cargo test --features pdf --test pdf_quality_pages && cargo test --features pdf"
max-rounds: 3
---

## Intent

Wire per-page quality into the PDF parser: one `PageTextQuality` per page with a routing verdict, a
document-level rollup, and a JSON mirror in `metadata.extra`. The critical invariant is that
`Escalate` never routes a page to OCR — it is a vision-LLM signal, because OCR cannot repair reading
order.

### Diagrams

```mermaid
flowchart TD
    A[page] --> B[richness: PdfTextQuality]
    A --> C[geometry_confidence]
    A --> D[layout_complexity]
    A --> E[trust reasons]
    B --> F{verdict}
    C --> F
    D --> F
    E --> F
    F -->|text low or untrusted| G[Ocr]
    F -->|trusted and complex| H[Escalate]
    F -->|trusted and simple| I[Trust]
    G --> J[ocr_pages]
    H --> K[page_quality only]
    I --> K
```

## Decisions

- **`Escalate` never adds a page to `ocr_pages`.** Only `Ocr` does (plus the existing
  `min_chars` / trust-reason paths). This is asserted by a dedicated negative test.
- The document-level `PdfTextQuality` stays as the summary enum and keeps its current values; the
  rollup for *routing* is worst-page-wins, computed from `page_quality`.
- `--ocr-min-chars` default (50) and `PAGE_TEXT_THRESHOLD` are **unchanged**. New signals add
  `Escalate`; they never re-tune `Ocr`.
- The per-page array is mirrored into `metadata.extra["pdf_quality"]` so JSON consumers get detail
  without a struct change, alongside the existing `metadata.extra["ocr"]`.
- `parse_pdf_split` and `parse_with_page_ocr_opts` must report **identical** per-page verdicts for the
  same input. Today they diverge (`src/formats/pdf.rs:2703` sets a per-page enum independently).
- Page numbers in `PageTextQuality.page` are 1-indexed, matching `SpatialPage.page_number` convention
  and the existing `metadata.extra["ocr"].low_confidence_pages` usage.
- No new dependency; no pdfium requirement for verdicts to be populated.

## Boundaries

### Allowed Changes
- `src/formats/pdf.rs`
- `tests/pdf_quality_pages.rs` (new)

### Forbidden
- `src/quality.rs` — types and scorer are fixed by TASK 5 and TASK 7
- `src/document.rs`, `src/lib.rs`, `src/render.rs`
- Do not change `PAGE_TEXT_THRESHOLD` or the `--ocr-min-chars` default
- Do not change `PageOcrEngine` (that is TASK 10)
- Do not add pdfium, a model, or any dependency
- Do not remove or reword the existing `metadata.extra["ocr"]` keys

## Acceptance Criteria

Scenario: one quality entry per page
  Test: test_page_quality_len_matches_page_count
  Given `tests/fixtures/pdf_mixed.pdf`
  When parsed
  Then `metadata.page_quality` has exactly `page_count` entries
  And their `page` values are `1..=page_count` in order

Scenario: an Escalate page is never sent to OCR
  Test: test_escalate_page_is_not_routed_to_ocr
  Given a page whose text layer is trusted and whose `layout_complexity` is at or above threshold
  When the document is parsed with an OCR engine available
  Then that page's verdict is `Escalate`
  And the page index is **absent** from `metadata.extra["ocr"]`'s OCR'd page set
  And `pages_ocr` does not count it

Scenario: split-pages and combined parsing agree
  Test: test_split_pages_reports_same_verdicts
  Given `tests/fixtures/pdf_mixed.pdf`
  When parsed both via `parse_pdf_split` and via the combined path
  Then the per-page verdicts are identical element for element

Scenario: extra mirror matches the struct
  Test: test_pdf_quality_mirrored_in_metadata_extra
  Given any parsed PDF
  When `metadata.extra["pdf_quality"]` is compared with `metadata.page_quality`
  Then they describe the same pages with the same verdict kinds

Scenario: a trusted simple page stays Trust (negative)
  Test: test_trusted_simple_page_has_no_ocr_reason
  Given a single-column, text-rich, unrotated page
  When parsed
  Then its verdict is `Trust`
  And it carries no OCR reason and no escalation score above threshold

Scenario: document rollup prefers the worst page
  Test: test_document_rollup_uses_worst_page_for_routing
  Given a document where page 1 is `TextRich` and page 2 is `ImageOnly`
  When the rollup runs
  Then the routing rollup sends page 2 to OCR
  And the public `pdf_text_quality` enum keeps its existing document-level value
