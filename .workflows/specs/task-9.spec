spec: task
name: "no-ink-trust-check"
tags: [pdf, trust, geometry, scan-detection]
verify: "cargo test --features pdf --test pdf_no_ink && cargo test --no-default-features --features pdf --test pdf_no_ink"
max-rounds: 3
---

## Intent

Close the TODO at `src/formats/pdf.rs:2571` — `(no-ink check pending: needs char-box geometry.)` —
using the ink coverage that TASK 4 geometry makes available, so a page whose text layer is invisible
is untrusted **without requiring pdfium**. Today the trust verdicts silently degrade to an empty
vector when `render-pdfium` is off; after this task they are non-empty on geometry alone.

## Diagrams

```mermaid
flowchart TD
    A[page] --> B[char count]
    A --> C[glyph box area from PageGeometry]
    A --> D[page area from MediaBox/CropBox]
    C --> E[ink coverage = glyph area / page area]
    D --> E
    E --> F{coverage implausibly low<br/>for this char count?}
    F -->|yes| G[untrust: no-ink]
    F -->|no| H[no verdict from this check]
    I[pdfium render-mode-3 / bad chars] -.->|optional extra signal| G
```

## Decisions

- The no-ink check reads **only** TASK 4 geometry; `render-pdfium` is an additional signal where
  present, never a precondition.
- A verdict produced by this check carries a machine-readable reason string distinct from the existing
  pdfium reasons, so consumers can tell the sources apart. Proposed literal: `"no-ink"`.
- The threshold is a named constant with a doc comment stating it is a heuristic calibrated on the
  five fixtures — not a calibrated probability.
- The existing pdfium-based reasons are preserved verbatim where pdfium is available; this task adds
  a source, it does not replace one.
- `force` still bypasses every trust check, including this one.
- Verdicts must never be silently empty because a feature is off: if no signal can run, that is a
  test failure, not a degraded pass.
- No new dependency.

## Boundaries

### Allowed Changes
- `src/formats/pdf.rs`
- `tests/pdf_no_ink.rs` (new)

### Forbidden
- `src/quality.rs`, `src/document.rs`, `src/lib.rs`, `src/render.rs`
- Do not change `PAGE_TEXT_THRESHOLD` or `--ocr-min-chars`
- Do not change the existing pdfium reason strings
- Do not make `render-pdfium` a required feature for the no-ink verdict
- Do not touch `PageOcrEngine` (TASK 10) or `parse_pdf_split` behaviour
- Do not add a dependency

## Completion Criteria

Scenario: a scan is untrusted with pdfium disabled
  Test: test_scanned_page_untrusted_without_pdfium
  Given `cargo test --no-default-features --features pdf --test pdf_no_ink`
  And `tests/fixtures/pdf_scanned.pdf`
  When the page's trust verdict is computed
  Then it is `Some("no-ink")` (or the documented reason literal)
  And the verdict did not require `render-pdfium`

Scenario: a real text page is not falsely untrusted (negative)
  Test: test_text_page_remains_trusted
  Given `tests/fixtures/pdf_mixed.pdf`'s text-rich pages
  When the trust verdict is computed
  Then every such page yields no no-ink verdict
  And its `PageTextQuality.verdict` remains `Trust`

Scenario: verdicts are not silently empty when pdfium is absent (regression)
  Test: test_trust_verdicts_not_empty_without_pdfium
  Given the crate built without `render-pdfium`
  When the trust-check vector is produced for `pdf_scanned.pdf`
  Then it is not an empty vector
  And the previous silent-degradation path is asserted gone

Scenario: force bypasses the new check too
  Test: test_force_bypasses_no_ink_check
  Given a scanned fixture parsed with `force = true`
  When trust verdicts are computed
  Then no no-ink verdict is produced
  And OCR runs regardless of geometry

Scenario: a blank page does not divide by zero (edge case)
  Test: test_zero_page_area_does_not_panic
  Given a page whose MediaBox and CropBox both yield zero area
  When ink coverage is computed
  Then the result is a defined value (no NaN, no panic)
  And no trust verdict is derived from a zero-area page
