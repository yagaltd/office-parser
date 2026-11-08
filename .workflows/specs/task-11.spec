spec: task
name: "rusto-core-tables"
tags: [pdf, ocr, feature-flag, tables]
verify: "cargo build && cargo test --features pdf,ocr-indra-native,ocr-rusto-core --test ocr_rusto_tables"
max-rounds: 3
---

## Intent

Add an **off-by-default** `ocr-rusto-core` feature that fills the new `OcrPage` optional fields using
`rusto-core` — layout regions, word boxes and table structure — for `LayoutType::Table` regions only.
Indra keeps the text path, so transcripts must be identical with the feature on and off.

## Diagrams

```mermaid
flowchart TD
    A[rendered page PNG] --> B[Indra: det + rec + order]
    B --> C[text blocks + confidence]
    A --> D{ocr-rusto-core enabled?}
    D -->|no| E[OcrPage with None extras]
    D -->|yes| F[rusto-core LayoutDetector]
    F --> G[LayoutType::Table regions]
    G --> H[rusto-core table structure]
    H --> I[table_html]
    F --> J[word boxes]
    C --> K[OcrPage]
    E --> K
    I --> K
    J --> K
```

## Decisions

- **`rusto-core` only.** Never the `rusto-rs` facade crate, whose `default = ["ffi", "rten"]` pulls in
  `uniffi`, `cbindgen`, `ureq` and `clap`.
- Feature name is `ocr-rusto-core`, **not** in `default`. With the feature off the dependency is absent
  from the resolved graph (`cargo tree` proves it; `Cargo.lock` is gitignored in this repo and cannot
  be the evidence).
- Indra recognition is untouched: the same det/rec/order models, the same transcripts. Any transcript
  difference is a failure, not a tuning opportunity.
- Scope is deliberately narrow: layout regions, word boxes, table structure for `LayoutType::Table`
  only. No unwarp, no layout-driven re-OCR, no model substitution.
- New module `src/formats/rusto_tables.rs`, gated `#[cfg(feature = "ocr-rusto-core")]`, registered in
  `src/formats/mod.rs` behind the same cfg.
- If `rusto-core` cannot be vendored under MIT-compatible terms as expected, stop and escalate — the
  AGPL ruling in `.workflows/CONTEXT.md` §Rulings governs the whole plan.

## Boundaries

### Allowed Changes
- `Cargo.toml` — the feature plus one optional dependency
- `src/formats/rusto_tables.rs` (new)
- `src/formats/mod.rs` — cfg-gated module registration
- `src/formats/pdf.rs` — only the `ocr-rusto-core` side of the `OcrPage` construction
- `tests/ocr_rusto_tables.rs` (new)

### Forbidden
- `src/formats/indra_native.rs` — recognition must not change
- `tests/pdf_ocr.rs` — the 6 existing OCR tests stay unmodified
- Do not add `rusto-rs`, `rusto-nn`, or any other rusto crate
- Do not put the feature in `default`
- Do not replace or re-tune Indra detection, recognition, or ordering
- Do not add pdfium or hayro

## Completion Criteria

Scenario: the feature is off by default and pulls nothing
  Test: `! cargo tree --no-default-features --features pdf -e normal | grep -q rusto-core`
  Given a default build in the repo root
  When `cargo build` runs and the dependency tree is inspected
  Then `rusto-core` is absent from the resolved graph
  And `cargo build` succeeds with no new dependency

Scenario: transcripts are identical with the feature on and off
  Test: test_transcripts_identical_with_feature_on_and_off
  Given `tests/fixtures/pdf_scanned.pdf`
  When OCR'd once with `ocr-rusto-core` off and once with it on
  Then the concatenated block text is identical in both runs
  And the per-page confidence values are identical

Scenario: table markup appears on the table fixture
  Test: test_table_html_populated_for_table_scan_fixture
  Given `tests/fixtures/pdf_table_scan.pdf` with the feature enabled
  When the page is OCR'd
  Then `OcrPage.table_html` is `Some` and contains a `<table` element
  And the engine with the feature disabled returns `None` for that field

Scenario: recognition path is untouched (negative)
  Test: test_indra_recognition_source_unchanged
  Given the repo state after this task
  When `git diff --name-only src/formats/indra_native.rs` is executed
  Then the output is empty
  And `cargo test --features pdf,ocr-indra-native` passes
