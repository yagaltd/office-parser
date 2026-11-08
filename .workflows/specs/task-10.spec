spec: task
name: "ocr-page-contract"
tags: [pdf, ocr, api, trait]
verify: "cargo test --features pdf,ocr-indra-native --test ocr_page_contract && cargo test --features pdf,ocr-indra-native"
max-rounds: 2
---

## Intent

Widen the OCR seam so a page can carry layout regions, word boxes and table markup in addition to
text — **without breaking a single existing implementor**. `OcrPage` plus a defaulted
`PageOcrEngine::ocr_page()` does that: the new fields are `None` for any engine that only implements
today's `ocr_page_png()`.

## Diagrams

```mermaid
flowchart TD
    A[PageOcrEngine] --> B[ocr_page_png: png to blocks + conf<br/>EXISTING, unchanged]
    A --> C[ocr_page: png to OcrPage<br/>NEW, defaulted]
    C -->|default impl| B
    B --> D[blocks]
    B --> E[confidence]
    C --> F[layout_regions: Option]
    C --> G[word_boxes: Option]
    C --> H[table_html: Option]
    F --> I[metadata.extra ocr]
    G --> I
    H --> I
```

## Decisions

- `ocr_page()` is a **defaulted** trait method wrapping `ocr_page_png()`; it is not a breaking change
  and implementors are not required to override it.
- `OcrPage` optional fields default to `None`. An engine that supplies them is opt-in.
- `ocr_page_png()` is retained unchanged — removing it would break every downstream implementor and
  is explicitly out of scope.
- The trait must stay object-safe (`&mut dyn PageOcrEngine` is used at
  `src/formats/pdf.rs:2457`). The defaulted method may not introduce a generic parameter.
- New non-`None` fields are threaded into `metadata.extra["ocr"]`; existing keys are preserved.
- `tests/pdf_ocr.rs` (6 tests) must stay green **and unmodified** — it is the behaviour lock.
- No new dependency.

## Boundaries

### Allowed Changes
- `src/formats/pdf.rs`
- `tests/ocr_page_contract.rs` (new)

### Forbidden
- `tests/pdf_ocr.rs` — must remain byte-identical and green (the 6 existing tests are the guard)
- `src/quality.rs`, `src/document.rs`, `src/lib.rs`, `src/render.rs`
- Do not remove or change the signature of `ocr_page_png`
- Do not add a generic parameter to the trait or to `ocr_page`
- Do not add a dependency (the rusto-core path is TASK 11)

## Completion Criteria

Scenario: a legacy engine still compiles and runs
  Test: test_legacy_engine_implementing_only_ocr_page_png_still_works
  Given a test-local struct implementing only `ocr_page_png`
  When it is used through `&mut dyn PageOcrEngine` and `ocr_page()` is called
  Then it returns an `OcrPage` whose three new fields are all `None`
  And its blocks and confidence match what `ocr_page_png` returned

Scenario: the trait is still object-safe
  Test: test_page_ocr_engine_is_object_safe
  Given the crate compiled with `pdf,ocr-indra-native`
  When a `&mut dyn PageOcrEngine` is constructed and `ocr_page()` is invoked on it
  Then it compiles and runs
  And no `where Self: Sized` bound is present on the new method

Scenario: supplied optional fields reach metadata (negative-ish, positive path)
  Test: test_optional_fields_threaded_into_ocr_metadata
  Given a test engine that returns `Some(table_html)` and `Some(word_boxes)`
  When a scanned fixture is parsed through it
  Then `metadata.extra["ocr"]` exposes the table markup and word boxes
  And an engine returning `None` for them leaves those keys absent rather than null

Scenario: the existing OCR behaviour lock is unmodified
  Test: `cargo test --features pdf,ocr-indra-native --test pdf_ocr`
  Given the repo state after this task
  When the existing OCR test file runs
  Then all 6 tests pass
  And `git diff --name-only tests/pdf_ocr.rs` is empty
