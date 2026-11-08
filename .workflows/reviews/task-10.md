# Verdict — TASK 10 (worker): OcrPage + defaulted ocr_page()

ok: true

Reviewer: orchestrator (gate), 2026-11-08.

Evidence:
- Applied: OcrWordBox / OcrLayoutRegion / OcrPage (serde, skip_serializing_if
  on Options) + the DEFAULTED `ocr_page()` wrapper on `PageOcrEngine` —
  object-safe, no generics, legacy implementors compile unchanged (the router
  calls `ocr_page()`; IndraNative + all test engines run via the default
  wrapper).
- Per-page payloads threaded into `metadata.extra["ocr"]` as
  layout_regions / word_boxes / table_html records ({"page": n, ...});
  keys ABSENT when no OCR'd page supplied the field (never null).
- tests/ocr_page_contract.rs: **3/3 pass** — legacy-engine compatibility,
  object safety, optional-fields threading (None keys absent).
- The behavior lock is untouched: `git diff tests/pdf_ocr.rs` empty; the
  pdf,ocr-indra-native matrix is green (98 passed / 0 failed).
- ocr_page_png signature unchanged; no new dependency; raster-pixel
  provenance for box coordinates documented in rustdoc.
