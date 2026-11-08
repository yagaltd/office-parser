# Verdict — TASK 7 (worker, RE-SCOPED): per-page quality + routing

ok:true

Reviewer: orchestrator (gate), 2026-11-08.

Evidence:
- Applied: quality.rs (page_verdict gate order + count_columns), pdf.rs (per-page
  quality population in plain + OCR paths, stale page_block_ranges fix via
  `ParsedPdfDocument::recompute_page_block_ranges` incl. metadata_json copy),
  tests/quality_routing.rs (7 tests).
- All 7 pass; full matrix green. Routing table matches the spike: 5 image-only
  pages → Ocr(image-only), pdf_cjk_scan p1 → Ocr via char threshold, pdf_mixed p1
  → Trust (columns=1, not rotated).
- Gate deviation (recorded): the delivered count_columns used greedy y-overlap line
  grouping, which merged same-row two-column spans into full-width lines and masked
  the gutter (the worker's own synthetic test failed; the fn's doc comment had
  flagged exactly this limitation). Orchestrator replaced it with a span-level
  union variant (full-width spans excluded, uncovered strips = gutters) — the
  worker's synthetic two-column test now passes at the asserted 2, the known
  limitation is gone, and all fixture verdicts are unchanged.
- Intentional deviation recorded: pdf.rs edits follow the TASK 3 re-scope brief
  (pre-rescope spec listed pdf.rs as forbidden).
- Scout defect closed: page_block_ranges recomputed after the OCR merge
  (combined + metadata_json copies), covered by two dedicated tests.

Docs impact for TASK 12: docs/formats.md PDF section (per-page page_quality
populated; split pages carry single-page records; page_block_ranges recomputed
post-OCR) + CHANGELOG entry with Plan ID.
