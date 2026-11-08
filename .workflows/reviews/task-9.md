# Verdict — TASK 9 (worker): pdfium-free no-ink trust check

ok:true

Reviewer: orchestrator (gate), 2026-11-08.

Evidence:
- Applied: quality.rs — `NO_INK_*` constants (documented heuristics) +
  `no_ink` implausibility check + `median_font_size` helper (shared with
  count_columns); pdf.rs — `build_page_quality` gains the geometry trust
  source with a `trust_checks: bool` param (`force` bypasses ALL trust
  checks per spec; plain parse passes true), overriding rich-but-boxless
  pages to `Ocr { reason: "no-ink" }`; tests/pdf_no_ink.rs: **10/10 pass**
  on BOTH `--features pdf` and `--no-default-features --features pdf`
  (the spec's silent-degradation defect is closed: trust verdicts are
  never an empty vector — `trust_verdicts_are_never_silently_empty_without_pdfium`
  walks all five fixtures).
- Regressions: none — full matrix 89/0; cjk keeps its sparse reason;
  scanned stays "image-only" (gate order: richness 0 → image-only before
  no-ink); pdf_mixed p1 stays Trust (measured ratio ≈ 1.31 ≥ 0.25).
- Gate fixes (delivery bugs, recorded):
  1. EDIT 6's whole-fn replacement was applied with a TEXT-ANCHOR boundary
     (the delivered brace-scan approach breaks on braces inside format
     strings — the first application deleted 3,237 lines; restored from
     HEAD and re-applied safely).
  2. EDIT 7's NEW block was identical to its OLD block (missing `!force`)
     — the worker's own summary stated `!force`; intent applied.
- Self-consistent small boxes (uniform 1pt spans) are NOT flagged — the
  check detects char-count-vs-geometry contradiction; pdfium render-mode-3
  remains the invisible-but-boxed detector (additional signal where
  render-pdfium is on; no regression).
