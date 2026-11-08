# Verdict — TASK 6 (worker): XY-Cut reading order

ok: true

Reviewer: orchestrator (gate), 2026-11-08.

Evidence:
- Applied: XY-Cut section (xy_cut_reading_regions / xy_cut_partition /
  xy_cut_leaf / xy_split_at_widest_gap, depth-capped), EDIT A/B doc updates,
  EDIT D integration in parse_pdf_full. tests/pdf_reading_order.rs: **9/9 pass**.
- KEY DESIGN FINDING accepted: per-region clustering at BOTH stages (lines +
  paragraphs) — the paragraph stage re-sorts lines by y globally, so atom
  reordering alone would re-interleave columns. No-cut pages run the identical
  original two calls (byte-guard structural).
- Synthetic two-column end-to-end: interleaved stream → left column then right,
  byte-pinned markdown; provably non-vacuous (fixture self-check pins the
  interleaved stream order). Pure-geometry tests: full-width header band peel,
  narrow-gap threshold, leaf geometric ordering, single-region no-op.
- Gate fixes (test expectations, not implementation):
  1. "empty page" assertion: an image-only page emits exactly one Image block
     (embedded-scan extraction) — expectation corrected from "no blocks"
     (rot90 precedent: same behavior, scout-verified).
  2. fn-pointer coercion for the axis closure pair (compile fix).
- Spec premise correction (recorded, mirrors TASK 4's): pdf_mixed is
  single-column per TASK 3 measurement — the genuine multi-column path is
  synthetic-verified; corpus behavior is a structural no-op byte-guard.
- Full matrix green (69 passed / 0 failed with pdf), default + rten configs 0
  failures, clippy clean.

TASK 8 note: pages_geometry + xy_cut regions + per-page quality (TASK 7) are
all in place for the verdicts/rollup task.
