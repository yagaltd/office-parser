---
spec: task
name: "pdf-columns-routing"
tags: [worker, pdf, quality, rescope]
verify: "cargo test; tests/quality_routing.rs asserts per-page routing verdicts on the five fixtures"
max-rounds: 2
---

## Intent

**RE-SCOPED by TASK 3 verdict (deletion-test branch 1, 2026-11-08):** the five-term
`layout_complexity` score is DROPPED — measured `gutters = 0` on every text-layer page and the
score agreed with `columns` on every page of the corpus (`.workflows/spikes/pdf-complexity/README.md`
§Measured verdict).

Implement instead:

1. `columns` per text-layer page: horizontal gutter count over body-line x-intervals
   (width ≥ max(2×median font size, 12pt)) — the one surviving signal (measured: gutters=0,
   columns=1 on all our pages).
2. `pdf_text_quality` routing unchanged (ImageOnly → `PageVerdict::Ocr { reason }`).
3. Standalone rotated-page flag read directly off the Tm matrix (dominant tilt > 5° when tilted
   atoms carry ≥50% of chars) — escalates rotation without span geometry.

## Completion Criteria

- `PageTextQuality` carries `columns` + `rotated` (per TASK 5's types); verdict logic:
  ImageOnly → Ocr; rotated → Escalate; else Trust.
- `tests/quality_routing.rs`: five-fixture routing table matches
  `.workflows/spikes/pdf-complexity/README.md` §Score table (ImageOnly pages → Ocr;
  text-layer pages → Trust; rotated fixture is image-only so Ocr wins).
- `cargo test` green.

## Allowed Changes

- `src/quality.rs`, `src/document.rs` (verdict fn), `tests/quality_routing.rs`

Forbidden: `src/formats/pdf.rs` (TASK 4/6 boundary), the five-term score formula.
