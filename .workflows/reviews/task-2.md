# Verdict — TASK 2 (worker): upgrade-09 (hayro 0.7.x → 0.9.0)

ok:true

Reviewer: orchestrator (gate), 2026-11-09. Two-stage delivery: the wave-1
read-only worker's fence set, applied by the gate, then a write-enabled
worker (T2b, isolated worktree branch 4a4503f, merged --no-ff 8f8f994)
that completed the F7r section AND repaired the gate's mechanical splices
(device.rs trait rebuild, form.rs/color-mod ink_tints+cmyk_to_srgb8,
ocg.rs /VE evaluator placement, context.rs InterpreterCache/budget,
codec API adaptations for the crates.io 0.5.0/0.4.0/0.2.0 generation,
lzw_flate/xref/stream/page splice fixes, zip 2.1.3 lockfile pin).

Evidence:
- vendor/ exact set: hayro 0.9.0, hayro-syntax 0.9.0, hayro-interpret
  0.9.0, hayro-cmap 0.1.1, hayro-font 0.5.0 (stale jbig2/postscript/vello
  dirs deleted — wiring discrepancy resolved per decision 1(b)).
- Suites: 90/0 pdf,render-hayro (incl. test_hayro_renders_fixtures_non_blank),
  98/0 pdf,ocr-indra-native, 89/0 default, 89/0 pdf; clippy 0 errors on the
  four-feature union; glue grep clean (no render_size/render_into remnants).
- **Transcript re-baseline: NOT NEEDED** — zero fixture expectation edits;
  0.9.0 renders kept every assertion green (decision 3a provision unused).
- cargo tree: 0 rusto crates; hayro-font present in the render-hayro graph.
- Gate-side fixes recorded: hayro-ccitt?/std dropped (published 0.5.0 has
  no features table), merged over the worker branch cleanly.

## Review addendum (bug-hunter-002, 2026-11-09)

10 findings on the 002 diff — **all fixed pre-ship, committed c07a231**:
P1-1 knockout-erase re-entrancy (one-line gate), P2-1 CMYK ink
unsupported-on-image/pattern, P2-2 straight-alpha metadata, P2-3 cmap
section resync (5 parsers), P2-4 dead pdfium sniff marker, P3-1 per-axis
clamp, P3-3 device-space glyph extent, P3-4 decode-before-gate, P3-5 WER
alignment doc. Accepted deviation: P3-2 clamp_budget threading (hayro's
budget clamp is a safety bound; 200 dpi OCR never approaches it).
Post-fix verification: check/build green, hayro non-blank ok, 90/0 +
98/0 matrices, harness selfcheck 0.
