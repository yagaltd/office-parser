# Verdict — TASK 8 (worker): per-page verdicts + document rollup

ok:true

Reviewer: orchestrator (gate), 2026-11-08.

Evidence:
- Applied: routing_rollup (pub) + pdf_quality_json (private) in pdf.rs;
  router refactored to verdict consumption (Escalate NEVER routes;
  pdfium trust reasons fold INTO the verdicts so page_quality carries the
  reason, preserving old routing precedence); worst-page-wins rollup;
  extra["pdf_quality"] mirror on plain parse + routed parse + per-page
  split docs; three eprintln sites updated to the new vec.
- tests/pdf_quality_pages.rs: **6/6 pass** — Escalate-never-OCRs negative
  (rich rotated synth page: verdict Escalate, engine.calls == 0, text layer
  intact, pages_ocr == 0), Trust-page negative, split/opts parity (always-on
  synth parity + renderer-gated pdf_mixed three-way incl. text-layer truth
  comparison), extra mirror (combined + per-page), worst-page-wins rollup
  (routing_rollup unit + end-to-end: only the worst page routed, doc-level
  enum unchanged).
- Full matrix green: 75 passed / 0 failed with pdf feature. Clippy clean.
- Routing equivalence argued + pinned by tests: Ocr verdict ⇔ (count==0 |
  count<min_chars); trust reason ⇒ Ocr; Escalate never routed; force
  unchanged.

Design decisions recorded: full verdict consumption (no escape hatch);
trust-reason fold preserves precedence; parity locked by regression tests.
