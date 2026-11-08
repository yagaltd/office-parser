# Verdict — TASK 12 (worker): docs + CHANGELOG sweep with the Plan ID

ok:true

Reviewer: orchestrator (gate), 2026-11-08.

Evidence:
- Spec verify: `grep -q '20261108-001' CHANGELOG.md` PASS (exactly 1
  occurrence, the Plan-ID subsection) + `test -f docs/quality.md` PASS
  (new 150-line doc: document-level summary, per-page
  `page_quality`/`extra["pdf_quality"]` JSON shapes, all verdict kinds
  with their reason strings (incl. "no-ink" + pdfium reasons),
  decision order, heuristic-status disclosure, routing/rollup/force,
  the full `extra["ocr"]` payload (absent-not-null rich keys), and the
  ocr-rusto-core opt-in semantics incl. model resolution + no-op).
- README: docs-list link added; Limitations PDF bullet now states
  per-page quality + verdicts + XY-Cut reading order (OCR paragraph
  already current from the task-11 gate).
- docs/formats.md: new "PDF quality and OCR" section (the four shipped
  surfaces were previously one line of prose).
- docs/library-usage.md: `PageOcrEngine` custom-engine contract (defaulted
  `ocr_page`) + a `page_quality` usage example with the non-PDF `None`
  discipline stated.
- CHANGELOG duplicate `## Unreleased` (stranded mind-map section between
  0.2.0 and 0.1.0) merged under the single top Unreleased as a `###`
  subsection — one heading remains (grep count 1), zero content lost;
  no release dates invented. One stray worker prose line was removed at
  the gate (delivery artifact).
- `scripts/check-archive.sh` does not exist in the repo — the spec's SHIP
  dry run has nothing to run; recorded here (create it at SHIP time or
  drop the criterion at review).
- Default build still green after the sweep (docs-only change).
