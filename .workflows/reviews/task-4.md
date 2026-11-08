# Verdict — TASK 4 (worker): PageGeometry emission

ok: true

Reviewer: orchestrator (gate), 2026-11-08.

Evidence:
- Anchor edits A–M applied to src/formats/pdf.rs (+10.3 KB: CoordinateSystem, SpanBox,
  PageGeometry types; push_text_atom consolidation collapsing the 5 duplicated atom-push
  blocks; page_inherited_boxes + build_page_geometry; parse_pdf_full wiring).
- tests/pdf_geometry.rs: **5/5 pass** — spans carry bbox/font/size; crop-origin subtraction
  exact (x0=50 = pen 150 − crop 100; reversible by adding the recorded origin);
  /Rotate 90/270 preserved on synthesized pages (the rot90 FIXTURE carries no /Rotate —
  spec premise corrected, fixture's true state pinned: landscape MediaBox, rotation 0);
  block-output pins (counts/texts/levels/image-id prefixes) all hold; empty pages yield
  empty geometry with boxes resolved.
- **Byte-identical proof: CLI output on pdf_scanned.pdf before vs after = 760 bytes, diff
  empty.** Full matrix green (53 passed/0 failed with pdf feature).

Accepted deviations (escalate-worthy findings, resolved):
1. Spec premise: pdf_rot90_scan.pdf has NO /Rotate — resolved via synthesized /Rotate pages.
2. No pre-change render goldens — replaced by full block-output pins + gate CLI diff.
3. pages_geometry always populated (downstream TASK 6/8/9 consume via parse); crop_box
   stores the RAW declared box (subtraction reversible).
- TASK 5 collision watch: the block pins are TASK 4's own (pdf.rs untouched by TASK 5).

TASK 6 (XY-Cut) may proceed on the new geometry.
