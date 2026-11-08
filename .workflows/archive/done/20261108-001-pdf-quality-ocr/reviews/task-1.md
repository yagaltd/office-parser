# Verdict — TASK 1 (scout): interpreter seams + fixture baseline

ok:true

Reviewer: orchestrator (gate), 2026-11-08.

Evidence:
- Baseline artifacts: `.workflows/spikes/scout/baseline/` (counts.txt, body-counts.txt,
  test-results.txt, per-fixture split-page md, mixed-nopdfium/ + mixed-pdfium/ OCR JSON).
- Test matrix green: 44 passed / 0 failed across targets (the "0 passed" line was the doctest tail).
- Findings persisted to `.workflows/CONTEXT.md` (committed in this wave's commit).

Key findings carried forward:
- Trust checks silently degrade without render-pdfium (pdf.rs:2573-2586) — pdfium-free no-ink
  check (TASK 9) is the fix.
- pdf_mixed digital page reads single-column in the text layer (no column interleaving) —
  matches TASK 3's measured gutters=0.
- pdf_cjk_scan p1 has a 16-char latin-1 mojibake text layer that defeats both the char
  threshold and trust check #2 (needs UTF-16/PUE detection, not valid-Latin-1).
- page_block_ranges stale after OCR merge in the COMBINED doc (split-pages unaffected) —
  defect-class, folded into TASK 8 scope.
- render.rs JSON + CLI frontmatter omit pdf_text_quality — TASK 5 must add per-page quality
  to the JSON shape.
- Registry deviation: deepseek/deepseek-flash not spawnable this session; runtime assigned
  zai/glm-5.3-flash (thinking high). Recorded; re-probe on deepseek before relying on its
  max-thinking behavior.
