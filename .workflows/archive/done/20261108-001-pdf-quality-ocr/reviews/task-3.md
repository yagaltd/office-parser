# Verdict — TASK 3 (prototype): complexity score calibration

ok:true

Reviewer: orchestrator (gate), 2026-11-08.

Evidence:
- Sources materialized verbatim from the prototype author's report; 4 compile fixes inside the
  spike dir only. `src/` untouched.
- Measured table (data/scores.json + score-table.md + README §Score table):
  gutters = 0 on EVERY text-layer page; pdf_mixed p1 (digital) score 0.039, columns 1;
  pdf_cjk_scan p1 score 0.095; all four image-only pages correctly Escalate via ImageOnly
  (score undefined — no fake 0s).
- Threshold 0.30 confirmed: no false escalations, no false keeps.
- Deletion test: **branch 1 wins** — score ≡ columns on this corpus → TASK 7 scope shrinks to
  `columns` + `pdf_text_quality` routing + optional standalone rotation flag. task-7.spec must
  be re-scoped by the orchestrator before TASK 7 dispatches.
