# Verdict — TASK 2 (prototype): per-span boxes + ink coverage without pdfium

ok:true

Reviewer: orchestrator (gate), 2026-11-08.

Evidence:
- Spike materialized verbatim from the prototype author's final report (session extraction);
  1 compile fix inside the spike dir (`resolve_stream` lifetime). `src/` untouched:
  `git diff --name-only src/` empty.
- Executed once (exit 0): scanned coverage exactly {0.000000, 0.000000} (structural: no /Font,
  no BT/Tj bytes); pdf_mixed p1 = 0.052249 (8 spans, 25,325 pt² of 484,704 pt²); p2 = 0.
- Pre-stated decision rule: GO iff coverage(mixed p1) > 0.005 while scanned == 0 → **GO**.
- Seam verdict: **GO — Track 1 is `extend`.** Box dump carries all 7 contract fields per span
  from state already held at emit points (Tj :1007, TJ :1034, ' :1059, " :1086, BDC :1117).
- Caveats accepted as stated: Tc/Tw/Tz untracked (±15% advance error, irrelevant to
  0-vs-threshold), typographic-box coverage ≠ raster ink, no tm advance after show ops.

TASK 4 proceeds as `extend` (per plan, not refactor-first).
