# Verdict — TASK 4 (human): bake-off verdict + flip decision

ok:true

Decision: **STAY** — pdfium remains the default OCR renderer.
Rule check: WER(hayro) 0.6964 > WER(pdfium) 0.6683 + 0.01 at 150 dpi
(+2.81pt, rule NOT satisfied); at 300 dpi +0.90pt (satisfied, margin
0.10pt). The human declined the flip (2026-11-09, option 1).
Evidence: .workflows/spikes/render-bakeoff/report-150dpi-30p.md and
report-300dpi-30p.md (30 pages, 55,621 ground-truth words, arms verified
distinct, committed 9182321). TASK 5 executes the stay branch.
