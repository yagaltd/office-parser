# Verdict — TASK 3 (worker): bake-off harness

ok:true

Reviewer: orchestrator (gate), 2026-11-09.

Evidence: harness applied from the worker's store event (4 files, 525 lines:
generate_corpus.py 252, wer.py 137, run_bakeoff.py 755, README.md 122;
real-scans/ slot). wer.py --selfcheck 9/9 hand-computed cases PASS
(headline 1/3 substitution case). run_bakeoff.py --selfcheck exit 0.

Gate fixes recorded (both gate-authored, worker-invisible environmental):
1. Generator byte-determinism was over-asserted: reportlab invariant=1 pins
   the timestamp but not the /ID (os.urandom) nor FormXob.<hex32> resource
   names nor the compressed XMP block — 776 bytes of metadata noise, zero
   content effect. Normalized the /ID + resource names in the generator AND
   re-pointed the selfcheck at CONTENT determinism (pdftoppm-re-rendered
   page hashes + words.json) — the criterion the measurement actually
   consumes.
2. raster_pages_sha hashed temp-file PATHS (differ per run); now hashes
   basename + bytes only.

Isolation honored: only .workflows/spikes/render-bakeoff/** touched.
Contract criteria: deterministic generation PASS, WER sanity PASS,
real-scans slot present (empty-tolerant per design), A/B end-to-end run
pending T2's 0.9.0 binary (the report.md gate is T2's, feeding TASK 4).
