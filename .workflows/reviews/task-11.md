# Verdict — TASK 11 (worker): ocr-rusto-core table-region enrichment

ok: true

Reviewer: orchestrator (gate), 2026-11-08.

History: this task was first ESCALATED (rusto-core "missing" from crates.io);
the human provided the local checkout; the dependency then went local-path →
git+rev (steer) → **vendored**. Amendments to the spec, all human-approved:
1. rusto-nn (pure-Rust rten 0.27 runtime) joins rusto-core — the spec's
   "no rusto-nn" line predates the discovery that rusto-core's detector needs
   it; no facade, no ffi, no ort.
2. Dependency = vendored crates (vendor/rusto-core, vendor/rusto-nn, MIT /
   MIT OR Apache-2.0, from byrizki/rusto-rs @ 991fedf3). The git+rev dep was
   measured infeasible: cargo clones that repo recursively INCLUDING its
   third-party submodules (RapidDoc, RapidUnDistort, alibaba/MNN,
   PaddlePaddle/PaddleX) — >1.2 GB per consumer/CI run, 52 min wall time.
   Vendoring two crates (452K + 1.4M) follows the hayro precedent.

Evidence:
- Feature `ocr-rusto-core` off by default: default `cargo build` 5.1 s;
  `cargo tree --no-default-features --features pdf -e normal` contains
  **0** rusto-core occurrences (the spec's proof line).
- Unit tests (lib): **5/5** — no-model no-op (payload = feature-off shape),
  bogus-model unchanged, Table-only region filter, html gating
  (region AND block), span/escape rendering.
- Integration (tests/ocr_rusto_tables.rs): **4/4** — transcript identity
  on/off (skip-gated on the absent DocOCR layout model, captured),
  table_html positive arm (same skip gate), feature-off payload shape on
  real Indra + fixture (pages_ocr == 2, rich keys absent), and
  indra_native.rs diff-empty lock.
- Transcript safety: enrichment runs AFTER engine.ocr_page(), only fills
  still-None rich fields; blocks/confidence pass through (struct update).
  Table cell content stays Indra's (rusto SLANet postprocess is a stub —
  verified by the worker; no re-recognition anywhere).
- Matrix: `pdf,ocr-indra-native` 98 passed / 0 failed; clippy 0 errors
  (1 pre-existing `pref` warning in pdf.rs renderer-env code, outside this
  task's files; vendored crates are not clippy targets).
- pdf_ocr.rs diff empty; word_boxes deliberately never populated
  (Indra seam exposes no per-word data) — asserted.
