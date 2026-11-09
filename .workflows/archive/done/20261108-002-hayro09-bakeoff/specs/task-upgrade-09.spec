spec: task
name: "upgrade-09"
tags: [render, hayro, vendor, upgrade]
verify: "cargo build --features render-hayro && cargo test --features pdf,ocr-indra-native,render-hayro && ! test -d vendor/hayro-jbig2 && ! test -d vendor/vello_cpu && grep -q 'hayro-font' Cargo.toml"
max-rounds: 3
---

## Intent

Move the vendored hayro stack from the PdfCraft 0.7.x patches to upstream 0.9.0, porting the 10
still-needed patch families onto the restructured tree, dropping the 4 unwired vendor dirs, and
rewriting the render glue. Human decision 1(b): upgrade the wired crates only + adopt hayro-font.
Human decision 3(a): transcript re-baselines are allowed, reviewer-gated, never silent.

## Decisions

- vendor/ after this task: hayro 0.9.0, hayro-syntax 0.9.0, hayro-interpret 0.9.0,
  hayro-cmap 0.1.1+, hayro-font 0.5.0 (NEW) — copied from upstream (laurenzv/hayro @
  360f4a2 or newer 0.9.x tag), MIT OR Apache-2.0 licenses kept. DELETED: vendor/hayro-jbig2,
  vendor/hayro-postscript, vendor/vello_common, vendor/vello_cpu (registry serves them;
  jbig2 0.4.0's memory budget supersedes our caps).
- The 10 still-needed families are re-applied per the TASK 1 anchor map
  (`.workflows/scout/20261108-002.md`), preserving the `PdfCraft patch:` mark convention:
  content byte budgets, renderer allocation caps, lenient cmap (skip malformed, not strict),
  Prev-as-indirect-ref, sync poison recovery, run-length decode limit, knockout/isolated +
  CMYK group passes, annotation flags, OCG /VE, offscreen-path skip.
- Glue rewrite (src/render.rs + the hayro call path in pdf.rs): `render_size()` →
  `render()`/Pixmap + PixmapSettings; `render_into(&mut [u8])` → RenderContext + Affine;
  `logging`/`simd` features dropped from feature wiring. One call path — no second renderer
  abstraction.
- `[patch.crates-io]` + workspace excludes updated for the new crate set (hayro-font added;
  jbig2/postscript/vello entries removed); hayro-font stays VENDORED (same doctrine as the
  rest — no crates.io pivot for patched crates).
- Transcript re-baseline (decision 3a): if 0.9.0 renders shift OCR transcripts on fixture
  tests, update the expected values with a PER-FIXTURE diff record in the task verdict
  (old text → new text, why the render change caused it). Silent expectation edits are a
  defect. tests/pdf_ocr.rs count (6) may grow only if a fixture needs a new assertion.

## Boundaries

### Allowed Changes
- `vendor/hayro*`, `vendor/vello_*` (replace/delete per Decisions), `vendor/hayro-font/` (new)
- `Cargo.toml`, `Cargo.lock` is gitignored — prove dependency changes with `cargo tree`
- `src/render.rs`, `src/formats/pdf.rs` (hayro glue only)
- fixture expected-value updates EXCLUSIVELY for transcript re-baseline (per Decision 3a,
  each recorded in the verdict)

### Forbidden
- Any change to Indra models/transcripts logic beyond what renders cause
- `src/quality.rs`, `src/table_structure module`, `src/formats/indra_native.rs`
- Re-adding vendored jbig2/postscript/vello copies
- New dependencies beyond the hayro workspace itself

## Acceptance Criteria

- Scenario: vendor layout upgraded
  Test: verify command (vendor dirs exact set; `grep -q hayro-font Cargo.toml`; versions 0.9.x in vendor Cargo.tomls)
- Scenario: the 10 families present with marks
  Test: `grep -rn "PdfCraft patch:" vendor/hayro* | wc -l` ≥ 10 distinct families (spot-check
  anchors: content budget in the interpret loop, cmap lenient path, Prev indirect read,
  knockout group pass in the render device)
- Scenario: builds and renders non-blank through 0.9.0
  Test: `test_hayro_renders_fixtures_non_blank` (all 5 fixtures, hayro path, non-white pixel ratio > 0)
- Scenario: glue migrated
  Test: `! grep -rn "render_size\|render_into(&mut \[" src/`
- Scenario: transcripts re-baselined with evidence, matrix green
  Test: full `pdf,ocr-indra-native` matrix green; every changed expectation listed in the verdict
  with old→new + cause
EOF
