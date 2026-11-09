# Plan: Pure-Rust rendering by default — hayro 0.9.0 upgrade + WER bake-off

> Plan ID: 20261108-002
> Created: 2026-11-09
> Status: APPROVED (human 2026-11-09 — scope (b), flip rule, re-baseline (a), corpus (b))

## Lifecycle
- Plan ID assigned per archive count: 1 done entry + 0 superseded + 1 = **002**. Never reused.
- SHIP (`/review`, all tasks ✅) → Status: DONE, bundle to `.workflows/archive/done/20261108-002-hayro09-bakeoff/`.
- `/abort` → SUPERSEDED to `.workflows/archive/superseded/`.

## Context
- Goal: make **hayro (pure Rust) the default OCR renderer**, pdfium the fallback, on the strength of measured evidence — and upgrade the vendored hayro stack from the PdfCraft 0.7.x patches to upstream 0.9.0 while porting the 10 still-needed patch families.
- Evidence in hand:
  - Dense 10-page A/B (2026-11-09, `--ocr-force`): hayro **+1.14% chars** vs pdfium, ~5% faster (6.99s vs 7.32s); light fixtures dead heat (`.workflows/research/hayro-0.9.0-upstream-report.md`).
  - Scout report (572 PdfCraft patch marks mapped to 0.9.0): **10 patch families still needed upstream** (content byte budgets, renderer allocation caps, lenient cmap, Prev-as-indirect-ref, sync poison recovery, run-length limit, knockout/CMYK group passes, annotation flags, OCG /VE, offscreen-path skip); API port MEDIUM — glue is one ~35-line fn (`src/formats/pdf.rs:2714`), `render_size()` removed, `PixmapSettings` + `render()`/Pixmap replace `render_into(&mut [u8])`; renderer.rs split into 8 modules, fonts moved to `hayro-font`.
  - **Wiring discrepancy (AGENTS.md report-only, now resolved by this plan)**: `[patch.crates-io]` wires only hayro/syntax/interpret/cmap — vendored jbig2/postscript/vello dirs are dead weight (registry versions build; jbig2 0.4.0 upstream carries the memory budget that supersedes ours).
- Key decisions (human, 2026-11-09):
  1. **Upgrade scope (b)**: upgrade the 4 wired crates (hayro, hayro-syntax, hayro-interpret, hayro-cmap) + adopt `hayro-font`; **drop the 4 unwired vendor dirs** (registry jbig2/postscript/vello serve).
  2. **Default-flip rule**: flip default to hayro iff WER(hayro) ≤ WER(pdfium) + 1pt on the bake-off corpus AND all re-baselined tests green — the flip itself is a **human-gated** task.
  3. **Transcript re-baseline (a)**: 0.9.0 renders shift OCR transcripts; deliberate re-baseline allowed, every per-fixture diff reviewer-gated in the verdict, never silent.
  4. **Bake-off corpus (b)**: synthetic seeded generator (exact WER ground truth) **plus a `real-scans/` drop-in slot** — human provides real scans, scored with char/word-delta metrics (no WER without ground truth).
- Constraints: `CARGO_BUILD_JOBS=4` (rten+ort+hayro OOM at higher parallelism); MSRV 1.92 (unchanged); no new runtime dependencies (hayro-font is part of the hayro workspace we vendor); fixture locks (`cli/tests/cli.rs`, `tests/docx_marks.rs`) untouched; AGENTS.md doctrue: vendors stay vendored (no crates.io pivot for the patched crates).

## Refactor assessment (mandatory)
Verdict: **extend/upgrade** — the render seam (`src/render.rs` + the one glue fn in pdf.rs) is extended in place; no second rendering path is created (hayro is already selectable; the plan changes WHICH engine is default and ON WHAT VERSION). The dead vendor dirs are SUBTRACTED.

## Docs impact (mandatory — read by the SHIP gate)
- `README.md` (renderer paragraph: default + bake-off numbers)
- `docs/formats.md` (renderer notes if present)
- `CHANGELOG.md` (Plan ID 20261108-002 entry)

## Exploration
See `.workflows/research/hayro-0.9.0-upstream-report.md` (committed 2b90b8a): version matrix, 20-family patch-by-patch table (10 FIXED/SUPERSEDED, 10 STILL NEEDED), API-port inventory, wiring discrepancy evidence (`[patch.crates-io]` covers 4 of 8 vendored crates), upstream-only wins (crypt filters, hinting, glyph caching, AA control, warning_sink).

## Spec Map

| Task | Agent | Goal | Bottleneck | Testing Strategy | Wave | Depends On |
|------|-------|------|------------|------------------|------|------------|
| TASK_1 | scout | Map the 10 patch families onto the 0.9.0 tree (file/function anchors per family) + glue-rewrite API surface + registry jbig2/postscript API check | ⚪ STANDARD | — | 0 | — |
| TASK_2 | worker | Upgrade vendored stack to 0.9.0: copy wired crates + hayro-font, re-apply the 10 families, drop unwired vendor dirs, rewrite glue, re-baseline fixture transcripts (verdict-gated) | 🔴 BLOCKING | example-based | 1 | TASK_1 |
| TASK_3 | worker | Bake-off harness: seeded dense-scan generator + WER scorer + pdfium/hayro A/B runner + `real-scans/` drop-in slot | 🟠 VERIFICATION_HEAVY | example-based | 1 | TASK_1 |
| TASK_4 | human | Read the bake-off report; decide the default flip per rule 2 | 🔴 BLOCKING | — | 2 | TASK_2, TASK_3 |
| TASK_5 | worker | Execute the T4 decision: flip default (hayro first, pdfium fallback) or document staying | 🟡 RISKY | example-based | 3 | TASK_4 |
| TASK_6 | worker | Docs sweep: README renderer paragraph, formats.md notes, CHANGELOG entry | ⚪ STANDARD | — | 4 | TASK_5 |

Parallel notes: TASK_2 (vendor/ + src/render.rs + Cargo.toml) and TASK_3 (`.workflows/spikes/render-bakeoff/` only) have disjoint Allowed Changes — same wave is safe.

## Tasks

### TASK 1 (⚪): Scout — 0.9.0 tree map for the 10 families
- **Agent**: scout
- **Depends on**: none
- **Wave**: 0
- **What**: For each of the 10 still-needed families (list in the research report), pin the exact 0.9.0 file/function anchors in /tmp/hayro-upstream (or a fresh clone) where the port lands — renderer.rs split into 8 modules, fonts in hayro-font. Also: the exact new render API surface for the glue rewrite (PixmapSettings, render(), RenderCache changes), and the registry hayro-jbig2 0.4.0 / hayro-postscript 0.2.0 API check (do they satisfy what our unwired vendored copies provided?). Findings to `.workflows/scout/20261108-002.md`.
- **Verify**: `.workflows/scout/20261108-002.md` exists; every family has ≥1 file:line anchor; glue section lists the new signatures.
- **context**: no changes
- **Status**: ✅ DONE (gate persisted .workflows/scout/20261108-002.md; anchors for all 10 families + glue spec + registry check incl. offline gap)

### TASK 2 (🔴): Upgrade vendored stack to 0.9.0
- **Agent**: worker
- **Depends on**: TASK 1
- **Wave**: 1
- **Contract**: `.workflows/specs/task-upgrade-09.spec`
- **Bottleneck**: 🔴 BLOCKING
- **Testing strategy**: example-based
- **What**: vendor/ = 0.9.0 copies of hayro, hayro-syntax, hayro-interpret, hayro-cmap + NEW hayro-font; re-apply the 10 families per the TASK 1 anchors (patch marks preserved, `PdfCraft patch:` convention); DELETE vendor dirs hayro-jbig2, hayro-postscript, vello_common, vello_cpu (registry serves; update `[patch.crates-io]` + workspace excludes); rewrite the glue in pdf.rs against render()/PixmapSettings; **re-baseline** any fixture-locked transcript diffs with per-fixture verdict records (decision 3).
- **Verify**: `cargo build --features render-hayro` green; all 5 fixtures render non-blank rasters through hayro; parser matrix green post re-baseline; `ls vendor/` shows exactly hayro, hayro-syntax, hayro-interpret, hayro-cmap, hayro-font (+hayro-jbig2 etc. REMOVED).
- **Status**: ⬜ PENDING

### TASK 3 (🟠): Bake-off harness + real-scans slot
- **Agent**: worker
- **Depends on**: TASK 1
- **Wave**: 1
- **Contract**: `.workflows/specs/task-bakeoff.spec`
- **Bottleneck**: 🟠 VERIFICATION_HEAVY
- **Testing strategy**: example-based
- **What**: `.workflows/spikes/render-bakeoff/`: (1) seeded dense-scan generator (formalize the reportlab→pdftoppm→drawImage flow, N pages configurable, exact ground-truth words.json emitted); (2) WER scorer (generated ground truth vs OCR transcript); (3) A/B runner: pdfium vs hayro over the synthetic corpus + the 5 fixtures + `.workflows/spikes/render-bakeoff/real-scans/` drop-in (any PDFs placed there are scored with char/word-delta, no WER claim); (4) markdown report table. CARGO_BUILD_JOBS=4 respected; measurement-only — no src/ changes, no CI wiring.
- **Verify**: harness runs end-to-end (both renderers × corpus), WER + delta table produced, real-scans slot demonstrably scored (≥1 scan present — the human drops them in).
- **Status**: ✅ DONE (gate: harness applied from the store event, WER selfcheck 9/9, runner --selfcheck exit 0 after two gate fixes — reportlab /ID+XMP byte noise normalized to rasterized-content hashes, path-hash bug; report.md generated at the T2 gate for TASK 4)
- **context**: no changes

### TASK 4 (🔴): HUMAN — bake-off verdict + flip decision
- **Agent**: human
- **Depends on**: TASK 2, TASK 3
- **Wave**: 2
- **What**: Read the bake-off report. Apply rule 2: flip default to hayro iff WER(hayro) ≤ WER(pdfium) + 1pt AND re-baselined tests green. Decision recorded in the plan (STATUS line under this task).
- **Status**: ⬜ PENDING

### TASK 5 (🟡): Execute the renderer-default decision
- **Agent**: worker
- **Depends on**: TASK 4
- **Wave**: 3
- **Contract**: `.workflows/specs/task-default-flip.spec`
- **Bottleneck**: 🟡 RISKY
- **Testing strategy**: example-based
- **What**: If flip: `render-pages` default order becomes hayro → pdfium (fallback); OFFICE_PARSER_RENDERER still overrides both ways; docs updated in the same change. If stay: the report + decision documented in CONTEXT.md §Rulings. Conditional contract — finalized after TASK 4.
- **Verify**: `cargo test --features pdf,ocr-indra-native,render-hayro,render-pdfium` green; the rendered-output A/B on fixtures matches the decision.
- **Status**: ⬜ PENDING

### TASK 6 (⚪): Docs sweep
- **Agent**: worker
- **Depends on**: TASK 5
- **Wave**: 4
- **What**: README renderer paragraph (default, bake-off numbers, hayro-font note), docs/formats.md renderer notes, CHANGELOG entry with Plan ID 20261108-002.
- **Status**: ⬜ PENDING

(No verify/quality tasks: `/next` verdict-gates every task; `/review` is the final gate at SHIP.)

## Contracts (generated after approval)
| Task | Contract File | Scenarios |
|------|--------------|-----------|
| TASK_2 | .workflows/specs/task-upgrade-09.spec | ~5 |
| TASK_3 | .workflows/specs/task-bakeoff.spec | ~4 |
| TASK_5 | .workflows/specs/task-default-flip.spec | ~3 (conditional, finalized post-T4) |

## Execution Notes
<per-task cost/duration/learnings appended during execution>
