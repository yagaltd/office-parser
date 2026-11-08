# Plan: Geometry-aware per-page PDF quality + optional richer OCR page contract

> Plan ID: 20261108-001
> Created: 2026-11-08
> Status: APPROVED (human approval 2026-11-08 — session A: build authorized)
> Proposal: `.workflows/proposals/2026-11-08-pdf-path-upgrade.md`
> Domain memory: `.workflows/CONTEXT.md` (§Rulings carries both locked decisions)

## Lifecycle

- Plan ID assigned once at bootstrap: `20261108-001` (archives empty → NNN = 1). Never reused.
- `/idea` and `/plan` refuse to overwrite a live plan (DRAFT/APPROVED/IN PROGRESS) — route it to
  `/review` (ship → `archive/done/`) or `/abort` (abandon → `archive/superseded/`) first.
- SHIP (`/review`, all tasks ✅) → Status: DONE and move the bundle (plan.md + specs/ + reviews/ +
  CONTEXT.snapshot.md) to `.workflows/archive/done/20261108-001-pdf-quality/`.
- SHIP gate: `bash scripts/check-archive.sh <archive-target-dir>` — a missing declared Docs impact
  path or a CHANGELOG without the Plan ID refuses the archive.
- `.workflows/CONTEXT.md`, `docs/adr/`, `knowledge/`, `LOG.md` are never archived.

## Context

`pdf_text_quality` is one enum per document, computed by counting alphanumeric characters across all
pages (`src/formats/pdf.rs:2099-2112`). It cannot see arrangement, so a three-column paper read as a
linear stream reports `TextRich` while being useless for retrieval; it cannot locate a scanned insert
inside a 200-page digital compilation; and it cannot answer the scan-vs-white-page question our own
source marks as pending:

> `pdf.rs:2571` — `(no-ink check pending: needs char-box geometry.)`

Meanwhile per-page routing *already* works (`page_char_counts`, `pdf.rs:2520`), so the reported
quality and the actual routing already disagree. This plan makes the reported quality per-page and
geometry-aware, and adds an `Escalate` verdict for pages that parse cleanly but structure badly —
the case where a vision model pays for itself.

Constraints (from `.workflows/CONTEXT.md` §Rulings, decided 2026-11-08):

- **No AGPL dependency.** `rasgapdf`'s design is borrowed; its code is not.
- **`rusto-core` only, off by default, tables only** (TASK_11). Never the `rusto-rs` facade crate.
- **No pdfium dependency added for quality.** Geometry comes from the text layer.
- `PdfTextQuality` stays source-compatible; new surface is additive.
- `--ocr-min-chars 50` default unchanged.

## Exploration

See `.workflows/proposals/2026-11-08-pdf-path-upgrade.md` §Evidence for the full anchor list. The
load-bearing findings:

- The content interpreter already tracks everything per-span boxes need: `state.tm`,
  `state.font_size` (`pdf.rs:925`, `:1005`, `:1024`, `:1059`), `effective_font_size()` (`:291`),
  `estimate_text_advance()` (`:471`). Geometry is a wiring job, not a rewrite — **to be proven by
  TASK_1p**, and if it fails this plan escalates as `refactor-first`.
- Trust checks silently degrade to `Vec::new()` without `render-pdfium` (`pdf.rs:2572-2585`).
- `--split-pages` has its own per-page quality path (`pdf.rs:2703`) that must stay consistent.
- Five PDF fixtures already exist and cover the interesting cases: `pdf_scanned.pdf`,
  `pdf_table_scan.pdf`, `pdf_rot90_scan.pdf`, `pdf_cjk_scan.pdf`, `pdf_mixed.pdf`.
- No dev-dependencies exist in this crate. All testing strategy below is **example-based on the
  existing fixtures** (house style); no test framework is introduced by this plan.

## Spec Map

| Task | Agent | Goal | Bottleneck | Testing Strategy | Wave | Depends On |
|------|-------|------|------------|------------------|------|------------|
| TASK_1 | scout | Inventory interpreter seams + fixture baseline | ⚪ STANDARD | — | 0 | — |
| TASK_2 | worker (prototype) | Can we emit per-span boxes + ink coverage without pdfium? | 🟡 RISKY | evidence artifact | 1 | TASK_1 |
| TASK_3 | worker (prototype) | Does deterministic complexity separate the 5 fixtures? | 🟡 RISKY | evidence artifact | 1 | TASK_2 |
| TASK_4 | worker | `PageGeometry` emission (`pdf.rs`) | 🔴 BLOCKING | example-based | 2 | TASK_2 |
| TASK_5 | worker | `PageTextQuality` types (`src/quality.rs`, `document.rs`) | 🟡 RISKY | example-based | 2 | TASK_1 |
| TASK_6 | worker | XY-Cut reading order (`pdf.rs`) | 🔴 BLOCKING | example-based | 3 | TASK_4 |
| TASK_7 | worker | Complexity scoring (`src/quality.rs`) | 🟡 RISKY | example-based | 3 | TASK_3, TASK_5 |
| TASK_8 | worker | Per-page verdicts + rollup (`pdf.rs`) | 🔴 BLOCKING | example-based | 4 | TASK_5, 6, 7 |
| TASK_9 | worker | pdfium-free no-ink trust check (`pdf.rs`) | 🟡 RISKY | example-based (both cfgs) | 5 | TASK_8 |
| TASK_10 | worker | `OcrPage` + defaulted `ocr_page()` (`pdf.rs`) | 🟡 RISKY | example-based | 6 | TASK_8 |
| TASK_11 | worker | `ocr-rusto-core` feature, tables only | 🟡 RISKY | example-based | 7 | TASK_10 |
| TASK_12 | worker (docs) | Docs + CHANGELOG with Plan ID | ⚪ STANDARD | — | 8 | TASK_4-11 |

Waves: **A** = TASK_4 ‖ TASK_5 (disjoint: `src/formats/pdf.rs` vs `src/quality.rs` + `src/document.rs`).
**B** = TASK_6 ‖ TASK_7 (disjoint: `src/formats/pdf.rs` vs `src/quality.rs`).
Every other task touches `src/formats/pdf.rs` and therefore runs serially — this plan is deliberately
sequential where the file is shared rather than pretending to parallelize.

Prototypes (TASK_2, 3p) are serial, Tier C routing, never in a parallel group.

## Tasks

### TASK 1 (⚪): Scout interpreter seams and fixture baseline
- **Agent**: scout
- **Depends on**: none
- **Wave**: 0
- **Bottleneck**: ⚪ STANDARD
- **What**: Confirm, by reading `src/formats/pdf.rs`, exactly which interpreter state is available at
  the point each text block is emitted, and produce a per-fixture baseline table: page count,
  `pdf_text_quality`, per-page `page_char_counts`, whether `render-pdfium` changes the trust verdict,
  and current reading-order output for `pdf_mixed.pdf` (the multi-column case). Append findings to
  `.workflows/CONTEXT.md` §Open Questions where they resolve or sharpen a question.
- **Verify**: findings section present in `.workflows/CONTEXT.md`; baseline table lists all five
  fixtures with non-empty rows; `cargo test --features pdf` still green (baseline is green).
- **Status**: ✅ DONE (verdict + baseline persisted to CONTEXT.md)

### TASK 2 (🟡): per-span boxes + ink coverage without pdfium
- **Agent**: worker (prototype)
- **Depends on**: TASK_1
- **Wave**: 1
- **Bottleneck**: 🟡 RISKY
- **Contract**: `.workflows/specs/task-2.spec`
- **Testing strategy**: evidence criteria (artifact on disk, measurements, not feature behavior)
- **What**: One bounded question — *does the existing content interpreter already hold enough state to
  emit per-span `[x0,y0,x1,y1]` boxes with `font_name`/`font_size`, and does glyph-area ÷ page-area
  separate `pdf_scanned.pdf` from `pdf_mixed.pdf`?* Spike code lives in
  `.workflows/spikes/pdf-geometry/**` only. Record: box dump for both fixtures, measured ink coverage
  for every page, and a GO / NO-GO on "Track 1 is `extend`". A NO-GO is a valid, useful answer — it
  re-tiers TASK_4 as `refactor-first` and stops the plan for human review.
- **Verify**: `.workflows/spikes/pdf-geometry/README.md` states GO/NO-GO with the measured coverage
  table; production paths (`src/**`) untouched (`git diff --stat src/` empty).
- **Status**: ✅ DONE (GO measured (coverage 0.0522 > 0.005 rule))

### TASK 3 (🟡): complexity score calibration on the five fixtures
- **Agent**: worker (prototype)
- **Depends on**: TASK_2
- **Wave**: 1
- **Bottleneck**: 🟡 RISKY
- **Contract**: `.workflows/specs/task-3.spec`
- **Testing strategy**: evidence criteria
- **What**: One bounded question — *does a deterministic `layout_complexity` built from span boxes
  (horizontal gutter count, table gutter density, span-overlap ratio, rotation ≠ 0, char density)
  rank the fixtures the way a human would?* Produce a table of scores for all five fixtures and a
  written justification for whichever threshold is proposed. Deletion test: if the score cannot
  separate `pdf_mixed.pdf` from a single-column text page, say so and propose dropping the score in
  favour of `columns` alone. Spike code in `.workflows/spikes/pdf-complexity/**` only.
- **Verify**: `.workflows/spikes/pdf-complexity/README.md` holds the five-fixture score table plus a
  threshold recommendation with its justification; `git diff --stat src/` empty.
- **Status**: ✅ DONE — measured: gutters=0 on all text-layer pages; score/columns agree
  everywhere → TASK 7 scope shrinks to columns + pdf_text_quality routing + standalone
  rotation flag (deletion test branch 1). See spikes/pdf-complexity/README.md §Measured verdict.

### TASK 4 (🔴): `PageGeometry` emission
- **Agent**: worker
- **Depends on**: TASK_2
- **Wave**: 2
- **Parallel group**: [PARALLEL-GROUP: A] — boundary `src/formats/pdf.rs`, disjoint from TASK_5
- **Contract**: `.workflows/specs/task-4.spec`
- **Bottleneck**: 🔴 BLOCKING
- **Testing strategy**: example-based
- **What**: Emit per-page geometry from the interpreter: one record per span with bbox, `font_name`,
  `font_size`, plus page `media_box`, `crop_box` (origin subtracted), `initial_rotation`, and
  `coordinate_system`. `PageGeometry` and `SpanBox` are defined **here, in `src/formats/pdf.rs`**
  (PDF-scoped), not in `src/document.rs` — that keeps this task disjoint from TASK_5. Strictly
  additive: `Document.blocks` and every existing block field are unchanged. Populate from the
  TASK_2 spike's proven approach, not from its throwaway code.
- **Verify**: `cargo test --features pdf`; new `tests/pdf_geometry.rs` asserts crop-origin subtraction
  on a fixture with a non-zero CropBox and that `git diff` on existing block output is empty.
- **Status**: ✅ DONE (gate: 5/5 geometry tests, byte-identical CLI proof, /Rotate spec premise corrected via synthesized pages)

### TASK 5 (🟡): `PageTextQuality` types
- **Agent**: worker
- **Depends on**: TASK_1
- **Wave**: 2
- **Parallel group**: [PARALLEL-GROUP: A] — boundary `src/quality.rs` + `src/document.rs`, disjoint
  from TASK_4
- **Contract**: `.workflows/specs/task-5.spec`
- **Bottleneck**: 🟡 RISKY
- **Testing strategy**: example-based
- **What**: Define `PageTextQuality { page, text_richness, geometry_confidence, layout_complexity,
  verdict, columns }` and `PageVerdict { Trust { geometry }, Ocr { reason }, Escalate { complexity } }`
  with serde derives and a documented JSON shape. Add `DocumentMetadata.page_quality:
  Option<Vec<PageTextQuality>>` — `None` for every non-PDF format, mirroring `pdf_text_quality`.
  Types only; no wiring into `pdf.rs` in this task. Defines no geometry types — those belong to
  TASK_4 in `src/formats/pdf.rs`.
- **Verify**: `cargo test`; new `tests/quality_types.rs` asserts `page_quality` is `None` for DOCX and
  that JSON round-trips through `render::to_json_value`.
- **Status**: ✅ DONE (gate green: 4 new tests + full matrix; JSON emits pdf_text_quality + page_quality)

### TASK 6 (🔴): XY-Cut reading order
- **Agent**: worker
- **Depends on**: TASK_4
- **Wave**: 3
- **Parallel group**: [PARALLEL-GROUP: B] — boundary `src/formats/pdf.rs`, disjoint from TASK_7
- **Contract**: `.workflows/specs/task-6.spec`
- **Bottleneck**: 🔴 BLOCKING
- **Testing strategy**: example-based
- **What**: Order born-digital page content by recursive XY-Cut over the TASK_4 boxes, replacing the
  linear stream join for multi-column pages. Single-column pages must produce **byte-identical**
  output (regression guard). Preserve the existing heading `level` walk on top of the new order; if
  it breaks, that is the escalation trigger named in TASK_2, not a local workaround.
- **Verify**: `cargo test --features pdf`; `tests/pdf_reading_order.rs` asserts single-column fixture
  output unchanged and that `pdf_mixed.pdf` no longer interleaves columns (asserted on a known
  sentence, not on a snapshot hash).
- **Status**: ✅ DONE (gate: 9/9 reading-order tests incl. synthetic two-column end-to-end; per-region clustering root-cause fix accepted; byte-guard structural)

### TASK 7 (🟡): Complexity scoring
- **Agent**: worker
- **Depends on**: TASK_3, TASK_5
- **Wave**: 3
- **Parallel group**: [PARALLEL-GROUP: B] — boundary `src/quality.rs`, disjoint from TASK_6
- **Contract**: `.workflows/specs/task-7.spec`
- **Bottleneck**: 🟡 RISKY
- **Testing strategy**: example-based (bounds + monotonicity invariants, no new dev-dep)
- **What**: Implement the pure scoring function from the TASK_3 recommendation:
  `fn layout_complexity(spans: &[..], page: &PageGeometry) -> f32` clamped to `0.0..=1.0`. Imports
  `PageGeometry`/`SpanBox` read-only from `src/formats/pdf.rs` (TASK_4) — never modifies that file.
  Document the constant as a heuristic with the calibration note from the spike. If TASK_3
  recommended dropping the score, implement `columns` only and record why in the spec.
- **Verify**: `cargo test`; `tests/quality_complexity.rs` asserts the range `[0,1]` over all five
  fixtures, monotonicity when a column gutter is added, and the threshold decision from TASK_3.
- **Status**: ✅ DONE (gate: 7/7 routing tests; count_columns gate-fixed to span-level union; stale-ranges defect closed)

### TASK 8 (🔴): Per-page verdicts + document rollup
- **Agent**: worker
- **Depends on**: TASK_5, TASK_6, TASK_7
- **Wave**: 4
- **Contract**: `.workflows/specs/task-8.spec`
- **Bottleneck**: 🔴 BLOCKING
- **Testing strategy**: example-based
- **What**: Wire `page_quality` into `parse_pdf_full` / `parse_with_page_ocr_inner`: one
  `PageTextQuality` per page, verdict derived from richness + geometry confidence + complexity.
  **`Escalate` never adds a page to `ocr_pages`.** Document-level rollup: worst page wins for
  routing; the existing `PdfTextQuality` stays as the summary enum. Keep `--split-pages`
  (`pdf.rs:2703`) consistent, and mirror the array into `metadata.extra["pdf_quality"]`.
- **Verify**: `cargo test --features pdf`; `tests/pdf_quality_pages.rs` asserts
  `pdf_mixed.pdf` yields a richer per-page array than the document enum implies, that a page with
  high complexity gets `Escalate` and is **absent** from `pages_ocr`, and that `--split-pages` mode
  reports the same per-page verdicts.
- **Status**: ⬜ PENDING

### TASK 9 (🟡): pdfium-free no-ink trust check
- **Agent**: worker
- **Depends on**: TASK_8
- **Wave**: 5
- **Contract**: `.workflows/specs/task-9.spec`
- **Bottleneck**: 🟡 RISKY
- **Testing strategy**: example-based, run under both cfgs
- **What**: Close `pdf.rs:2571`. Use TASK_4 ink coverage to untrust a text layer whose glyph area is
  implausibly small for its character count, and keep it independent of `render-pdfium`. The pdfium
  path stays as an additional signal where present; where absent, verdicts must now be non-empty
  instead of silently `Vec::new()`.
- **Verify**: `cargo test --features pdf` **and** `cargo test --no-default-features --features
  pdf`; `tests/pdf_no_ink.rs` asserts a scan is untrusted with pdfium disabled.
- **Status**: ⬜ PENDING

### TASK 10 (🟡): `OcrPage` + defaulted `ocr_page()`
- **Agent**: worker
- **Depends on**: TASK_8
- **Wave**: 6
- **Contract**: `.workflows/specs/task-10.spec`
- **Bottleneck**: 🟡 RISKY
- **Testing strategy**: example-based
- **What**: Add `OcrPage { blocks, confidence, layout_regions, word_boxes, table_html }` and a
  **defaulted** `PageOcrEngine::ocr_page()` that wraps `ocr_page_png()` and returns `None` for the new
  optional fields. No existing implementor breaks. `parse_with_page_ocr_inner` switches to the new
  method and threads the new fields into `metadata.extra["ocr"]` where present.
- **Verify**: `cargo test --features pdf,ocr-indra-native`; `tests/pdf_ocr.rs` (6 existing tests)
  unchanged and green; a test asserts a legacy `ocr_page_png`-only implementor still compiles and runs.
- **Status**: ⬜ PENDING

### TASK 11 (🟡): `ocr-rusto-core` feature, tables only
- **Agent**: worker
- **Depends on**: TASK_10
- **Wave**: 7
- **Contract**: `.workflows/specs/task-11.spec`
- **Bottleneck**: 🟡 RISKY
- **Testing strategy**: example-based
- **What**: Add an **off-by-default** `ocr-rusto-core` feature depending on `rusto-core` only — never
  the `rusto-rs` facade crate. New module implements `layout_regions`, `word_boxes`, and
  `table_html` for `LayoutType::Table` regions. Indra keeps the text path: recognition is untouched,
  so transcripts must be **identical** with the feature on and off.
- **Verify**: `cargo build` (feature off — zero new deps in `Cargo.lock`) and `cargo test --features
  pdf,ocr-indra-native,ocr-rusto-core`; a test asserts transcripts are identical with the feature
  on and off, and that `table_html` is populated for `pdf_table_scan.pdf`.
- **Status**: ⬜ PENDING

### TASK 12 (⚪): Docs + CHANGELOG
- **Agent**: worker (cheap dispatch per registry)
- **Depends on**: TASK_4-11
- **Wave**: 8
- **Contract**: `.workflows/specs/task-docs.spec`
- **Bottleneck**: ⚪ STANDARD
- **What**: Update exactly the paths declared under the proposal's `## Docs impact`:
  `docs/formats.md` (per-page quality semantics, three verdicts, `Escalate` is not an OCR trigger,
  plus the V3 ingestion guidance), `docs/library-usage.md` (`page_quality`, `ocr_page`),
  `README.md` (Limitations bullet + the AGPL design-only decision), `cli/README.md`
  (confirm `--ocr-min-chars` text still matches), and `CHANGELOG.md` — new `0.3.0` heading mentioning
  **Plan 20261108-001**. State plainly that the complexity threshold is calibrated on five fixtures
  and is a heuristic. `docs/v3-ingestion.md` is deliberately **not** touched: it is gitignored.
- **Verify**: `git diff --name-only` lists every declared Docs-impact path, and
  `grep -c '20261108-001' CHANGELOG.md` returns 1.
- **Status**: ⬜ PENDING

(No verify/quality tasks — `/next` verdict-gates every task automatically; `/review` is the final
gate at SHIP.)

## Contracts (generated after approval)

| Task | Contract File | Scenarios |
|------|--------------|-----------|
| TASK_2 | `.workflows/specs/task-2.spec` | 2 evidence scenarios + NO-GO scenario |
| TASK_3 | `.workflows/specs/task-3.spec` | 2 evidence scenarios + deletion-test scenario |
| TASK_4 | `.workflows/specs/task-4.spec` | 4 (incl. crop-offset negative) |
| TASK_5 | `.workflows/specs/task-5.spec` | 3 (incl. `None` for DOCX) |
| TASK_6 | `.workflows/specs/task-6.spec` | 4 (incl. single-column byte-identical) |
| TASK_7 | `.workflows/specs/task-7.spec` | 3 (bounds, monotonicity, deletion) |
| TASK_8 | `.workflows/specs/task-8.spec` | 5 (incl. `Escalate` ≠ `Ocr`, split-pages parity) |
| TASK_9 | `.workflows/specs/task-9.spec` | 3 (both cfgs) |
| TASK_10 | `.workflows/specs/task-10.spec` | 3 (incl. legacy implementor) |
| TASK_11 | `.workflows/specs/task-11.spec` | 3 (identical transcripts, table HTML) |
| TASK_12 | `.workflows/specs/task-docs.spec` | 2 (paths touched, Plan ID present) |

## Execution Notes

Filled in as tasks complete — cost, duration, and learnings per task.

- `lessons:` none applicable — no lessons ledger exists in this repo yet.
- Standing risk: TASK_2 NO-GO re-tiers TASK_4 as `refactor-first` and stops the plan for human
  review. That is the designed behaviour, not a failure.
- Standing risk: TASK_6 touching the heading `level` walk is the most likely surprise. If it breaks,
  escalate rather than patch locally.
- **2026-11-08 finding (pre-contract):** `docs/v3-ingestion.md` is matched by the
  `v3-ingestion.md` pattern in `.gitignore` (`git check-ignore` confirms), so it cannot appear in a
  shipping Docs-impact declaration. Removed from the proposal and from TASK_12; the V3 guidance
  moves into `docs/formats.md`. Same trap applies to `roadmap.md`.
- **2026-11-08 type placement (pre-contract):** `PageGeometry`/`SpanBox` live in
  `src/formats/pdf.rs` (TASK_4); `PageTextQuality`/`PageVerdict` live in `src/quality.rs` (TASK_5).
  This split is what makes parallel group A and B legal — TASK_5 must not define geometry types and
  TASK_7 imports them read-only.
