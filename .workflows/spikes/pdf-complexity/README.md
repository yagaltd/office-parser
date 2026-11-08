# spike: pdf-complexity — TASK 3 (plan 20261108-001)

**STATUS: MEASURED (orchestrator executed the spike author's command; see Score table).**
The spike author had no execution tools; sources were materialized verbatim by the
orchestrator and run once. Nothing in `src/` was touched by this spike.

## Question

Does a deterministic `layout_complexity` built from span boxes (horizontal gutter count, table
gutter density, span-overlap ratio, rotation ≠ 0, char density) rank the five fixtures the way a
human would — high for multi-column text, low for single-column text, ambiguous for scans? If not,
`columns` alone wins the deletion test and the score is dropped.

## Methodology (span boxes only; no pdfium, no model)

`main.rs` mirrors `src/formats/pdf.rs`'s proven text-layer walk (lopdf 0.35: Tm/cm/Td/TD/TL/T*/
Tj/TJ state machine, `estimate_text_advance`, NUL-stripping font decode), then:

1. groups atoms into lines in a rotation-aligned frame (u along the dominant baseline,
   v perpendicular); page rotation = median angle of tilted atoms when they carry ≥50% of chars;
2. computes five signals per page:
   - **horizontal gutter count** — uncovered vertical strips of width ≥ max(2×median font size,
     12pt) in the union of *body-line* x-intervals (body = lines ≤70% of text width, so
     full-width headings don't mask column gutters);
   - **table gutter density** — intra-line gaps ≥ max(0.6×font size, 2.5pt) clustered by x
     (tolerance max(0.5×median font, 4pt)); "strong" cuts are shared by ≥40% of lines (≥6 lines
     required); density = strong_cuts/2, clamped to 1;
   - **span-overlap ratio** — fraction of line boxes colliding in both axes;
   - **rotation ≠ 0** — 1.0 when the dominant tilt exceeds 5°;
   - **char density** — non-space chars ÷ text-bbox area (pt²), saturating at 0.02 chars/pt².

## Score formula and weights

```
layout_complexity = clamp01(0.30·rot + 0.25·gutters_norm + 0.20·table_density
                          + 0.10·overlap_norm + 0.15·density_norm)      # 0.0..=1.0
gutters_norm    = min(1, gutters/2)
table_density   = min(1, strong_cuts/2)
overlap_norm    = min(1, overlap_ratio/0.3)
density_norm    = min(1, char_density/0.02)
```

## Score table

Measured by orchestrator per spike author's command (see §Reproduce); `data/scores.json` holds
the provenance dump.

```
score = clamp01(0.30*rot + 0.25*gutters_norm + 0.20*table_density + 0.10*overlap_norm + 0.15*density_norm); escalate at >= 0.30

| fixture | page | text_layer | chars | rot_deg | lines | gutters | aligned_cuts | strong_cuts | table_density | overlap_ratio | char_density | columns | score | verdict |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| pdf_scanned.pdf | 1 | false | 0 | - | 0 | - | - | - | - | - | - | - | (undefined) | Escalate (ImageOnly via pdf_text_quality; score undefined) |
| pdf_scanned.pdf | 2 | false | 0 | - | 0 | - | - | - | - | - | - | - | (undefined) | Escalate (ImageOnly via pdf_text_quality; score undefined) |
| pdf_table_scan.pdf | 1 | false | 0 | - | 0 | - | - | - | - | - | - | - | (undefined) | Escalate (ImageOnly via pdf_text_quality; score undefined) |
| pdf_rot90_scan.pdf | 1 | false | 0 | - | 0 | - | - | - | - | - | - | - | (undefined) | Escalate (ImageOnly via pdf_text_quality; score undefined) |
| pdf_cjk_scan.pdf | 1 | true | 21 | 0.0 | 1 | 0 | 0 | 0 | 0.00 | 0.000 | 0.0126 | 1 | 0.095 | Keep text layer |
| pdf_mixed.pdf | 1 | true | 310 | 0.0 | 8 | 0 | 0 | 0 | 0.00 | 0.000 | 0.0052 | 1 | 0.039 | Keep text layer |
| pdf_mixed.pdf | 2 | false | 0 | - | 0 | - | - | - | - | - | - | - | (undefined) | Escalate (ImageOnly via pdf_text_quality; score undefined) |
```

New fact: `pdf_cjk_scan` p1 carries a real 21-char text layer (below the 50-char routing
threshold — it still routes to OCR via the char gate).

## Threshold

Proposed: **escalate when `score ≥ 0.30` (inclusive); pages with `text_layer = false` escalate via
the existing `pdf_text_quality = ImageOnly` signal, not via this score** (their score is undefined,
not 0 — reporting a fake 0 would be silent-pass evidence, a defect class per AGENTS.md).

Gate confirmation against the measured table: recorded in §Measured verdict below.

**This threshold is a starting value measured against five fixtures — a hypothesis, not a
calibrated probability.** The calibration note: signals are measurable only on pdf_mixed's digital
page(s); the other four fixtures are image-only scans invisible to span-box signals.

## Deletion test

Gate-checked against the measured table — see §Measured verdict.

## Reproduce

From the repo root (single command):

```
cargo run --release --offline --manifest-path .workflows/spikes/pdf-complexity/Cargo.toml
```

Writes `data/scores.json` + `data/score-table.md` under the spike dir and prints the table. To
score other files: append paths as args.

## Measured verdict

GATE CHECK (orchestrator, 2026-11-08):

- Threshold: confirmed. All text-layer pages score < 0.30 (0.039/0.095); no false escalations;
  ImageOnly pages escalate via the existing signal with score undefined (no fake 0s).
- Deletion test: **branch 1 wins.** Measured `gutters = 0` on every text-layer page
  (pdf_mixed's digital page is single-column, as the author predicted). Score and `columns`
  agree on every page → **TASK 7 scope shrinks to: `columns` (gutter count) + `pdf_text_quality`
  routing + optional standalone rotation flag read off the Tm matrix.** The five-term score is
  dropped (deletion, not failure — the spec's own terms).
- Materialization note: 4 small compile fixes were needed to run the delivered sources
  (`&Option<Signals>` borrow, missing `Default` for TextState, lopdf 0.35 `&Dictionary` clone,
  page_fonts iteration by reference). All inside the spike dir; `src/` untouched
  (`git diff --stat src/` empty).
