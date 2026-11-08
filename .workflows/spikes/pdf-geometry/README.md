# Spike: per-span boxes + ink coverage without pdfium (TASK 2, plan 20261108-001)

Spec: `.workflows/specs/task-2.spec`. Throwaway — never merged. Status: **GO on the seam
("Track 1 is `extend`")**; the ink-coverage question resolves mechanically via the pre-stated
Decision rule below (one measured number, pasted into the marked slot at gate time).

## Reproduce (one command — every number comes from this)

```
cargo run --quiet --manifest-path .workflows/spikes/pdf-geometry/Cargo.toml --release -- tests/fixtures/pdf_scanned.pdf tests/fixtures/pdf_mixed.pdf
```

No `render-pdfium`, no `render-hayro`, no OCR features. Deps: `lopdf 0.35` (workspace lock:
0.35.0) + `anyhow`. On an offline machine append `--offline` (crates are in the local cargo
cache). `git diff --name-only src/` stays empty — the spike only reads.

## Box dump

> ▼ PASTE GATE OUTPUT HERE: the `BOX DUMP — first 20 spans ...` block (measured by orchestrator
> per spike author's command). Contract check: each line carries all seven fields —
> `x0 y0 x1 y1 font_name font_size "text"`.
```
BOX DUMP — first 20 spans of tests/fixtures/pdf_mixed.pdf page 1 (fields: x0 y0 x1 y1 font_name font_size "text"):
  [  0] x0=   72.00 y0=  720.00 x1=  281.88 y1=  738.00 font=R7           size= 18.00 text="Quarterly Field Report"
  [  1] x0=   72.00 y0=  680.00 x1=  259.33 y1=  693.00 font=R7           size= 13.00 text="Network Maintenance Summary"
  [  2] x0=   72.00 y0=  650.00 x1=  401.12 y1=  661.00 font=R8           size= 11.00 text="During the second quarter the field team replace..."
  [  3] x0=   72.00 y0=  635.00 x1=  391.66 y1=  646.00 font=R8           size= 11.00 text="patch panels across the northern ring. Downtime ..."
  [  4] x0=   72.00 y0=  620.00 x1=  355.58 y1=  631.00 font=R8           size= 11.00 text="eleven hours, split over four maintenance window..."
  [  5] x0=   72.00 y0=  585.00 x1=  361.19 y1=  596.00 font=R8           size= 11.00 text="Customer impact was limited to two buildings on ..."
  [  6] x0=   72.00 y0=  570.00 x1=  366.25 y1=  581.00 font=R8           size= 11.00 text="Riverside campus. Affected tenants received cred..."
  [  7] x0=   72.00 y0=  555.00 x1=  293.65 y1=  566.00 font=R8           size= 11.00 text="according to the standard SLA schedule."
```

Field provenance — every value comes from state the production interpreter already holds at its
existing emit points (all cites `src/formats/pdf.rs`):

| Span field | Interpreter source |
|---|---|
| `x0`, `y0` | `state.current_position()` :283-288 = `(ctm×tm).apply_to_origin()` :263-265 — used at emit Tj :1007, TJ :1034 |
| `font_size` | `state.effective_font_size()` :290-293 (via `tm.scale_hint()` :267-272) — Tj :1006, TJ :1025 |
| `font_name` | `state.font_alias` — set by `Tf` :923, read at emit :998 / :1024 |
| `x1` (advance) | `estimate_text_advance(text, fs)` :471-485 — the estimator production already uses |
| direction (rotation) | `Matrix2D` components :219-256 from `state.ctm × state.tm` |
| state inputs | `PdfTextState` :208-216; `Tm` :933-947, `Td` :964-973, `TD` :974-984, `T*` :990-993, `cm` :948-963, `q/Q` :909-916 |
| emit points | `Tj` :994-1017 · `TJ` :1018-1047 · `'` :1048-1075 · `"` :1076-1099 · `BDC`/ActualText :1100-1133 |
| per-page walk | `parse_pdf_full` :1658-1701 (`get_pages` :1660, `get_page_fonts` :1680, encodings :1692-1697, `handle_content_ops` call :1700) |
| extend seam | `TextAtom` :199-205 already carries `(seq, x, y, font_size, text)` per span — boxes are additive fields at the same emit points |

## Ink coverage table

> ▼ PASTE GATE OUTPUT HERE: the `COVERAGE TABLE (tab-separated)` block (measured by orchestrator
> per spike author's command).
```
fixture	page	spans	glyph_area_pt2	page_area_pt2	coverage
tests/fixtures/pdf_scanned.pdf	1	0	0.00	484768.91	0.000000
tests/fixtures/pdf_scanned.pdf	2	0	0.00	484768.91	0.000000
tests/fixtures/pdf_mixed.pdf	1	8	25325.07	484704.00	0.052249
tests/fixtures/pdf_mixed.pdf	2	0	0.00	484766.26	0.000000
```

Pre-measured structural facts (proven by byte-grep over the whole files, WITHOUT running):

- **pdf_scanned.pdf — coverage is exactly 0.000000 on both pages.** The file contains no
  `/Font` object and no `BT`/`Tj` bytes anywhere (rg over the raw file: zero matches; page 1's
  content stream is stored uncompressed and reads `q 612.041 0 0 792.053 0 0 cm /Im0 Do Q`).
  Both pages `MediaBox [0 0 612.041 792.053]` → page area 484,776.60 pt², zero glyph boxes.
- **pdf_mixed.pdf page 1** (`4 0 obj`): born-digital text page — `/Font 10 0 R` in resources,
  `MediaBox [0 0 612 792]` (484,704 pt²) → coverage = the one number the run adds.
- **pdf_mixed.pdf page 2** (`11 0 obj`): image-only page (`ProcSet [/PDF /ImageC]`, XObject
  `14 0 R` 1275×1650 DeviceGray, no `/Font`) → coverage = 0 structurally; the run confirms.

## Decision rule (pre-stated — the GO/NO-GO stands or falls on this rule, per orchestrator protocol)

- scanned range: exactly {0.000000, 0.000000} (structural, above).
- mixed range: {measured p1, 0.000000 (p2)}.
- **Rule: ink coverage identifies a scan without pdfium iff coverage(mixed p1) > 0.005 while
  scanned == 0.** Threshold basis: a single short text line (~300 chars at 11 pt ≈ 19,800 pt² ≈
  0.041 of page area) clears it 8×; a lone 5-char watermark word (~330 pt² ≈ 0.0007) would not —
  that boundary is stated, not hidden.
- If measured > 0.005 → GO: "ink coverage identifies a scan without pdfium" holds for the plan's
  fixtures. If ≤ 0.005 → NO-GO for the coverage question only (refactor required: glyph-metric
  advances + `Tc`/`Tw`/`Tz` tracking); the seam verdict below is unaffected.

## Verdict

**GO — Track 1 is `extend`.** Every field of a per-span box (`x0,y0`: current_position :283-288;
size: effective_font_size :290-293; name: font_alias :923/:998; `x1`: advance :471-485 plus
matrix direction :219-256) is already held at the existing emit points :994-1133. The spike
replicates the walk and adds only box arithmetic (`push_span`) on top of state production
already maintains; no new text path is proposed. The ink-coverage verdict resolves via the
Decision rule once the one command runs; its structural half (scanned == 0) is already proven.

## Caveats (bounded, stated)

- `Tc`/`Tw`/`Tz` (char/word spacing, horizontal scaling) have no arms in :905-1199 → advance
  error per span is bounded (~±15% typical); irrelevant to a 0-vs-threshold decision.
- Production does not advance `tm` after show ops → consecutive `Tj` without an intervening
  positioning op share an origin (overlaps double-count in the raw coverage sum). Both fixtures
  position every show op.
- Coverage = Σ span-box area (advance × em) ÷ MediaBox area — a typographic-box metric, not
  raster ink.
- Rotated text: box is the axis-aligned bbox of the 4 corners (slight overestimate); both
  fixtures are unrotated.

## Boundaries compliance

Files: this directory only. `git diff --name-only src/` is empty. No pdfium, no hayro, no OCR.
No test framework introduced.
`
