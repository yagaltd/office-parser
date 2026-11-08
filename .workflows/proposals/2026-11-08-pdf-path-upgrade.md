# Proposal: geometry-aware PDF page quality + optional richer OCR page contract

- Date: 2026-11-08
- Status: PROPOSAL — investigation complete, zero code, zero decisions locked
- Origin: comparison of `rasgapdf` v0.4.2 (pure-Rust spatial PDF parser) and `rusto-rs` v0.4.0 (pure-Rust OCR engine) against this crate
- Plan ID (reserved, assigned for real at `/idea` bootstrap): `20261108-001`
- Doctrine reminder: every 🟡 item below needs a prototype task with evidence on disk before wiring. No dependency is added without the license gate in §9.

## Problem

`pdf_text_quality` answers one question — "how many alphanumeric characters did we get?" — and it
answers it **per document, not per page**, from a signal that cannot see structure.

```rust
// src/formats/pdf.rs:2099
fn classify_pdf_text_quality(blocks: &[Block]) -> PdfTextQuality {
    let n = count_alphanumeric(blocks);          // sums ALL pages
    if n > 500 { PdfTextQuality::TextRich }
    else if n >= 50 { PdfTextQuality::TextSparse }
    else { PdfTextQuality::ImageOnly }
}
```

Four consequences, in order of severity:

1. **`TextRich` is not "good".** A three-column academic PDF read as one linear stream yields
   3,000 alphanumeric characters → `TextRich`, while the reading order is interleaved garbage.
   Character count measures *presence*, never *arrangement*.
2. **Document-level aggregation hides mixed documents.** One scanned insert in a 200-page digital
   compilation cannot be located from a single enum. Page routing already works per page
   (`page_char_counts`, `pdf.rs:2520`), but the *reported* quality does not — the two disagree.
3. **The geometry gap is recorded as a TODO in our own source.** `pdf.rs:2571`:
   `(no-ink check pending: needs char-box geometry.)` Without char boxes we cannot tell a scan
   from a white page, so we depend on `pdfium` (and silently get `Vec::new()` trust verdicts when
   it is absent — `pdf.rs:2581`).
4. **No complexity signal reaches the vision-LLM fallback.** We already emit low-confidence pages
   (`conf < 0.60` → `metadata.extra["ocr"].low_confidence_pages`), but never "this page parsed
   cleanly and is still structurally wrong" — the exact case where escalating to a VLM pays.

## Evidence

- Quality classifier, thresholds, and document-level aggregation: `src/formats/pdf.rs:2099-2112`,
  `src/document.rs:62-71`, `src/formats/pdf.rs:2432`, `:2566`, `:2684`, `:2703`.
- Per-page char counts already exist and drive routing: `pdf.rs:2520-2526`, filter at `:2586-2601`.
- Trust checks and their pdfium dependency: `pdf.rs:2568-2585` (comment + `#[cfg(feature =
  "render-pdfium")]` / `#[cfg(not(...))] { Vec::new() }`).
- Header TODO this proposal closes: `pdf.rs:2571`.
- The content interpreter already tracks what geometry needs: `state.tm`, `state.font_size`
  (`pdf.rs:925`, `:1005`, `:1024`, `:1059`), `effective_font_size()` (`:291`),
  `estimate_text_advance()` (`:471`), heading `level` walk (`:739`, `:807`, `:853-870`).
- OCR seam to widen: `PageOcrEngine` (`pdf.rs:2449-2452`), `&mut dyn` usage (`:2457`), single call
  site (`:2647`).
- Fixtures available for the prototypes, no new corpus needed: `tests/fixtures/pdf_scanned.pdf`,
  `pdf_table_scan.pdf`, `pdf_rot90_scan.pdf`, `pdf_cjk_scan.pdf`, `pdf_mixed.pdf`.
- Existing gate to keep green: `tests/pdf_ocr.rs` (6 tests).

External, for the record — **not adopted as code, see §9**:

- `rasgapdf` v0.4.2, AGPL-3.0-or-later: Recursive XY-Cut reading order, CropBox-offset + `/Rotate`
  contract (`coordinate_system: source|viewer_oriented`, `initial_rotation` preserved),
  deterministic geometric tables (Y-band `row_tolerance 6.0pt` → X projection `column_tolerance
  16.0pt` → `min_table_columns 3`), `complexity_score` with a **0.45** escalation threshold,
  `is_complex_layout` per page and per block. ~9 ms/page, 5 deps, no ML.
- `rusto-rs` v0.4.0, MIT: `rusto-core` is usable standalone (7 transitive deps) and adds 10-class
  PicoDet layout, 3D UVDoc unwarp, RT-DETR + SLANet table structure → HTML, word-level boxes.
  It runs **the same Indra models we already use** (45.98 % OmniDocBench, identical transcripts) —
  so it improves *structure*, never *characters*.
- Neither renders PDFs. `render_pages_png_opts(bytes, dpi, clamp_budget=false)` stays ours
  (pdfium, or the vendored hayro stack).

## Refactor assessment (mandatory)

**Scan** — files this plan touches and what already covers the need:

| Existing | Covers | Verdict |
|---|---|---|
| `classify_pdf_text_quality` + `PdfTextQuality` | char-count quality, document level | **extend** — add geometry/complexity, keep the enum |
| `page_char_counts` + OCR filter (`pdf.rs:2586`) | per-page routing | **extend** — feed it the complexity signal |
| `pdfium_text_layer_reasons` trust checks | scan detection when pdfium present | **extend** — geometry makes the no-ink check pdfium-free |
| `PageOcrEngine` + `parse_with_page_ocr_inner` | per-page OCR seam | **extend** — defaulted methods, no signature break |
| `indra_native.rs` (1,905 LOC) | scanned-page OCR (det/rec/order/orient/deskew/CJK) | **untouched** |
| `render_pages_png_opts` | rasterization | **untouched** |

**Verdict: `extend`.** Every capability above has a working flow; the gap is signal, not structure.
No refactor lands first. Nothing here duplicates `pdf_text_quality` — it enriches it.

**Block rule acknowledged:** if a mid-plan prototype shows the content interpreter *cannot* emit
per-span boxes without rewriting the text path, that contradicts this verdict → stop and escalate,
do not absorb.

## Proposed changes

Three independent tracks. Tracks 1 and 2 are pure-Rust, dependency-free, and land first; Track 3 is
optional and gated.

### Track 1 — page geometry (unblocks the header TODO)

Emit, per page, a `PageGeometry`: one record per span with `bbox [x0,y0,x1,y1]`, `font_name`,
`font_size`, derived from the interpreter state we already track. Store `coordinate_system`
(`source` | `viewer_oriented`), `initial_rotation`, `media_box`, `crop_box` with the crop origin
subtracted. Then:

- **XY-Cut reading order** for born-digital multi-column pages, replacing the current linear
  stream join. This is the fix for consequence (1) — and it is pure geometry, no model, no license.
- **Ink coverage**: `glyph_area / page_area` from char boxes. Closes the header TODO and makes the
  no-ink scan check work without pdfium (consequence 3).

### Track 2 — page quality becomes a struct (the actual ask)

Keep `PdfTextQuality` unchanged for compatibility; add a sibling detail:

```
page_quality: Vec<PageTextQuality>      // new, per page, Pdf-only; None elsewhere
PageTextQuality {
    page: usize,
    text_richness: PdfTextQuality,          // existing signal, now per page
    geometry_confidence: f32,               // box coverage, junk-glyph ratio, font sanity
    layout_complexity: f32,                 // 0..1
    verdict: PageVerdict,                   // Trust { geometry } | Ocr { reason } | Escalate { complexity }
    columns: u8,                            // from XY-Cut
}
```

`layout_complexity` is **deterministic**, from signals we already have or get in Track 1: column
count, table gutter density (geometric projection, Track 1 boxes), span-overlap ratio, rotation
≠ 0, char count. The 0.45 escalation threshold is borrowed as a *starting* constant and must be
calibrated on our five PDF fixtures in a prototype before it gates anything.

Document-level rollup: worst page wins for routing; the existing document enum stays as the
"best page" summary. `metadata.extra["pdf_quality"]` carries the per-page array for consumers that
want detail without a struct change.

**Non-negotiable:** complexity alone never triggers OCR. Escalation is a *vision-LLM* signal
(`Escalate`), separate from `Ocr`, because OCR cannot fix reading order.

### Track 3 — richer OCR page contract (optional, gated)

Widen the seam with defaulted methods so no downstream impl breaks:

```rust
// PageOcrEngine gains, with defaults that return None:
fn ocr_page(&mut self, png: &[u8]) -> Result<OcrPage, String> { /* wraps ocr_page_png */ }
// OcrPage { blocks, confidence, layout_regions: Option<..>, word_boxes: Option<..>, table_html: Option<..> }
```

Then, behind an **off-by-default** `ocr-rusto-core` feature, implement `layout_regions`,
`word_boxes` and `table_html` via **`rusto-core` only** — never the `rusto-rs` facade crate
(`default = ["ffi", "rten"]` drags `uniffi`, `cbindgen`, `ureq`, `clap`). Indra keeps the text path;
`rusto-core` is used for `LayoutType::Table` regions and table structure only.

## Decision tree

1. **Is AGPL-3.0-or-later acceptable for a dependency of an MIT crate served over a network?**
   - Recommended: **no**. Take `rasgapdf`'s *design* (coordinate contract, XY-Cut, complexity
     thresholds, multi-rect chunks) and implement in our own words. Zero dependency, zero license
     change, same behaviour.
   - Evidence: `rasgapdf` `license = "AGPL-3.0-or-later"`; this crate is MIT (`Cargo.toml`);
     AGPL §13 reaches hosted ingestion services.
   - If you choose otherwise: office-parser and CognitiveOS V3 must relicense, or a commercial
     license must be bought before any code is copied.
   - Blocks: Track 1's adoption route (not its design).
2. **Do we add `rusto-core` at all?**
   - Recommended: **yes, later, behind `ocr-rusto-core`, off by default**, scoped to tables.
   - Evidence: 7 transitive deps; same Indra models ⇒ no recognition gain; only SLANet tables and
     word boxes are new. Our README already flags complex tables as a limitation.
   - If you choose otherwise: Track 3 is dropped and table quality stays as-is. Nothing else changes.
   - Blocks: Track 3 only.
3. **Does `PdfTextQuality` change shape?**
   - Recommended: **no** — additive `page_quality` + `metadata.extra`. Keeps 0.2.0 consumers compiling.
   - If you choose otherwise: breaking major bump and a migration note in README.
   - Blocks: Track 2's public surface.
4. **Who owns the 0.45 constant?**
   - Recommended: calibrated in the Track 2 prototype, documented as a heuristic, overridable.
   - Blocks: Track 2's routing test.

## Prototype tasks (required before the build tasks that depend on them)

- **TASK 0p (prototype, 🟡)** — *Does our existing content interpreter already have enough state to
  emit per-span boxes and ink coverage without pdfium?*
  Allowed changes: `.workflows/spikes/pdf-geometry/**`. Evidence: box dump for
  `pdf_scanned.pdf` vs `pdf_mixed.pdf`; measured ink coverage separating scan from digital page.
  This is a bounded question with a binary answer and it decides whether Track 1 is `extend` or
  `refactor-first`.
- **TASK 1p (prototype, 🟡)** — *Does a deterministic `layout_complexity` separate our fixtures the
  way a human would?* Evidence: score for all five PDF fixtures + a table of which page the score
  would route to `Escalate`; calibration note for the threshold.

## Blindspot pass

1. **`render-pdfium`-absent builds must stay first-class.** Today trust checks degrade to
   `Vec::new()`. Track 1/2 must not introduce a pdfium dependency for quality — geometry comes from
   the text layer, not the raster. Evidence: `pdf.rs:2572-2585`.
2. **`--split-pages` has its own quality path** (`pdf.rs:2703` sets per-page docs). Per-page
   `page_quality` must agree with it, or the two modes disagree — a silent inconsistency bug.
3. **Four-format parity.** `pdf_text_quality` is `None` outside PDF. `page_quality` must be too, and
   `tests/docx_odt_parity.rs` discipline applies.
4. **New public structs are API surface.** `PageTextQuality` and `PageVerdict` need serde derives and
   a documented schema, or `to_json_value()` consumers break on shape drift.
5. **`--ocr-min-chars 50` stays the default.** Changing the default silently changes output for
   every existing user. New signals add `Escalate`, they do not re-tune `Ocr`.
6. **Fixture honesty.** Five fixtures is a smoke test, not a corpus. Complexity thresholds validated
   on five files are a hypothesis — say so in docs, do not present them as calibrated.
7. Likely safe assumptions
   - The heading `level` walk keeps working on top of XY-Cut ordering — verify in Track 1 tests.
   - `PdfTextQuality` semantics (500/50) stay unchanged for compatibility — safe, additive only.

## Docs impact (mandatory)

```
Docs impact:
- `docs/formats.md` — PDF quality is now per-page: document the `page_quality` semantics,
  the three verdicts, and that complexity escalation is a vision-LLM signal, not an OCR trigger.
  Also carries the V3 ingestion guidance for `Escalate` pages.
- `docs/library-usage.md` — new `DocumentMetadata.page_quality` field and the
  `PageOcrEngine::ocr_page` defaulted method.
- `README.md` — Limitations bullet: quality is per-page and geometry-aware; state the AGPL
  decision (design-borrowed, no dependency).
- `cli/README.md` — no new flags in this plan; confirm `--ocr-min-chars` text still matches.
- `CHANGELOG.md` — 0.3.0: "PDF text quality is per-page and geometry-aware; XY-Cut reading
  order for multi-column born-digital pages (Plan 20261108-001)".

`docs/v3-ingestion.md` was considered and **excluded**: it is matched by the
`v3-ingestion.md` pattern in `.gitignore`, so it cannot be part of a shipping Docs-impact
declaration (the content lands in `docs/formats.md` instead).
```

## Explicitly NOT decided

- Whether `layout_complexity` ever gates OCR (proposed: never).
- Whether Track 3 ships in 0.3.x or 0.4.x (depends on Decision 2).
- Whether the heading-level walker survives XY-Cut unchanged (Track 1 prototype decides).
- Adopting any `rasgapdf` or `rusto-rs` code — both blocked pending §9 Decisions 1 and 2.
- Multi-rect chunk citations (`Vec<BoundingBox>` per chunk) — valuable, but belongs in
  `render::to_chunks`, i.e. a separate proposal.
