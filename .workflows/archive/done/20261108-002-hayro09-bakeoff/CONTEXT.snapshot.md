# Context

Shared domain language for office-parser. Keep this file meaningful to domain experts;
avoid implementation trivia.

## Glossary

| Term | Meaning | Notes |
|---|---|---|
| Text layer | The glyphs a PDF carries natively, as opposed to pixels | Present in born-digital PDFs; absent or junk in scans |
| Born-digital page | A page whose text layer is trustworthy | Routed to text extraction, never to OCR |
| Scanned page | Image-only page, or a page whose text layer is untrusted | Routed to OCR |
| `pdf_text_quality` | Document-level classification of extracted text: `TextRich` (>500 alnum), `TextSparse` (50–500), `ImageOnly` (<50) | `src/formats/pdf.rs:2099`, `src/document.rs:62` |
| `page_quality` | **Proposed** per-page quality detail: richness, geometry confidence, layout complexity, verdict | New; see `.workflows/proposals/` for the design |
| Verdict | Per-page routing outcome: `Trust { geometry }`, `Ocr { reason }`, or `Escalate { complexity }` | `Ocr` and `Escalate` are distinct |
| `Escalate` | Page parsed cleanly but is structurally wrong → vision-LLM fallback | **Never** an OCR trigger: OCR cannot fix reading order |
| `layout_complexity` | Deterministic 0.0–1.0 score from columns, table gutters, span overlap, rotation | Heuristic, calibrated on fixtures — not a calibrated probability |
| XY-Cut | Recursive horizontal/vertical cut ordering multi-column text into reading order | Pure geometry, no model |
| Ink coverage | Glyph box area ÷ page area | Distinguishes a scan from a white page without pdfium |
| Trust check | Heuristic that untrusts a text layer (invisible render-mode-3 text, replacement/private-use chars, broken font maps) | Must work with and without `render-pdfium` |
| Geometry confidence | Confidence in extracted box geometry: box coverage, junk-glyph ratio, font sanity | Feeds the verdict |
| Coordinate system | `source` (unrotated page space) or `viewer_oriented` (rotation applied) | Must be reported with any box; crop origin subtracted |
| Page furniture | Headers, footers, page numbers, rules — dropped from output | Distinguish from real content |
| Render budget | `clamp_budget` in `render_pages_png_opts`; OCR needs full-resolution source pixels | `false` for OCR |
| PageOcrEngine | Trait turning one rendered page PNG into blocks + confidence | `src/formats/pdf.rs:2449` |

## Domain Rules

- **Structure vs characters are different failures.** OCR fixes missing characters; it never fixes
  interleaved columns. A page can be `TextRich` and still worthless. Complexity is an `Escalate`
  signal, never an `Ocr` signal.
- **Quality is per page; the document enum is a rollup.** Mixed documents (digital compilation with
  scanned inserts) are the normal case, not the exception.
- **Geometry comes from the text layer, not the raster.** Quality signals must not add a pdfium
  dependency; `render-pdfium`-absent builds are first-class.
- **`None` outside PDF.** Format-specific metadata never leaks into other formats (four-format
  parity discipline, `tests/docx_odt_parity.rs`).
- **Silent default changes are forbidden.** `--ocr-min-chars 50` stays; new signals add `Escalate`,
  they never re-tune `Ocr`.
- **Every box carries its coordinate system and rotation.** A box without provenance is not citable.
- **Heuristic constants are named as heuristics.** A threshold validated on five fixtures is a
  hypothesis; docs must say so.

## Rulings

### Renderer default ruling (plan 20261108-002, 2026-11-09)

**pdfium stays the default OCR renderer; hayro remains the opt-in/fallback**
(`render-hayro` + `OFFICE_PARSER_RENDERER=hayro`). Evidence (bake-off,
`.workflows/spikes/render-bakeoff/report-*.md`): WER on the seeded dense
2-column synthetic corpus — 150 dpi: pdfium 0.6683 vs hayro 0.6964
(+2.81pt); 300 dpi: 0.6734 vs 0.6824 (+0.90pt, inside the 1pt flip rule);
pdfium ~6% faster end-to-end. The flip rule (WER(hayro) <= WER(pdfium)
+ 1pt) is satisfied only at office-grade density; degraded scans keep a
real pdfium edge, and degraded scans are why the OCR path exists.
Revisit-when: upstream hayro lands knockout/blending group passes, the
WER gap closes at 150 dpi, or the deployment corpus standardizes on
>= 300 dpi scans. The vendored stack is upgraded to 0.9.0 + 10 hardening
families regardless (that work is independent of which engine is default).


Verbatim human decisions for this feature. Specs touching these quote them.

- **2026-11-08 — dependencies:** *"i agree with your recommendation for both"*, resolving:
  1. **AGPL-3.0-or-later is not acceptable** for a dependency of this MIT crate. Take `rasgapdf`'s
     *design* only (coordinate contract, XY-Cut reading order, 0.45-style escalation threshold,
     multi-rect chunk citations) and implement it in our own words. Zero dependency.
  2. **`rusto-core` is accepted, later, behind an off-by-default `ocr-rusto-core` feature, scoped to
     tables only.** Never the `rusto-rs` facade crate (`default = ["ffi", "rten"]`). Indra keeps the
     text path; `rusto-core` supplies layout regions, word boxes, and table structure.

## Plan 20261108-001 — durable learnings (sweep 2026-11-08)

- Trust checks never silently degrade: a feature-off build must yield
  either a working signal or a loud no-op — an empty verdict vector is a
  defect class (TASK 9: the no-ink check runs from geometry alone; pdfium
  reasons fold in only where the feature exists).
- Escalate pages are NEVER OCR'd: OCR cannot repair a reading order the
  text layer mangled; escalation is a vision-LLM signal (TASKS 7/8).
- Routing rollup is worst-page-wins: exactly the Ocr-verdict pages go
  through the engine; the document-level quality enum stays a summary.
- OcrPage optional fields are absent — never null — when no OCR'd page
  supplied them (TASK 10).
- Table ownership: Indra produces the text (cells) and the geometry
  (grid bbox, cell rects); the structure model only organizes. A
  structure model that cannot read cells (rusto SLANet stub) is a
  gate, not a source (TASKS 11/13).
- The runtime for OCR-adjacent models is rten (pure Rust); ONNX loads
  through rten's onnx_format. Model vocabularies live in the model's
  own ONNX metadata — extract, never assume (TASK 13 gate: the EN/CH
  dict mismatch).
- Vendoring is the last resort, measured first: the rusto git dep cost
  >1.2 GB via transitive submodules; two MIT crates (452K + 1.4M) were
  vendored, then removed entirely once the logic was ported (TASK 13).

## Open Questions

- Does the existing content interpreter's font metrics survive XY-Cut reordering unchanged?
  Owner: TASK 1p/TASK 6 evidence. Why it matters: heading-level inference depends on it.
- Should `layout_complexity` ever gate OCR? Current ruling: no.


## Open Questions — resolved by TASK 1 scout (2026-11-08, plan 20261108-001)

- **Interpreter state at emit points** (Tj:993/TJ:1017/'/"/BDC:1098): tm+ctm full affine,
  effective_font_size (rotation-safe), current_position (pen origin), leading, font_alias,
  in_text. NOT tracked: render mode, font name (alias only), TJ kerning advances, Tc/Tw, /Rotate,
  MediaBox/CropBox. TextAtom stores only {seq,x,y,font_size,text} — matrices dropped at push.
- **Line assembly drops x** (Line {y,font_size,text}); paragraph split is y-gap-only → column
  separation must live between atoms and line clustering (TASK 4/6 design constraint).
- **Heading inference is order-insensitive** (median/ratio) — survives XY-Cut reordering.
- **Trust checks degrade silently without render-pdfium** (Vec::new() at :2573-2586) → TASK 9.
- **pdf_cjk_scan p1 = 16-char latin-1 mojibake layer** — passes char checks as valid Latin-1;
  defeats threshold intent + bad-char trust check → note for TASK 8/9.
- **Combined-doc page_block_ranges stale after OCR merge** (split-pages uses pages_from_merge,
  unaffected) → defect-class, fold into TASK 8.
- **render.rs JSON + CLI frontmatter omit pdf_text_quality** → TASK 5 must include per-page
  quality in the JSON shape.
- **Routing on all 5 fixtures is fully determined by char counts** (trust deltas nil) — the new
  signals change nothing here; calibration needs richer pages.
- **Registry deviation**: deepseek/deepseek-flash not spawnable this session; runtime assigned
  zai/glm-5.3-flash thinking high for wave 1. Re-probe deepseek before relying on max-thinking.

### Fixture baseline (TASK 1)

| fixture | pages | doc quality | per-page body alnum | trust delta |
|---|---|---|---|---|
| pdf_scanned | 2 | ImageOnly | 0 / 0 | none |
| pdf_table_scan | 1 | ImageOnly | 0 | none |
| pdf_rot90_scan | 1 | ImageOnly | 0 | none |
| pdf_cjk_scan | 1 | ImageOnly (16<50, mojibake) | 16 | none |
| pdf_mixed | 2 | TextSparse | 305 / 0 | none (pages_ocr=1 both builds) |

Raw: `.workflows/spikes/scout/baseline/` (counts.txt, body-counts.txt, test-results.txt,
per-fixture md, mixed-*-ocr json).
