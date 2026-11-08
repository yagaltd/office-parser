# PDF Text Quality and Page Verdicts

How `office-parser` assesses PDF text quality per page and routes pages
through OCR. Applies to the `pdf` feature; `pdf_text_quality` and
`page_quality` are `None` for every non-PDF format.

## Document-level summary

`metadata.pdf_text_quality` is the classic char-count enum (`TextRich`
> 500 alphanumeric chars, `TextSparse` 50–500, `ImageOnly` < 50). It stays
a plain summary — per-page routing no longer derives from it.

## Per-page quality

Every PDF parse path (`parse()`, the OCR routers, split pages) emits one
`PageTextQuality` per page in page order in `metadata.page_quality`,
mirrored to `metadata.extra["pdf_quality"]` for JSON consumers:

```json
[
  {
    "page": 1,
    "text_richness": 310,
    "columns": 1,
    "rotated": false,
    "verdict": { "kind": "trust" }
  },
  {
    "page": 2,
    "text_richness": 0,
    "columns": 0,
    "rotated": false,
    "verdict": { "kind": "ocr", "reason": "image-only page" }
  },
  {
    "page": 3,
    "text_richness": 214,
    "columns": 1,
    "rotated": true,
    "verdict": { "kind": "escalate", "complexity": "rotated page" }
  },
  {
    "page": 4,
    "text_richness": 604,
    "columns": 2,
    "rotated": false,
    "verdict": { "kind": "ocr", "reason": "no-ink" }
  }
]
```

- `page`: 1-based page number.
- `text_richness`: alphanumeric characters extracted from the page's text
  layer (`0` for image-only pages).
- `columns`: text-column count from the span-box geometry (`1` single
  column, `0` when the page has no text spans).
- `rotated`: the page carries a non-zero native `/Rotate`. A
  content-stream tilt is not recoverable from axis-aligned span boxes, so
  `/Rotate` is the standalone flag.
- `verdict`: routing decision, see below.

With `--split-pages` / `parse_pdf_split`, each per-page document carries
its own single-record array.

## Verdicts

`PageVerdict` is internally tagged JSON on `kind`:

- `trust` — geometry and text layer agree; use the extracted text as-is.
- `ocr { reason }` — the page's text layer is insufficient; the page is
  routed through the OCR engine when one is attached. Reasons:
  - `"image-only page"` — the page's text layer has zero characters.
  - `"sparse text layer (N chars < M)"` — under the char gate
    (`--ocr-min-chars`, default `50`).
  - `"no-ink"` — the pdfium-free geometry trust check: the page claims
    rich text (`text_richness >= min_chars`) but its span boxes hold less
    than 25% of the ink area that char count implies (typical producers:
    invisible OCR layers over scans, spans with degenerate text matrices).
    Runs from text-layer geometry alone — no pdfium required — so trust
    verdicts are never silently empty in `render-pdfium`-absent builds.
  - pdfium reasons (feature `render-pdfium`, additional signal):
    `"invisible text N% (OCR layer)"` (> 5% render-mode-3 text) and
    `"broken font map (N% bad chars)"` (replacement/private-use/control
    characters).
- `escalate { complexity }` — the text layer is readable but the layout
  needs more than OCR (shipped signal: `"rotated page"`; `complexity` is
  `null` when no single signal is known). `escalate` pages are NEVER
  routed to OCR — OCR cannot repair a reading order the text layer
  mangled; escalation is a vision-LLM signal.

Decision order: the image-only gate first, then the sparse gate, then
rotation escalation; the no-ink check overrides a would-be `trust`
verdict, and pdfium trust reasons (where present) fold in last and can
flip any verdict to `ocr`.

Heuristic status: the no-ink constants (0.55 advance factor per glyph,
11 pt nominal body font, 0.25 minimum coverage ratio) and the column
gutter rule are calibrated on the five repository fixtures
(`pdf_scanned`, `pdf_mixed`, `pdf_table_scan`, `pdf_rot90_scan`,
`pdf_cjk_scan`) — they are heuristics, not calibrated probabilities.

## Routing, rollup, force

- A page is OCR'd iff forced or its own verdict is `ocr` (worst-page-wins
  rollup over the per-page records). `escalate` never adds a page.
- The document-level `pdf_text_quality` stays the summary enum; the
  per-page verdicts drive routing.
- Force (`--ocr-force` / `force = true`) bypasses every trust check —
  `no-ink` and the pdfium reasons — and OCRs every page. The image-only /
  sparse gates and rotation escalation are not trust checks: verdicts
  still record them.
- Verdicts are computed and reported even when no engine is attached;
  routing only happens when an OCR engine is supplied (CLI `--ocr`,
  library `pdf_ocr::parse_auto*` / `pdf::parse_with_page_ocr*`).
- `metadata.extra["ocr"]` appears only when at least one page was OCR'd.

## OCR payload (`metadata.extra["ocr"]`)

Always present: `engine`, `pages_ocr`. Plus:

- `low_confidence_pages`: pages whose mean line confidence fell below
  0.60 — the vision-LLM fallback signal.
- `page_reasons`: `[{ "page": N, "reason": "..." }]` when a trust check
  fired.
- Rich per-page keys from `OcrPage` (engines may supply them; absent —
  never `null` — when no OCR'd page did):
  - `layout_regions`: `[{ "page": N, "regions": [{ "label": "Table",
    "x0": …, "y0": …, "x1": …, "y1": … }] }]`
  - `word_boxes`: `[{ "page": N, "words": [{ "text": "…", "x0": …, "y0": …, "x1": …, "y1": … }] }]`
  - `table_html`: `[{ "page": N, "html": "<table>…</table>" }]`

Region and word coordinates are pixels in the rendered page PNG
(top-left origin, y down), rasterized at the requested dpi with no
megapixel clamp (OCR wants full-resolution source pixels).

## Table enrichment (`ocr-rusto-core`, opt-in)

Off by default. When enabled, OCR'd pages are enriched after the engine
ran: rusto-core's `LayoutDetector` detects `LayoutType::Table` regions
(surfaced via `layout_regions`, label `"Table"`) and the engine's own
table blocks render as `<table>` markup into `table_html` (gated on a
detected table region). Blocks and confidence pass through untouched —
transcripts are identical with the feature on or off. `word_boxes` are
never populated.

Model resolution: `OFFICE_PARSER_LAYOUT_MODEL` (exact file path) wins,
else `models/DocOCR/layout.onnx` relative to the process working
directory. Without a model file the enrichment is a documented no-op:
one warning, rich keys stay absent (the feature-off payload shape).
