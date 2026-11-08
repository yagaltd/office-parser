# Formats and Behavior

## Office documents

- `docx`, `odt`: headings, paragraphs, lists, tables, images, links; chart/diagram semantics when detectable.
- `pptx`, `odp`: slide headings + text/image/table content, chart and diagram semantics.

## Spreadsheets

- `xlsx`, `ods`, `csv`, `tsv` are parsed into semantic tables.
- Large sheets are constrained by row/column/cell limits.

## Other formats

- `pdf`: text extraction and page metadata.
- `rtf`: paragraph/list/table/image extraction.

## PDF quality and OCR

- Per page: `metadata.page_quality` (JSON mirror `extra["pdf_quality"]`)
  holds `page`, `text_richness`, `columns`, `rotated`, and a `verdict`.
- Born-digital pages are ordered by recursive XY-Cut over span boxes —
  multi-column pages read column-by-column; single-column output is
  byte-identical to the previous linear join.
- Verdicts: `trust` (use as-is), `ocr { reason }` (routed to OCR when an
  engine is attached), `escalate { complexity }` (vision-LLM signal —
  never OCR'd).
- Reasons: `image-only page`, `sparse text layer (N chars < M)`,
  `no-ink` (text layer claims ink its span geometry cannot hold; checked
  without pdfium), and with `render-pdfium` also invisible-text /
  broken-font-map reasons.
- `--ocr-force` bypasses the trust checks and OCRs every page.
- OCR'd pages report `metadata.extra["ocr"]`: `engine`, `pages_ocr`,
  `low_confidence_pages`, plus rich keys `layout_regions`,
  `word_boxes`, `table_html` when a page supplied them.
- Shapes, reasons, and examples: [quality.md](quality.md).

(The `- rtf:` line is unique in the file; inserting after it keeps the new section inside "Other formats" flow before "## Config/data formats".)
- `epub`: spine XHTML and embedded images.
- `xmind`, `mmap`: mindmap topic trees mapped to heading + nested list blocks, with full tree in `metadata.extra.mindmap`.

## Config/data formats

- `json`, `yaml`, `toml`: converted to heading/KV/list/table/jsonl-style block structure.
- `xml`: generic XML mapping to heading/KV/list/table blocks; attributes as `@attr`; CDATA HTML converted to markdown best-effort.

## Known limits

- Slide snapshot image generation is not currently implemented.
- Chart tables depend on cached values embedded by source apps.
- Complex visuals (SmartArt, hand-drawn diagrams) are intentionally not fully interpreted.
