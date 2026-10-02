# office-parser

[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Rust crate for parsing documents into a normalized `Document` AST that preserves useful hierarchy and semantic structure for downstream ingestion, chunking, retrieval, and inspection.

`office-parser` is for office-like documents and structured data files:
`DOCX`, `ODT`, `PPTX`, `ODP`, `XLSX`, `ODS`, `CSV`, `TSV`, `PDF`, `RTF`, `EPUB`, `JSON`, `YAML`, `TOML`, `XML`, `XMIND`, `MMAP`.

It is intended to keep context organized instead of flattening everything into one large text blob. Headings, tables, images, charts, diagrams, sheets, and metadata remain available for downstream systems such as CognitiveOS V3.

It is not an email parser, chat parser, or store-ingest crate.

## Installation

Add to your `Cargo.toml`:

```toml
[dependencies]
office-parser = { git = "https://github.com/<user>/<repo>" }
```

To pick specific formats (disable defaults):

```toml
[dependencies]
office-parser = { git = "https://github.com/<user>/<repo>", default-features = false, features = ["docx", "pdf", "json"] }
```

## What It Extracts

- Ordered content blocks: headings, paragraphs, lists, tables, images, links
- Document metadata: format, title, page/slide counts, format-specific extras
- Embedded assets: extracted images with stable IDs and source references
- Semantic structures when detectable: spreadsheet segments, charts, diagrams
- Mind maps (`xmind`, `mmap`) as root heading + nested list hierarchy, with full tree preserved in `metadata.extra.mindmap`

## Why This Helps V3 Ingestion

V3 already consumes `office-parser` directly from bytes and derives:

- extracted text from normalized blocks
- section and hierarchy structure
- image hashes and metadata
- document-level metadata for retrieval and UI

The value of `office-parser` is not “avoid truncation” by itself. The value is preserving structure so downstream chunking and retrieval can select the right context with less loss of meaning.

## Minimal Usage

Parse by hint:

```rust
let bytes = std::fs::read("input.pptx")?;
let doc = office_parser::parse(&bytes, "input.pptx")?;
```

Render text or JSON:

```rust
let md = office_parser::render::to_markdown(&doc);
let chunks = office_parser::render::to_chunks(&doc, 2000);
let json = office_parser::render::to_json_value(&doc);
```

## CLI

A CLI wrapper is included in `./cli/` for inspection and export:

```bash
cd cli && cargo run -- document.docx --out ./output
```

See [cli/README.md](cli/README.md) for details.

## Docs

- [Library usage](docs/library-usage.md)
- [CLI usage](cli/README.md)
- [Formats and behavior](docs/formats.md)
- [Spreadsheets](docs/spreadsheets.md)
- [Charts and diagrams](docs/charts-diagrams.md)
- [Release guide](docs/release.md)

## MorphEditor Dialect (Markdown Output)

`--format markdown` emits MorphEditor-compatible markdown (the dialect
`BlockModel.js` parses) wrapped in OKF frontmatter (`type: document`,
optional `title`/`description`/`tags`, and for presentations `slides:
<ratio>` + optional `transition:`, validated against MorphEditor's
`SLIDE_RATIOS`/`SLIDE_TRANSITIONS`). Presentations (PPTX/ODP) become
decks: one `#` section per slide, `---` hr-delimited, speaker notes as
`::: note` fences inside their slide section. DOCX table column widths
become `::: table {widths=[..]}` wrappers and multi-column DOCX sections
become `::: columns` fences (even-volume column split — see Limitations).
Images emit MorphEditor media lines (`![alt](asset/…)`); inline
bold/italic/strikethrough/links serialize as `**`/`*`/`~~`/`[]()` marks
parsed from DOCX runs, ODT text spans, PPTX `rPr` attributes and ODP
styles — round-trip byte-exact through MorphEditor.

Mermaid diagram blocks are hardened: connector arrow directions are
honored (`<-->` for bidirectional, reversed connectors swap endpoints),
grouped shapes (`grpSp`) are flattened so every member becomes a node,
and ~30 OOXML preset geometries map to mermaid shapes (stadium, hex,
cyl, das, document, trap-b/t, lean-r, lin-rect, …).

Validation: `node tools/verify_morph_roundtrip.mjs <file.md>` (needs the
MorphEditor checkout) plus the Rust golden tests in `cli/tests/cli.rs`
and `tests/docx_marks.rs`.

### Limitations

- **PDF page images**: PDFs are parsed for text (lopdf) and embedded
  image objects — pages are **not rasterized**. Scanned/image-only PDFs
  are flagged via `pdf_text_quality: ImageOnly` so downstream pipelines
  can OCR them from the extracted embedded scans, but no composed page
  PNG is produced.
- **Presentations**: slide **transitions** are not auto-detected (an
  explicit `--slide-transition` flag sets the MorphEditor `transition:`
  frontmatter). **WordArt, 3-D, shadows, gradients, animations/timing**
  and freeform (`custGeom`) shape geometry are not rendered — diagram
  slides export as mermaid topology (what connects to what), not visual
  layout; decorative unconnected shapes are dropped.
- **Image width/position**: MorphEditor image lines carry no width attr
  (`![a](u){width=..}` fails image detection and parses as a paragraph) —
  sizes are dropped rather than emitted as corrupting markup.
- **DOCX flowed multi-column sections** (`w:cols num>=2`): emitted as
  `::: columns` + `::: column` fences. DOCX columns are flowed — the
  source carries no explicit split points — so column boundaries are an
  even-volume approximation of the print layout (round-trips cleanly;
  also captured in JSON metadata as `multi_column_sections`).
- **colspan/rowspan**: MorphEditor's table grammar (pipe tables +
  `widths`/`rowHeights`/`align` wrapper props only) cannot represent
  spans; cells are flattened.
- **DOCX headers/footers/TOC fields**: page chrome and duplicated TOC
  content are skipped.
- **`--chunk-size` output**: the CognitiveOS chunked view — NOT
  MorphEditor markdown.
