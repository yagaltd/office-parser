# office-parser

[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Standalone Rust crate for extracting a structured document AST (`Document` + `Block`s + extracted images) from:
`DOCX`, `ODT`, `PPTX`, `ODP`, `PDF`, `RTF`, `EPUB`, and config/data files: `JSON`, `YAML`, `TOML`, `XML`.

## Features

Formats are feature-gated:

- `docx`, `odt`, `pptx`, `odp`, `pdf`, `rtf`, `epub`, `json`, `yaml`, `toml`, `xml`, `xlsx`, `ods_sheet`, `csv`, `tsv` (enabled by default)

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

## Library usage

### Parse (auto-detect)

```rust
let bytes = std::fs::read("input.pptx")?;
let doc = office_parser::parse(&bytes, "input.pptx")?;
```

### Parse (explicit format)

```rust
let bytes = std::fs::read("report.pdf")?;
let doc = office_parser::parse_as(&bytes, office_parser::Format::Pdf)?;
```

### Render outputs

```rust
let md = office_parser::render::to_markdown(&doc);
let chunks = office_parser::render::to_chunks(&doc, 2000);
let json = office_parser::render::to_json_value(&doc);
```

## Spreadsheets (XLSX/ODS/CSV/TSV)

Supported formats:
- XLSX (Office Open XML)
- ODS (ODF)
- CSV / TSV

Key behaviors:
- Spreadsheets are rendered as **semantic tables** (multiple `Table` blocks as needed), with repeated `#### Rows ...` sub-headings.
- Row/column limits are enforced to avoid pathological memory/output.

Important options (see `office_parser::spreadsheet::ParseOptions`):
- `max_rows_per_sheet`, `tail_rows`: head+tail truncation for large sheets.
- `max_cols_per_sheet`, `max_cell_chars`.
- `max_table_rows_per_segment`: max data rows per emitted table segment.
- `group_by`: split segments when a key column changes (header name or Excel column letter like `A`, `AA`).

Cell formatting:
- XLSX: best-effort **style-aware** date/time formatting (reads `xl/styles.xml` + per-cell style index).
- Fallback heuristic: if styles are missing, headers containing `date|time|timestamp|datetime|when` may trigger formatting of Excel serials.

## Charts

For PPTX/DOCX (and some ODT embedded objects), if the file includes **cached chart values** (`<c:strCache>` / `<c:numCache>` or ODF table backing data), we emit:

- A `### Chart` heading
- A `Paragraph` note like `Note: bar chart; units: %`
- A `Table` block with the chart’s categories + series values

If a chart references an external workbook and the cache is missing, you will still get chart metadata (when detectable), but may not get a `Table`.

Chart metadata is also stored in `doc.metadata.extra["charts"]`.

## Diagrams

For PPTX/ODP/DOCX/ODT, when we detect a simple “diagram” (text boxes + connector arrows/lines), we emit:

- `### Diagram`
- A Mermaid `flowchart LR` code block (as a `Paragraph` block)

Notes/limits:

- We intentionally skip SmartArt / hand-drawn / complex visuals.
- Connector text labels (when present) are included on Mermaid edges.
- Slide snapshot images are currently not generated. `pptx::ParseOptions::include_slide_snapshots` and `odp::ParseOptions::include_slide_snapshots` are reserved for future support.

Diagram graph JSON is also stored in `doc.metadata.extra["diagram_graphs"]`.

## JSON / YAML / TOML

`office-parser` parses JSON/YAML/TOML into a Markdown-KV-like block structure (useful for ingestion + chunking):

- Nested objects become headings up to **depth 4** (`#`..`####`)
- Leaf scalar fields become `key: value` lines
- Arrays:
  - array of scalars → `List`
  - array of flat objects (scalar fields) → `Table` (first column `#` is the item index)
  - otherwise → fenced ```jsonl``` payload

Key order is preserved (source order), rather than sorted.

## XML

Generic XML parsing (e.g. WordPress WXR exports) is supported:

- Elements become headings/keys; attributes become `@attr` KV lines
- Repeated tags become arrays (flat arrays-of-objects become `Table`; otherwise repeated sections)
- CDATA that looks like HTML is converted to Markdown best-effort

## CLI

A CLI wrapper is included in `./cli/` for command-line usage:

```bash
cd cli && cargo run -- document.docx --out ./output
```

See [cli/README.md](cli/README.md) for details.

For extended docs, see `./docs/`.
