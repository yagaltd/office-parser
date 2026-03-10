# office-parser

[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Rust crate for parsing documents into a normalized `Document` AST that preserves useful hierarchy and semantic structure for downstream ingestion, chunking, retrieval, and inspection.

`office-parser` is for office-like documents and structured data files:
`DOCX`, `ODT`, `PPTX`, `ODP`, `XLSX`, `ODS`, `CSV`, `TSV`, `PDF`, `RTF`, `EPUB`, `JSON`, `YAML`, `TOML`, `XML`.

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

- [V3 ingestion notes](docs/v3-ingestion.md)
- [Library usage](docs/library-usage.md)
- [CLI usage](cli/README.md)
- [Formats and behavior](docs/formats.md)
- [Spreadsheets](docs/spreadsheets.md)
- [Charts and diagrams](docs/charts-diagrams.md)
- [Release guide](docs/release.md)
