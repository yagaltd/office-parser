# Library Usage

`office-parser` exposes a normalized `Document` AST for downstream systems that need more than plain text. Use it when you need hierarchy-preserving extraction for ingestion, retrieval, inspection, or export.

## Parse by hint

```rust
let bytes = std::fs::read("report.docx")?;
let doc = office_parser::parse(&bytes, "report.docx")?;
```

## Parse by explicit format

```rust
let bytes = std::fs::read("slides.pptx")?;
let doc = office_parser::parse_as(&bytes, office_parser::Format::Pptx)?;
```

The parser is deterministic and side-effect free. It returns a `Document` containing ordered blocks, extracted images, and metadata.

## Spreadsheet options

```rust
let bytes = std::fs::read("data.xlsx")?;
let opts = office_parser::spreadsheet::ParseOptions {
    max_rows_per_sheet: 20_000,
    tail_rows: 2_000,
    max_table_rows_per_segment: 500,
    group_by: Some("order_id".to_string()),
    ..Default::default()
};
let doc = office_parser::xlsx::parse_with_options(&bytes, opts)?;
```

## Presentation parse options

`pptx::ParseOptions` and `odp::ParseOptions` currently expose `include_slide_snapshots`, but snapshot generation is reserved and not yet emitted. Parsing still extracts semantic slide content, embedded images, charts, and diagram metadata.

## Render outputs

```rust
let md = office_parser::render::to_markdown(&doc);
let chunks = office_parser::render::to_chunks(&doc, 2000);
let json = office_parser::render::to_json_value(&doc);
let light_json = office_parser::render::to_json_value_with_options(
    &doc,
    office_parser::render::JsonRenderOptions {
        include_image_bytes: false,
    },
);
```

Use:

- Markdown when you want a readable text projection
- chunk rendering when you want generic bounded text segments
- JSON when you want structured inspection/debug output

The lightweight JSON option is useful when you want image metadata without embedding base64 image bytes into the output.

## Core model

- `Document`:
  - `blocks`: ordered content blocks (`Heading`, `Paragraph`, `List`, `Table`, `Image`, `Link`)
  - `images`: extracted binary assets and metadata
  - `metadata`: format/title/page/slide counts + format-specific `extra`

## Why Structure Matters

The main value of `office-parser` is preserving enough structure for downstream retrieval to select the right context:

- headings define sections
- tables stay tables instead of being flattened into noisy prose
- sheets/slides/pages remain recoverable through metadata and structure
- charts and diagrams surface semantic content when possible
