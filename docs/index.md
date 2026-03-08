# office-parser

`office-parser` parses Office and data/config files into a normalized document AST for ingestion pipelines (including CognitiveOS) and export workflows.

## Crates

- `office-parser`: library crate for parsing and rendering.
- `office-parser-cli`: CLI wrapper that writes Markdown/JSON plus extracted assets.

## Supported inputs

- Office docs: `docx`, `odt`, `pptx`, `odp`
- Spreadsheets: `xlsx`, `ods`, `csv`, `tsv`
- Other docs: `pdf`, `rtf`, `epub`
- Config/data: `json`, `yaml`, `toml`, `xml`

## Outputs

- Structured AST (`Document` with `Block`s)
- Markdown (`render::to_markdown`)
- Chunked markdown text (`render::to_chunks`)
- JSON (`render::to_json_value`)

## Design goals

- Deterministic extraction for ingestion and indexing.
- Feature-gated parsers per format.
- Stable behavior for large spreadsheets via semantic segmentation.
- Semantic chart/diagram extraction when the source contains usable structured data.
