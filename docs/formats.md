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
- `epub`: spine XHTML and embedded images.
- `xmind`, `mmap`: mindmap topic trees mapped to heading + nested list blocks, with full tree in `metadata.extra.mindmap`.

## Config/data formats

- `json`, `yaml`, `toml`: converted to heading/KV/list/table/jsonl-style block structure.
- `xml`: generic XML mapping to heading/KV/list/table blocks; attributes as `@attr`; CDATA HTML converted to markdown best-effort.

## Known limits

- Slide snapshot image generation is not currently implemented.
- Chart tables depend on cached values embedded by source apps.
- Complex visuals (SmartArt, hand-drawn diagrams) are intentionally not fully interpreted.
