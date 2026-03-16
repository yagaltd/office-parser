# Changelog

## Unreleased

- Added input parsing support for mind map formats: `xmind` and `mmap`.
- Added new `Format` variants and hint detection for `.xmind`, `.mmap`, and `application/vnd.xmind.workbook`.
- Added new parser entry points: `office_parser::xmind::parse` and `office_parser::mmap::parse`.
- Added Cargo features `xmind` and `mmap` (enabled by default).
- Mind map parsing maps content into the normalized `Document` AST as root heading + nested list items.
- Added full mind map tree metadata under `metadata.extra.mindmap` (`root`, `node_count`, `max_depth`, and `sheet_index` for XMind).
- Improved XMind XML parsing to preserve entity-referenced punctuation and spacing in topic titles.
- Added parser/unit/integration coverage for XMind JSON/XML, MMAP XML, and format hint detection.

## 0.1.0

- Initial public release.
- Supports Office formats (`docx`, `odt`, `pptx`, `odp`), spreadsheets (`xlsx`, `ods`, `csv`, `tsv`), `pdf`, `rtf`, `epub`, and config/data formats (`json`, `yaml`, `toml`, `xml`).
- Exposes normalized `Document` AST and renderers for Markdown/chunks/JSON.
