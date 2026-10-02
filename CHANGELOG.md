## Unreleased

### MorphEditor dialect (`--format markdown`)

- OKF frontmatter on every export: `type: document`, optional
  `title`/`description`/`tags`; presentations additionally emit
  `slides: <ratio>` (from slide size) and support `--slide-transition`.
- Deck mode for PPTX/ODP: one `#` section per slide, `---` hr-delimited
  (MorphEditor hr-delimited deck sections); speaker notes from
  `notesSlide` (PPTX) / `presentation:notes` (ODP) emitted as
  `::: note` fences inside their slide section.
- Inline marks across DOCX/ODT/PPTX/ODP: bold/italic/strikethrough from
  run properties, hyperlinks as link marks; ODT automatic-style parent
  chains resolved. Serialized as `**`/`*`/`` ` ``/`~~`/`[]()`.
- DOCX tables emit `::: table {widths=[..]}` wrappers from `gridCol`.
- Multi-column DOCX sections (`w:cols num>=2`) emit `::: columns` +
  `::: column` fences with an even-volume column split; ranges also
  captured in metadata (`multi_column_sections`).
- Images emit MorphEditor media lines (`![alt](asset/…)`) instead of the
  debug form.
- MorphEditor sigil escaping for paragraph lines; mermaid fences pass
  through as code-fence blocks.

### Mermaid diagram hardening

- Connector arrow directions honored: bidirectional `<-->`, reversed
  connectors swap endpoints; standalone line arrows normalized.
- Grouped shapes (`grpSp`) flattened recursively so every member becomes
  a diagram node.
- OOXML preset-geometry map expanded (~30 presets: terminator, document,
  delay, manual op, storage, magnetic media, preparation, connectors…).

### Docs & validation

- New round-trip tooling: `tools/verify_morph_roundtrip.mjs` and
  `tools/golden_roundtrip.mjs` (validates emitted markdown through
  MorphEditor's own BlockModel).
- New fixtures + integration tests: run marks (DOCX/ODT/PPTX/ODP),
  speaker notes, column sections, table widths.

# Changelog

## 0.2.0

- **Fixed:** Panic on malformed RTF input when a HYPERLINK/INCLUDEPICTURE field with an empty or data-image URL occurs at a group nesting depth that empties the internal `group_stack`. The replenishment check now runs before field processing, preventing a subsequent `unwrap()` panic on an empty stack.

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
