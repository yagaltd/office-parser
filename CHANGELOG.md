## Unreleased

### PDF OCR — per-page routing for mixed PDFs

- `--ocr` no longer requires the whole document to classify as
  `ImageOnly`. Page text is now counted per page: pages under the
  50-character threshold are routed through Indra while text-rich pages
  keep their text layer — digital compilations with scanned inserts
  (visa bundles, signed contracts with attached scans) recover every
  page in order. Driven by the existing per-page block ranges; both
  engines (native and subprocess) route through the same
  `PageOcrEngine` seam.
- Fixed a temp-dir race when two renders ran in one process (parallel
  tests): render/OCR scratch dirs now carry a sequence number.
- New fixture `tests/fixtures/pdf_mixed.pdf` (digital page + rasterized
  page) with a routing test asserting both sources survive.

### PDF OCR — native Indra engine (Mode B, feature `ocr-indra-native`)

- The full Indra OCR pipeline ported to Rust over ONNX Runtime
  (`ort`, pinned rc.13): line detection (connected-component core
  decoding + adaptive second pass), CTC recognition (graph-decoded
  ids/conf/lengths), reading order and block classes (OrderNet heads:
  average-linkage same-block clustering, XY-cut / top-down inner
  orders, insertion-sweep pairwise ordering), and paragraph assembly
  with hyphen-joining. No Python, no subprocess.
- Models resolve from `INDRA_OCR_MODEL_DIR` or the standard Hugging
  Face cache; vocab comes from the engine's own `meta.json`.
- `--ocr` now prefers the native engine when built with the feature
  and models are present, falling back to the subprocess engine.
- V1 omissions vs the Python engine (degrade gracefully): orientation
  correction, deskew, CJK/script recognizers, math, CenterMask
  windows, table cell splitting, digital-text-layer passthrough.
- Validated against the Python engine's own JSONL output on the
  scanned-fixture corpus (same headings, same paragraph text).

### PDF OCR — Indra engine (Mode A, feature `ocr-indra`)

- Scanned PDFs (`pdf_text_quality: ImageOnly`) can now be recovered
  end-to-end inside the parser: render pages → run `indra-ocr`
  (subprocess; `INDRA_OCR_BIN` override) → map layout JSON to blocks
  (Title/Section-Header → headings, Text → paragraphs, consecutive
  List-items merged, page furniture and pictures dropped, unknown labels
  degrade to paragraphs).
- Per-page mean line confidence recorded in
  `metadata.extra["ocr"]` with `low_confidence_pages` for downstream
  vision-LLM fallback routing.
- New CLI flag `--ocr` (PDF only). Zero new Rust dependencies — the
  gate exists so library consumers can exclude the OCR code path.
- Real-fixture tests: a genuine image-only scanned PDF
  (`tests/fixtures/pdf_scanned.pdf` — rasterized text pages re-wrapped)
  plus the engine's actual JSONL output as mapping golden. The
  end-to-end test auto-skips when no indra binary is present.
- Note: Indra model weights are OpenRAIL-M licensed (free for research/
  personal/startups under $20M; not usable competitively with Datalab's
  API) — review before commercial shipping.
- Fix pre-existing clippy deny lint in xlsx number-format detection
  (logically redundant disjunct removed; behavior identical).

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
