# office-parser-cli

CLI wrapper around the `office-parser` crate.

## Installation

```bash
git clone https://github.com/<user>/<repo>
cd <repo>/cli
cargo install --path .
```

## Usage

```bash
office-parser-cli <input> [options]
```

### Basic examples

```bash
# Parse to Markdown (default)
office-parser-cli document.docx --out ./output

# Parse to JSON
office-parser-cli slides.pptx --out ./output --format json

# Parse to lightweight JSON (omit embedded base64 image bytes)
office-parser-cli slides.pptx --out ./output --format json --json-no-image-bytes

# Config/data files
office-parser-cli config.yaml --out ./output
office-parser-cli data.toml --out ./output --format json

# Mindmap files
office-parser-cli roadmap.xmind --out ./output --format json
office-parser-cli weekly.mmap --out ./output
```

### Spreadsheet options

```bash
# Limit rows (head+tail truncation)
office-parser-cli large.xlsx --out ./output --max-rows 10000 --tail-rows 500

# Split by key column
office-parser-cli orders.xlsx --out ./output --group-by order_id
office-parser-cli sheet.xlsx --out ./output --group-by A  # Excel column letter
```

## Options

| Flag | Default | Description |
|------|---------|-------------|
| `--out <dir>` | `.` | Output directory |
| `--format <fmt>` | `markdown` | Output format: `markdown` or `json` |
| `--json-no-image-bytes` | off | JSON output: omit embedded base64 image bytes |
| `--chunk-size <n>` | - | Chunk size (chars) for generic text splitting |
| `--max-rows <n>` | - | Max rows per sheet (spreadsheets) |
| `--max-cols <n>` | - | Max columns per sheet (spreadsheets) |
| `--tail-rows <n>` | - | Include last N rows when truncating |
| `--table-rows <n>` | - | Max data rows per table segment |
| `--group-by <col>` | - | Split when this column changes |
| `--ocr` | off | PDF: route image-like pages through the in-process Indra OCR engine (Rust + ONNX Runtime; models from `INDRA_OCR_MODEL_DIR` or the HF cache) |
| `--ocr-min-chars <n>` | `50` | PDF + `--ocr`: OCR pages with fewer than N alphanumeric chars |
| `--ocr-force` | off | PDF + `--ocr`: OCR every page, discard text layers |
| `--split-pages` | off | PDF: also write one document per page (`<stem>-p<N>.md`; JSON also emits `<stem>.jsonl`). With `--ocr`, every page goes through the engine |
| `--page-markers` | off | PDF + `--split-pages`: delimit the combined markdown with `<!-- page N -->` |
| `--source <name-or-url>` | - | Document name/URL recorded in the OKF `source:` frontmatter field |
| `--slide-transition <t>` | - | MorphEditor deck mode: slide transition (`up`, `down`, `left`, `right`, `fade`) |

### PDF options (scans, pagination)

```bash
# Scanned PDF: OCR image-like pages (text-layer pages keep their text)
office-parser-cli scan.pdf --out ./output --ocr

# Mixed bundle: every page through the engine, one document per page,
# page-delimited combined markdown, provenance in frontmatter
office-parser-cli bundle.pdf --out ./output --split-pages --ocr --page-markers \
  --source "Bundle display name or URL"

# JSON consumers get a paginated JSONL by default (one document per line)
office-parser-cli bundle.pdf --out ./output --format json
#   bundle.json       combined document
#   bundle.jsonl      one page per line (metadata.extra.page = N)
```

- Auto routing per page: pages failing the text-layer trust checks
  (invisible OCR layers, broken font maps) or under
  `--ocr-min-chars` go to OCR; text-rich pages keep their text layer.
  `--ocr-force` overrides the checks entirely.
- Engine: in-process ONNX Runtime (`--features ocr-indra-native`);
  rendering via pdfium (`--features render-pdfium`, preferred) or
  `pdftoppm`. Sideways scans are turned upright and deskewed;
  CJK lines route through the CJK recognizer.
- Provenance lands in `metadata.extra["ocr"]`: engine, per-page trust
  verdicts, low-confidence pages (the vision-LLM escalation signal).

## Output layout

Given input `slides.pptx` with `--out ./output`:

```
output/
  slides.md           # or slides.json
  asset/
    image_0001.png
    image_0002.jpg
    ...
```

Extracted images are written to `<out>/asset/` and referenced as `asset/<file>` in the output.

## Notes

- Format detection is based on input filename extension.
- The CLI is for inspection/export, not store ingestion.
- `xmind`/`mmap` are parsed from the first root map/sheet and rendered via the normal markdown/json outputs.
- For format-specific behavior (charts, diagrams, spreadsheets, JSON/YAML/TOML, XML, PDF
  limitations) see the [library README](../README.md) and [docs/](../docs/).
