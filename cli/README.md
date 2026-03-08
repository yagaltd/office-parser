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

# Config/data files
office-parser-cli config.yaml --out ./output
office-parser-cli data.toml --out ./output --format json
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
| `--chunk-size <n>` | - | Chunk size (chars) for generic text splitting |
| `--max-rows <n>` | - | Max rows per sheet (spreadsheets) |
| `--max-cols <n>` | - | Max columns per sheet (spreadsheets) |
| `--tail-rows <n>` | - | Include last N rows when truncating |
| `--table-rows <n>` | - | Max data rows per table segment |
| `--group-by <col>` | - | Split when this column changes |

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
- For format-specific behavior (charts, diagrams, spreadsheets, JSON/YAML/TOML, XML), see the [library README](../README.md).
