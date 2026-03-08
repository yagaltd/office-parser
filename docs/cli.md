# CLI Usage

## Run from repository root

```bash
cargo run -p office-parser-cli -- <input> --out <out_dir>
```

## Examples

```bash
# Markdown output (default)
office-parser-cli input.docx --out out

# JSON output
office-parser-cli input.pdf --out out --format json

# Spreadsheet tuning
office-parser-cli sheet.xlsx --out out --max-rows 20000 --tail-rows 5000
office-parser-cli sheet.xlsx --out out --group-by order_id
office-parser-cli sheet.xlsx --out out --group-by A
```

## Output layout

```text
<out>/
  <basename>.md or <basename>.json
  asset/
    <extracted files>
```

- Assets are written to `<out>/asset/`.
- Image references are rewritten to `asset/<file>`.
- Mermaid `office-image:sha256:...` placeholders are rewritten to actual asset paths.

## Chunking behavior

- Non-spreadsheets: use `--chunk-size` to emit chunk separators.
- Spreadsheets: default output is semantic table segmentation (no extra generic chunking unless `--chunk-size` is set).
