# CLI Usage

The canonical CLI reference lives in [cli/README.md](../cli/README.md).

Use that page for:

- installation
- examples
- flags
- output layout
- JSON export options

## MorphEditor markdown output

`--format markdown` emits MorphEditor-compatible markdown wrapped in OKF
frontmatter. Presentation extras:

- `slides: <ratio>` is derived from the slide size (16:9 / 4:3).
- `--slide-transition <up|down|left|right|fade>` sets the
  `transition:` frontmatter key (validated against MorphEditor's
  `SLIDE_TRANSITIONS`).
- Speaker notes, `::: columns`, `::: table {widths=[..]}` and image
  media lines are emitted where the source carries them.

Note: `--chunk-size` output is the CognitiveOS chunked view and is NOT
MorphEditor markdown.

## PDF OCR (`--ocr`)

`--ocr` runs the Indra OCR engine over scanned PDFs (only when the
document classifies as `ImageOnly`):

```sh
office-parser-cli scan.pdf --out results/ --ocr
```

- Two engines: built with `--features ocr-indra-native`, the Indra
  pipeline runs in-process over ONNX Runtime (models from
  `INDRA_OCR_MODEL_DIR` or the Hugging Face cache — no Python).
  Otherwise it shells out to `pip install indra-ocr`.
- A page renderer is required either way: build with
  `--features render-pdfium` or have `pdftoppm` on PATH.
- Override the binary with `INDRA_OCR_BIN` (e.g. a venv install).
- Output flows through the normal markdown/JSON render path; OCR
  provenance and low-confidence pages are recorded in
  `metadata.extra["ocr"]`. Table-class pages come back as real pipe
  tables (native engine).
- Tuning: `--ocr-min-chars N` routes pages with fewer than N
  alphanumeric chars to OCR (default 50 — lower for thin/garbage text
  layers); `--ocr-force` OCRs every page and discards text layers.
