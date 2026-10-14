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

`--ocr` routes image-like PDF pages through the Indra OCR engine
(per-page routing: pages that fail the text-layer trust checks —
invisible OCR layers over scans, broken font maps — or fall under the
char threshold are OCR'd; text-rich pages keep their text layer):

```sh
office-parser-cli scan.pdf --out results/ --ocr
```

- Rust-only: the Indra pipeline runs **in-process** over ONNX Runtime
  (`--features ocr-indra-native`; models from `INDRA_OCR_MODEL_DIR` or
  the Hugging Face cache — no Python, no external OCR binary).
- A page renderer is required: `--features render-pdfium` (preferred —
  engine-faithful pixels) or `pdftoppm` on PATH.
- Output flows through the normal markdown/JSON render path.
  `metadata.extra["ocr"]` records provenance, per-page trust verdicts
  (`page_reasons`) and low-confidence pages. Table-class pages come
  back as real pipe tables; sideways scans are turned upright
  (orient.onnx) and deskewed; CJK lines route through rec_cjk.
- Tuning: `--ocr-min-chars N` (default 50), `--ocr-force` (OCR every
  page, discard text layers).

## Per-page output (`--split-pages`, `--page-markers`)

- `--split-pages` (PDF): writes `<stem>-p<N>.md` per page, each with
  its own OKF frontmatter and `tags: [pdf, page:<n>]`; the combined
  file is still written. With `--ocr`, every page is routed through the
  engine (no text-layer shortcuts).
- `--format json` on a PDF **always** also writes `<stem>.jsonl` — one
  document per line (`metadata.extra.page = N`) — for paginated
  ingestion, no flag needed.
- `--page-markers`: delimits the combined markdown with
  `<!-- page N -->` between pages.
- `--source <name-or-url>`: recorded in the OKF `source:` field of the
  combined and per-page documents.

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
