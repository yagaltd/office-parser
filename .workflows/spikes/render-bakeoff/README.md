# render-bakeoff — pdfium vs hayro measurement harness (TASK 3)

Measurement-only spike: it produces the numbers for the TASK 4 human decision
(default-renderer flip). It changes nothing in src/, tests/ or Cargo.toml and
is not wired into CI or the test suite (contract boundary:
.workflows/spikes/render-bakeoff/** only).

## Layout

    generate_corpus.py   seeded dense-scan generator: reportlab text layer
                         (A4, 2-column, 8pt Helvetica) -> /tmp pages ->
                         pdftoppm -r 150 -png -> image-only dense_scan.pdf
                         + exact ground truth words.json
    wer.py               WER scorer (Levenshtein / reference length)
                         + --selfcheck
    run_bakeoff.py       A/B runner (pdfium vs hayro via
                         OFFICE_PARSER_RENDERER) over the synthetic corpus,
                         tests/fixtures/pdf_*.pdf and real-scans/ -> report.md
                         + --selfcheck
    real-scans/          drop-in slot for real scans (char/word delta only,
                         never WER)
    report.md            written by the runner; input to TASK 4

## Requirements

- python3 — stdlib ONLY for wer.py and run_bakeoff.py. Exception: the
  generator imports the system-wide reportlab package (the one allowed
  non-stdlib dependency, generator only).
- pdftoppm (poppler-utils) — generator rasterization; also the CLI's
  last-resort render backend.
- office-parser-cli built with the render + OCR features (from the repo
  root; produces target/release/office-parser-cli, the default --cli):

    CARGO_BUILD_JOBS=4 cargo build --release --manifest-path cli/Cargo.toml \
        --features render-pdfium,render-hayro,ocr-indra-native \
        --target-dir target

- Indra OCR models resolvable (INDRA_OCR_MODEL_DIR or the standard
  Hugging Face cache, datalab-to/indra-ocr). OFFICE_PARSER_TABLE_MODEL is
  respected when set, never required.

## Usage (generate -> run -> report)

    python3 .workflows/spikes/render-bakeoff/generate_corpus.py
    python3 .workflows/spikes/render-bakeoff/run_bakeoff.py

- Generator defaults: seed 42, 10 pages, A4 2-column 8pt Helvetica,
  150 dpi image-only output. --seed / --pages / --out-dir override.
- The runner auto-generates the corpus when dense_scan.pdf / words.json are
  missing (--no-generate forbids that). Same seed -> byte-identical corpus
  and words.json, so WER is exactly reproducible.

Self-checks (both exit 0 on pass):

    python3 .workflows/spikes/render-bakeoff/wer.py --selfcheck
    python3 .workflows/spikes/render-bakeoff/run_bakeoff.py --selfcheck

run_bakeoff.py --selfcheck verifies (1) generator determinism — two seeded
runs produce identical dense_scan.pdf + words.json sha256 checksums — and
(2) the WER scorer micro-case. Headline value:

    ref ["a","b","c"] vs hyp ["a","x","c"]  ->  WER = 1/3 = 0.333333

## Ground truth and WER

words.json records the per-page word list EXACTLY as drawn (draw order =
column-major: left column top-to-bottom, then right column). WER = standard
Levenshtein edit distance over the page word sequence, normalized by the
reference length; tokens are whitespace-split and casefolded; micro-averaged
across pages (total edit distance / total reference words). Hypothesis =
per-page OCR transcripts extracted from the CLI's JSON output
(--format json --json-no-image-bytes --ocr --ocr-force --split-pages; the
JSONL carries one document per page via metadata.extra.page).

## Decision rule (TASK 4)

Flip the default renderer to hayro iff WER(hayro) <= WER(pdfium) + 1pt
(0.010) on the bake-off corpus. report.md states the rule verdict directly
under the synthetic table. The flip itself is human-gated (TASK 4) and
executed by TASK 5 — this harness only measures.

## real-scans drop-in

    cp my-scan.pdf .workflows/spikes/render-bakeoff/real-scans/
    python3 .workflows/spikes/render-bakeoff/run_bakeoff.py

Any PDF there joins the A/B as its own rows, scored with char/word delta
ONLY (no WER column — no ground truth exists for real scans). Runs use
--ocr --ocr-force (real scans are image-like by definition). The runner
succeeds with the directory empty (section renders as a placeholder note).

## Arm integrity (why the runner is picky)

Renderer selection in src/formats/pdf.rs falls back SILENTLY
(pdfium -> hayro -> pdftoppm). A binary built without render-pdfium, or
run without libpdfium.so installed, would silently serve the "pdfium" arm
with pdftoppm — a fake bake-off (AGENTS.md: never optional-but-silent).
The runner therefore:

1. sniffs the CLI binary for feature-gated markers (hayro crate paths;
   pdfium-render FPDF_* libloading symbols);
2. mirrors pdfium_lib_path() (PDFIUM_PATH env, else
   ~/.cache/office-parser/pdfium/libpdfium.so);
3. hard-fails before measuring when a pinned backend could not have served
   its arm (exit 3 with the rebuild/install instruction);
4. warns loudly in report.md if both arms produce byte-identical synthetic
   transcripts (same-backend collapse). Residual caveat: a runtime render
   failure of exactly one backend can still swap that arm to pdftoppm —
   compare the per-arm ocr engine/pages_ocr rows and the distinctness
   check when in doubt.

## Notes

- Wall time = full CLI process (startup + render + OCR + JSON write),
  measured per corpus item, arms interleaved; one discarded warm-up run
  precedes measurement (OS page cache + ONNX session init).
- Heavy artifacts (CLI JSON outputs, per-page JSONL) live in
  /tmp/render-bakeoff (--work-dir); only the corpus and report live here.
- Fixture runs use plain --ocr so per-page routing (text layer vs OCR) is
  respected; routing is text-layer based and therefore renderer-independent,
  which keeps the A/B valid.
