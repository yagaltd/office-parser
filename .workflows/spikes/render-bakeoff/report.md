# Render bake-off: pdfium vs hayro

- Generated: 2026-11-09 04:56:39 (UTC)
- CLI: /home/aurel/.cache/mbx/targets/v1/1bb8277a5e1b1755a9d97f56b631443cec7a4faac2d7ade3d4f919eb893ad9ce/release/office-parser-cli (sha256 ecf414fa0f76fe78, 55573848 bytes)
- Arms: OFFICE_PARSER_RENDERER=pdfium / hayro, interleaved per corpus item; wall time covers the full CLI process (startup + render + OCR + JSON write); a warm-up run preceded measurement.
- Corpus: seed 42, 10 pages, 18538 words - dense_scan.pdf sha256 09a7d5866c4e84c4, words.json sha256 62e16f6d13927a26
- Env passthrough: OFFICE_PARSER_TABLE_MODEL=/tmp/slanet-plus.onnx - both are respected when set, never required.

## Arm integrity

| check | value |
|---|---|
| binary has pdfium markers (FPDF_/pdfium-render) | yes |
| binary has hayro markers | yes |
| pdfium lib (PDFIUM_PATH or ~/.cache/office-parser/pdfium/) | /home/aurel/.cache/office-parser/pdfium/libpdfium.so |
| synthetic ocr engine / pages_ocr (pdfium) | indra-native / 10 |
| synthetic ocr engine / pages_ocr (hayro) | indra-native / 10 |
| synthetic transcripts distinct across arms | yes |

## Synthetic corpus - dense_scan.pdf (10 pages, 18538 ground-truth words)

| metric | pdfium | hayro | Delta (hayro minus pdfium) |
|---|---:|---:|---:|
| WER - micro-avg (lower is better) | 0.6760 | 0.7121 | +0.0361 |
| WER - per page min..max | 0.594..0.812 | 0.617..0.832 | - |
| words | 8126 | 8086 | -40 |
| chars | 46517 | 46272 | -245 |
| blocks | 39 | 46 | +7 |
| low-confidence pages | 0 | 0 | +0 |
| wall time (s) | 4.95 | 5.32 | +0.38 |

Decision rule (TASK 4): flip default to hayro iff WER(hayro) <= WER(pdfium) + 0.01 -> WER(pdfium)=0.6760, WER(hayro)=0.7121 -> rule **NOT satisfied** (margin -0.0261). The flip itself remains human-gated.

WER method: ground truth = words.json (exactly as drawn, column-major reading order); hypothesis = per-page OCR transcripts from the CLI JSONL; Levenshtein distance over word sequences / reference length, casefolded tokens, micro-averaged across pages.

## Fixtures - tests/fixtures/pdf_*.pdf (plain --ocr, routing respected; no WER column - no ground truth)

| fixture | metric | pdfium | hayro | Delta |
|---|---|---:|---:|---:|
| pdf_cjk_scan.pdf | words | 2 | 2 | +0 |
| pdf_cjk_scan.pdf | chars | 7 | 6 | -1 |
| pdf_cjk_scan.pdf | blocks | 2 | 1 | -1 |
| pdf_cjk_scan.pdf | low-confidence pages | 1 | 1 | +0 |
| pdf_cjk_scan.pdf | wall time (s) | 0.21 | 0.18 | -0.03 |
| pdf_mixed.pdf | words | 104 | 104 | +0 |
| pdf_mixed.pdf | chars | 723 | 723 | +0 |
| pdf_mixed.pdf | blocks | 8 | 8 | +0 |
| pdf_mixed.pdf | low-confidence pages | 0 | 0 | +0 |
| pdf_mixed.pdf | wall time (s) | 0.43 | 0.45 | +0.02 |
| pdf_rot90_scan.pdf | words | 52 | 52 | +0 |
| pdf_rot90_scan.pdf | chars | 361 | 361 | +0 |
| pdf_rot90_scan.pdf | blocks | 4 | 4 | +0 |
| pdf_rot90_scan.pdf | low-confidence pages | 0 | 0 | +0 |
| pdf_rot90_scan.pdf | wall time (s) | 0.41 | 0.42 | +0.01 |
| pdf_scanned.pdf | words | 76 | 76 | +0 |
| pdf_scanned.pdf | chars | 524 | 524 | +0 |
| pdf_scanned.pdf | blocks | 6 | 6 | +0 |
| pdf_scanned.pdf | low-confidence pages | 0 | 0 | +0 |
| pdf_scanned.pdf | wall time (s) | 0.62 | 0.61 | -0.01 |
| pdf_table_scan.pdf | words | 24 | 24 | +0 |
| pdf_table_scan.pdf | chars | 131 | 131 | +0 |
| pdf_table_scan.pdf | blocks | 3 | 3 | +0 |
| pdf_table_scan.pdf | low-confidence pages | 0 | 0 | +0 |
| pdf_table_scan.pdf | wall time (s) | 0.52 | 0.51 | -0.01 |

## real-scans/ - drop-in scans (char/word delta only; no WER column - no ground truth exists)

_(empty - drop .pdf files into the harness' real-scans/ directory and rerun run_bakeoff.py)_

## Notes

- Fixtures run with plain --ocr: page routing (text layer vs OCR) is text-layer based, hence renderer-independent - the A/B stays valid. --ocr-force is used only on the synthetic corpus and real scans.
- Renderer selection falls back silently (pdfium -> hayro -> pdftoppm, src/formats/pdf.rs); a runtime render failure of one backend can therefore still swap a pinned arm to pdftoppm. The transcript-distinctness check above catches the common both-arms-collapse case.
- WER(hayro) <= WER(pdfium) + 1pt is the TASK 4 flip rule; this report only measures - the flip is a separate human-gated task.
