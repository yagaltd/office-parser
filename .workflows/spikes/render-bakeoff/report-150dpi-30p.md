# Render bake-off: pdfium vs hayro

- Generated: 2026-11-09 05:09:06 (UTC)
- CLI: /home/aurel/.cache/mbx/targets/v1/1bb8277a5e1b1755a9d97f56b631443cec7a4faac2d7ade3d4f919eb893ad9ce/release/office-parser-cli (sha256 ecf414fa0f76fe78, 55573848 bytes)
- Arms: OFFICE_PARSER_RENDERER=pdfium / hayro, interleaved per corpus item; wall time covers the full CLI process (startup + render + OCR + JSON write); a warm-up run preceded measurement.
- Corpus: seed 42, 30 pages, 55621 words - dense_scan.pdf sha256 21bc7c179192de4f, words.json sha256 a83f1a2eb2bd7202
- Env passthrough: OFFICE_PARSER_TABLE_MODEL=/tmp/slanet-plus.onnx - both are respected when set, never required.

## Arm integrity

| check | value |
|---|---|
| binary has pdfium markers (FPDF_/pdfium-render) | yes |
| binary has hayro markers | yes |
| pdfium lib (PDFIUM_PATH or ~/.cache/office-parser/pdfium/) | /home/aurel/.cache/office-parser/pdfium/libpdfium.so |
| synthetic ocr engine / pages_ocr (pdfium) | indra-native / 30 |
| synthetic ocr engine / pages_ocr (hayro) | indra-native / 30 |
| synthetic transcripts distinct across arms | yes |

## Synthetic corpus - dense_scan.pdf (30 pages, 55621 ground-truth words)

| metric | pdfium | hayro | Delta (hayro minus pdfium) |
|---|---:|---:|---:|
| WER - micro-avg (lower is better) | 0.6683 | 0.6964 | +0.0281 |
| WER - per page min..max | 0.566..0.857 | 0.595..0.858 | - |
| words | 24482 | 24153 | -329 |
| chars | 140319 | 138482 | -1837 |
| blocks | 123 | 135 | +12 |
| low-confidence pages | 0 | 0 | +0 |
| wall time (s) | 15.61 | 16.60 | +0.99 |

Decision rule (TASK 4): flip default to hayro iff WER(hayro) <= WER(pdfium) + 0.01 -> WER(pdfium)=0.6683, WER(hayro)=0.6964 -> rule **NOT satisfied** (margin -0.0181). The flip itself remains human-gated.

WER method: ground truth = words.json (exactly as drawn, column-major reading order); hypothesis = per-page OCR transcripts from the CLI JSONL; Levenshtein distance over word sequences / reference length, casefolded tokens, micro-averaged across pages.

## Fixtures - tests/fixtures/pdf_*.pdf (plain --ocr, routing respected; no WER column - no ground truth)

| fixture | metric | pdfium | hayro | Delta |
|---|---|---:|---:|---:|
| pdf_cjk_scan.pdf | words | 2 | 2 | +0 |
| pdf_cjk_scan.pdf | chars | 7 | 6 | -1 |
| pdf_cjk_scan.pdf | blocks | 2 | 1 | -1 |
| pdf_cjk_scan.pdf | low-confidence pages | 1 | 1 | +0 |
| pdf_cjk_scan.pdf | wall time (s) | 0.24 | 0.17 | -0.08 |
| pdf_mixed.pdf | words | 104 | 104 | +0 |
| pdf_mixed.pdf | chars | 723 | 723 | +0 |
| pdf_mixed.pdf | blocks | 8 | 8 | +0 |
| pdf_mixed.pdf | low-confidence pages | 0 | 0 | +0 |
| pdf_mixed.pdf | wall time (s) | 0.44 | 0.44 | +0.00 |
| pdf_rot90_scan.pdf | words | 52 | 52 | +0 |
| pdf_rot90_scan.pdf | chars | 361 | 361 | +0 |
| pdf_rot90_scan.pdf | blocks | 4 | 4 | +0 |
| pdf_rot90_scan.pdf | low-confidence pages | 0 | 0 | +0 |
| pdf_rot90_scan.pdf | wall time (s) | 0.42 | 0.42 | +0.00 |
| pdf_scanned.pdf | words | 76 | 76 | +0 |
| pdf_scanned.pdf | chars | 524 | 524 | +0 |
| pdf_scanned.pdf | blocks | 6 | 6 | +0 |
| pdf_scanned.pdf | low-confidence pages | 0 | 0 | +0 |
| pdf_scanned.pdf | wall time (s) | 0.66 | 0.63 | -0.02 |
| pdf_table_scan.pdf | words | 24 | 24 | +0 |
| pdf_table_scan.pdf | chars | 131 | 131 | +0 |
| pdf_table_scan.pdf | blocks | 3 | 3 | +0 |
| pdf_table_scan.pdf | low-confidence pages | 0 | 0 | +0 |
| pdf_table_scan.pdf | wall time (s) | 0.49 | 0.52 | +0.03 |

## real-scans/ - drop-in scans (char/word delta only; no WER column - no ground truth exists)

_(empty - drop .pdf files into the harness' real-scans/ directory and rerun run_bakeoff.py)_

## Notes

- Fixtures run with plain --ocr: page routing (text layer vs OCR) is text-layer based, hence renderer-independent - the A/B stays valid. --ocr-force is used only on the synthetic corpus and real scans.
- Renderer selection falls back silently (pdfium -> hayro -> pdftoppm, src/formats/pdf.rs); a runtime render failure of one backend can therefore still swap a pinned arm to pdftoppm. The transcript-distinctness check above catches the common both-arms-collapse case.
- WER(hayro) <= WER(pdfium) + 1pt is the TASK 4 flip rule; this report only measures - the flip is a separate human-gated task.
