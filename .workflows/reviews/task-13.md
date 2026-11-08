# Verdict — TASK 13 (worker): own-table-structure (rusto-core removed)

ok:true

Reviewer: orchestrator (gate), 2026-11-08.

Delivery: the whole table-structure pipeline is ours —
`src/formats/table_structure.rs` (SLANet-plus via the rten runtime already
shipped by `ocr-rten`; embedded structure vocabulary; TableLabelDecode port
with end-token stop + `<td` bbox decode; TableMatch port (IoU + L1,
first-claim-wins, multi-det accumulation); own `<table>` rendering with
`html.escape` semantics). IndraNative exposes the geometry it already
computed (grid bbox → `layout_regions`, cell rects → `word_boxes`) under a
cfg-gated `ocr_page` override; recognition byte-identical. Deletions done:
`vendor/rusto-core`, `vendor/rusto-nn`, `src/formats/rusto_tables.rs`,
`tests/ocr_rusto_tables.rs`. Feature renamed `ocr-table-structure`
(= `pdf` + `image` + `ocr-rten`; zero new deps).

Gate corrections (recorded):
1. VOCABULARY (delivery bug): the worker embedded the EN 28-token dict; the
   shipped SLANet-plus.onnx itself carries the CH 48-token dict in its ONNX
   `character` metadata prop (read via the onnx module: 48 tokens → 50
   classes; outputs confirmed [?, ?, 8] loc + [?, ?, 50] probs). Gate
   replaced the vocab + indexes (END_IDX 49, head-dim match 50) from the
   model's own metadata — the binary now uses exactly what the model expects.
2. Test-expectation bugs (3): a 1×1 crop is upscaled to fill the whole
   488×488 canvas (no pad to assert — test now uses a 2×1 crop so the
   zero-pad path is real); the `<td` bbox loc row is step 0 (eos at step 1
   breaks before row 1 is read); the synthetic shape_list's ratio slots
   (x / ratio_w = 0.5, y / ratio_h = 0.25 per `_bbox_decode`).

Evidence:
- Unit: **14/14** (vocab incl. model-metadata parity, eos-stop, bbox decode,
  IoU/L1 math vs upstream, multi-det accumulation, escape semantics,
  preprocess normalization + zero-pad, no-model/engine-supplied/
  no-geometry/bogus-model no-ops).
- Integration: **5/5** — and 5/5 AGAIN with
  `OFFICE_PARSER_TABLE_MODEL=/tmp/slanet-plus.onnx` (SHA256 d57a942a…34514b
  = RapidAI v2.0.0 pin, verified): the REAL Indra engine + REAL SLANet-plus
  inference produced structure-correct `table_html` on pdf_table_scan
  (td open/close parity asserted; transcript-identity arms env-toggled).
- Lean-down proof: `! grep rusto-core Cargo.toml` PASS, no `vendor/rusto-*`,
  `cargo tree --no-default-features --features pdf -e normal` → 0 rusto.
- Default matrix (feature off): 98 passed / 0 failed; `tests/pdf_ocr.rs` +
  `tests/ocr_page_contract.rs` diffs empty; clippy 0 errors (2 pre-existing
  warnings outside this task's files).
