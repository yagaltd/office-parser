spec: task
name: "own-table-structure"
tags: [pdf, ocr, feature-flag, tables, rten]
verify: "cargo build && cargo test --features pdf,ocr-indra-native,ocr-table-structure --test table_structure && ! grep -q 'rusto-core' Cargo.toml && ! test -d vendor/rusto-core"
max-rounds: 3
---

## Intent

Own the table-structure pipeline end to end: port the PaddleOCR SLANet decode + TableMatch
into our code, delete the vendored rusto-core + rusto-nn, and run SLANet-plus through the
rten runtime we already ship. Human ruling 2026-11-08: "keeps the Rust core lean if
PaddleOCR logic ports" — the crate goes once the logic is ours.

### Diagrams

```mermaid
flowchart TD
    A[rendered page PNG] --> B[Indra det + rec + order]
    B --> C[blocks + confidence - unchanged]
    B --> G[grid bbox = layout_regions<br/>cell rects = word_boxes - NEW, geometry only]
    G --> H{ocr-table-structure on<br/>and slanet-plus.onnx present?}
    H -->|no| I[no-op - rich keys stay absent]
    H -->|yes| J[crop grid bbox, 488x488 preprocess]
    J --> K[rten session - SLANet-plus]
    K --> L[TableLabelDecode port<br/>tokens + cell bboxes]
    L --> M[TableMatch port<br/>IoU + L1 vs Indra cell rects]
    M --> N[own table rendering to table_html]
```

## Decisions

- Feature renamed `ocr-table-structure` (replaces `ocr-rusto-core` entirely). Deps: only
  `image` (already optional, PNG) + the existing optional `rten` (via `ocr-rten` wiring).
  No ndarray — plain `Vec<f32>` math.
- Model: SLANet-plus onnx (RapidAI v2.0.0, SHA256 d57a942af6a2f57d6a4a0372573c696a2379bf5
  857c45e2ac69993f3b334514b, verified at gate). Resolution: `OFFICE_PARSER_TABLE_MODEL`
  (exact file) > `models/slanet-plus.onnx` (CWD). Absent = documented no-op (one warning,
  rich keys stay absent) — the TASK 11 contract preserved.
- Structure vocabulary (`ppocr/utils/dict/table_structure_dict.txt`, 28 tokens) is EMBEDDED
  as a Rust const — no external file.
- IndraNative exposes the geometry it already computes: when `build_table_grid` emits a
  Table block it also records the grid bbox (`layout_regions`, label "Table") and cell
  rects (`word_boxes`, OcrWordBox { text, x0..y1 }). RECOGNITION UNTOUCHED — blocks and
  confidence byte-identical (locked by the existing suite). This amends task-11's
  indra_native.rs prohibition: the lock was on transcripts, not on geometry.
- Pipeline (with model + a table region present): crop region = 488x488 preprocess (exact
  ops from predict_structure.py / SLANet_plus.yml), then rten session, TableLabelDecode port
  (argmax tokens, td bbox decode x*(w/ratio_w)/(h/ratio_h)), TableMatch port (IoU + L1,
  Indra cell rects + texts mapped to SLANet cells), our own table rendering (structure rows,
  cols, spans; HTML-escaped text). Without the model: no-op contract (rich keys absent).
- Delete `vendor/rusto-core`, `vendor/rusto-nn`, the `rusto-core` dependency, and
  `src/formats/rusto_tables.rs`. `src/formats/mod.rs` registers the new module behind
  `ocr-table-structure`. Docs sweep within task: docs/quality.md + README rename the
  feature/model env (OFFICE_PARSER_TABLE_MODEL), CHANGELOG gains one subsection noting the
  rename + lean-down under the same Unreleased (history policy: current state only).

## Boundaries

### Allowed Changes
- `Cargo.toml`, `vendor/` (deletion), `src/formats/table_structure.rs` (new),
  `src/formats/rusto_tables.rs` (deleted), `src/formats/mod.rs`, `src/formats/pdf.rs`
  (cfg call site), `src/formats/indra_native.rs` (geometry exposure ONLY),
  `tests/table_structure.rs` (new), `docs/quality.md`, `README.md`, `CHANGELOG.md`

### Forbidden
- Any change to Indra's det/rec/order models, transcript content, or confidence values
- `tests/pdf_ocr.rs`, `tests/ocr_page_contract.rs`; `tests/ocr_rusto_tables.rs` is
  DELETED (its contract transfers to the new test file; its no-op/object-safety/
  transcript-identity scenarios must be re-covered)
- New dependencies beyond what `ocr-rten` already pulls

## Acceptance Criteria

- Scenario: the crate is gone — `! grep -q 'rusto-core' Cargo.toml`, `! test -d vendor/rusto-core`.
- Scenario: transcripts identical — the `pdf,ocr-indra-native` matrix is green and
  indra_native block/confidence output is unchanged (existing tests cover it).
- Scenario: no-model no-op — payload shape identical to feature-off.
- Scenario: real model — with OFFICE_PARSER_TABLE_MODEL set, pdf_table_scan produces
  table_html with structure-correct rows/cells (gate validates with the verified model).
- Scenario: feature off pulls nothing — `cargo tree --no-default-features --features pdf
  -e normal` shows no rusto-core; feature ON requires the `ocr-rten` feature in the graph.
