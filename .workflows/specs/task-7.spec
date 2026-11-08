spec: task
name: "complexity-score"
tags: [pdf, layout, quality, scoring]
verify: "cargo test --test quality_complexity && cargo test"
max-rounds: 2
---

## Intent

Implement the pure, deterministic `layout_complexity` function that turns TASK 4 geometry into a
`0.0..=1.0` score, using the calibration recommended by TASK 3p — including its deletion test, which
may reduce this task to a column count. No parser wiring here: this is a pure function plus its
invariants.

## Diagrams

```mermaid
flowchart LR
    A[PageGeometry] --> B[horizontal gutter count]
    A --> C[span overlap ratio]
    A --> D[table gutter density]
    A --> E[rotation != 0]
    A --> F[char density]
    B --> G[weighted sum]
    C --> G
    D --> G
    E --> G
    F --> G
    G --> H[clamp 0.0..=1.0]
```

## Decisions

- Input types come from `src/formats/pdf.rs` (TASK 4): `PageGeometry`, `SpanBox`. Import read-only;
  never modify that file.
- The function is **pure**: no I/O, no pdfium, no rendering, no globals.
- Output is clamped to `0.0..=1.0`; the clamp is part of the contract, not the caller's job.
- The signal set and weights are exactly those recommended by
  `.workflows/spikes/pdf-complexity/README.md`. If the spike's deletion test recommends dropping the
  composite score, implement `columns()` only, keep the same public name returning the column-derived
  score, and state the deletion in the spec's implementation note.
- The escalation threshold lives with the caller (TASK 8), **not** in this module — the scorer
  returns a number, it does not decide routing.
- The threshold constant carries a calibration doc comment naming the five fixtures it was measured
  on, and states that it is a heuristic, not a calibrated probability.
- No new dependency (the crate has no dev-dependencies; invariants are asserted example-based).

## Boundaries

### Allowed Changes
- `src/quality.rs`
- `tests/quality_complexity.rs` (new)

### Forbidden
- `src/formats/pdf.rs` — owned by TASK 6 (parallel group B); import only
- `src/document.rs`, `src/lib.rs`, `src/render.rs`
- Do not implement routing, OCR triggering, or verdict selection — that is TASK 8
- Do not add pdfium, a model, or any dependency
- Do not "tune" the score to make a fixture pass; record a disagreement in a code comment instead

## Completion Criteria

Scenario: score is always in range
  Test: test_complexity_score_within_bounds_for_all_fixtures
  Given all five fixtures in `tests/fixtures/`
  When every page's geometry is scored
  Then every score is `>= 0.0` and `<= 1.0`
  And no score is `NaN`

Scenario: adding a column gutter increases complexity (monotonicity)
  Test: test_added_column_gutter_increases_score
  Given a synthetic single-column `PageGeometry`
  When the same spans are split into two columns by moving a subset to the right by one column width
  Then the two-column score is strictly greater than the single-column score
  And both remain within bounds

Scenario: the spike's threshold decision is honoured
  Test: test_threshold_matches_calibrated_recommendation
  Given the threshold constant and the score table in
    `.workflows/spikes/pdf-complexity/README.md`
  When the fixture scores are compared against the constant
  Then every page the spike predicted as Escalate scores at or above the constant
  And every page it predicted as Trust scores below it

Scenario: scorer is deterministic and pure (negative)
  Test: test_scoring_same_input_twice_is_identical
  Given one `PageGeometry`
  When it is scored twice in a row, and once again after scoring a different page
  Then all three results are bit-identical
  And no file, environment variable, or global state affects the result
