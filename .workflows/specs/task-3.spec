spec: task
name: "pdf-complexity-spike"
tags: [prototype, pdf, layout, spike]
verify: "test -f .workflows/spikes/pdf-complexity/README.md && test -z \"$(git diff --name-only src/)\""
max-rounds: 2
---

## Intent

Answer one bounded question: **does a deterministic `layout_complexity` score built from span boxes
rank the five PDF fixtures the way a human would** — high for `pdf_mixed.pdf` (multi-column),
low for a single-column text page, and correctly ambiguous for the scans? The calibration note and
threshold recommendation produced here are the evidence the TASK 7 impl spec is drafted from. If the
score cannot beat `column count` alone, say so — that is a deletion, not a failure.

## Diagrams

```mermaid
flowchart LR
    A[span boxes] --> B[horizontal gutter count]
    A --> C[span overlap ratio]
    A --> D[table gutter density]
    A --> E[rotation != 0]
    A --> F[char density]
    B --> G[layout_complexity 0..1]
    C --> G
    D --> G
    E --> G
    F --> G
    G --> H{ranks fixtures like a human?}
    H -->|no| X[deletion test: keep columns only]
    H -->|yes| Y[recommend threshold]
```

## Decisions

- Signals are limited to what TASK 2p proved available from the text layer; no pdfium, no model.
- The score is clamped to `0.0..=1.0`.
- The escalation threshold is a **starting value only** and must carry the calibration note
  (which fixtures were measured); it is never presented as a calibrated probability.
- Evaluation set is the five existing fixtures — no new corpus, no downloads:
  `pdf_scanned.pdf`, `pdf_table_scan.pdf`, `pdf_rot90_scan.pdf`, `pdf_cjk_scan.pdf`, `pdf_mixed.pdf`.
- `pdf_rot90_scan.pdf` is expected to score high on the rotation component: a scan that is also
  rotated is genuinely the hardest case and the score should reflect that.
- Spike code is confined to `.workflows/spikes/pdf-complexity/**`.

## Boundaries

### Allowed Changes
- `.workflows/spikes/pdf-complexity/**`
- May **read** `.workflows/spikes/pdf-geometry/**` (consumes the TASK 2p box dump; never modify it)

### Forbidden
- `src/**`, `cli/**`, `tests/**`, `Cargo.toml`, `Cargo.lock`
- `.workflows/spikes/pdf-geometry/**` — read-only
- Do not tune the score against a fixture's expected label; record disagreement instead

## Completion Criteria

Scenario: all five fixtures scored
  Test: `.workflows/spikes/pdf-complexity/README.md` — section "Score table"
  Given the box dump from `.workflows/spikes/pdf-geometry/`
  When the spike scores every page of all five fixtures
  Then the README holds a table with one row per page, columns for fixture, page, each signal
    component, and the final score
  And `pdf_mixed.pdf` pages score higher than single-column pages from the same fixture set

Scenario: threshold recommendation is justified, not asserted
  Test: `.workflows/spikes/pdf-complexity/README.md` — section "Threshold"
  Given the score table
  When the spike proposes an escalation threshold
  Then the README places that threshold in the table, names every page it would route to Escalate,
    and justifies the cut with the measured numbers
  And the README states plainly that five fixtures is a hypothesis, not a calibration

Scenario: deletion test is answered
  Test: `.workflows/spikes/pdf-complexity/README.md` — section "Deletion test"
  Given the full score and the `column count` signal alone
  When the spike compares them on the five fixtures
  Then the README states whether the extra signals change any verdict
  And if they do not, the README recommends dropping the score in favour of `columns` alone

Scenario: production tree is untouched
  Test: `test -z "$(git diff --name-only src/)"`
  Given the spike ran to completion
  When `git diff --name-only src/` is executed in the repo root
  Then the output is empty
