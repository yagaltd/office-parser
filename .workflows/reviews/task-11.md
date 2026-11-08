# Escalation — TASK 11 (rusto-core tables): STOPPED, human review required

Status: ❌ ESCALATED (spec's own stop-clause triggered)

2026-11-08, orchestrator. **`rusto-core` does not exist** — crates.io API:
`crate 'rusto-core' does not exist`. The task-11 spec's API surface
(LayoutDetector, LayoutType::Table) references a crate that was never published.
The spec's own escalation clause fires ("If rusto-core cannot be vendored … stop
and escalate"). Nothing was implemented; no dependency added; `src/` untouched.

The only existing "rusto" crate is `rusto-rs` 0.3.2 (PaddleOCR-based OCR) —
**explicitly forbidden** by the spec ("Do not add rusto-rs … Never the rusto-rs
facade crate").

Human options at review:
(a) DROP TASK 11 — tables stay on the existing paths (born-digital PDF tables via
    the text-layer table heuristic; OcrPage.table_html seam is shipped and ready
    for a future engine). No functionality lost that any fixture exercises.
(b) RE-POINT to an in-house span-geometry table detector (no new crate): we
    already ship xy_cut + count_columns + gutter machinery; a table detector over
    aligned-cut clustering (the TASK 3 spike's aligned_cuts signal) can fill
    OcrPage.table_html under the same off-by-default feature. Honors every
    constraint (transcript-identical, tables only, no rusto-rs).
(c) RE-POINT to a real published crate with table-structure output, if the human
    names one (license must clear the AGENTS.md AGPL ruling).
