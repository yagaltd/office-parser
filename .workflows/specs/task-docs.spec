spec: task
name: "docs-and-changelog"
tags: [docs, changelog]
verify: "grep -c '20261108-001' CHANGELOG.md && test -z \"$(git diff --name-only src/ tests/ cli/src/)\""
max-rounds: 2
---

## Intent

Update exactly the docs declared under the proposal's `## Docs impact`, and add the CHANGELOG entry
that names the Plan ID. The archive gate at SHIP reads that declaration: a declared path that was not
updated, or a CHANGELOG without the Plan ID, refuses the archive.

## Decisions

- The declaration is authoritative and was corrected on 2026-11-08: **`docs/v3-ingestion.md` is
  excluded** because `.gitignore` matches the `v3-ingestion.md` pattern, so it cannot be committed.
  Its intended content (how V3 consumes `Escalate`) moves into `docs/formats.md`. `roadmap.md` has
  the same trap and is also out of scope.
- `Escalate` is documented as a **vision-LLM escalation signal, never an OCR trigger** — the most
  likely misreading.
- The complexity threshold is documented as a **heuristic calibrated on five fixtures**, never as a
  calibrated probability.
- `--ocr-min-chars` default 50 is documented as unchanged; this plan adds `Escalate`, it does not
  re-tune `Ocr`.
- Docs describe current state only; history lands in `CHANGELOG.md` (repo docs discipline).
- No source, test, or CLI code changes in this task.

## Boundaries

### Allowed Changes
- `docs/formats.md`
- `docs/library-usage.md`
- `README.md`
- `cli/README.md`
- `CHANGELOG.md`

### Forbidden
- `src/**`, `tests/**`, `cli/src/**`
- `docs/v3-ingestion.md` and `roadmap.md` — gitignored, cannot ship
- Do not restructure unrelated sections of any doc
- Do not change the version number in `Cargo.toml` (release is a separate step)

## Completion Criteria

Scenario: every declared doc path is updated
  Test: `git diff --name-only` lists each of docs/formats.md, docs/library-usage.md, README.md, cli/README.md, CHANGELOG.md
  Given the repo state after this task
  When `git diff --name-only` runs
  Then all five declared paths appear
  And no path outside the Allowed list appears

Scenario: the CHANGELOG names the Plan ID
  Test: `grep -c '20261108-001' CHANGELOG.md`
  Given the updated CHANGELOG
  When grepped for the Plan ID
  Then the count is exactly 1
  And the entry sits under a new `0.3.0` heading

Scenario: the two likely misreadings are documented explicitly
  Test: `grep -q 'Escalate' docs/formats.md && grep -qi 'not an OCR trigger' docs/formats.md`
  Given the updated `docs/formats.md`
  When searched for the verdict semantics
  Then it states that `Escalate` routes to a vision model and is not an OCR trigger
  And it states that the complexity threshold is a heuristic measured on five fixtures

Scenario: no code changed (negative)
  Test: `test -z "$(git diff --name-only src/ tests/ cli/src/)"`
  Given the repo state after this task
  When source and test paths are diffed
  Then the output is empty
