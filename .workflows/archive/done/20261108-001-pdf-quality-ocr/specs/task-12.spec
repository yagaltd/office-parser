spec: task
name: "pdf-plan-docs"
tags: [docs, pdf]
verify: "grep -q '20261108-001' CHANGELOG.md && test -f docs/quality.md"
max-rounds: 1
---

## Intent

Document the shipped plan surface: per-page `PageTextQuality` + verdicts in metadata, XY-Cut
reading order, `OcrPage` contract, the no-ink trust check, complexity scoring, and the
`ocr-rusto-core` tables feature. Record the Plan ID in CHANGELOG.

## Acceptance Criteria

- README.md + docs/ updated for every user-visible surface above.
- CHANGELOG entry carries the Plan ID `20261108-001`.
- `scripts/check-archive.sh` passes (SHIP gate dry run).

## Constraints (Allowed Changes)

- `README.md`, `docs/**`, `CHANGELOG.md`

Forbidden: `src/**`.
