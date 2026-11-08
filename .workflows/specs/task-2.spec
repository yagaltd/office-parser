spec: task
name: "pdf-geometry-spike"
tags: [prototype, pdf, geometry, spike]
verify: "test -f .workflows/spikes/pdf-geometry/README.md && test -z \"$(git diff --name-only src/)\""
max-rounds: 2
---

## Intent

Answer one bounded question before any build task is drafted: **does the existing content
interpreter in `src/formats/pdf.rs` already hold enough state to emit per-span `[x0,y0,x1,y1]`
boxes with `font_name`/`font_size`, and does glyph-area ÷ page-area separate a scanned page from a
born-digital one without `render-pdfium`?** A GO turns Track 1 of the proposal into an `extend`;
a NO-GO re-tiers TASK 4 as `refactor-first` and stops the plan. Throwaway code only — production
paths are forbidden by construction.

### Diagrams

```mermaid
flowchart LR
    A[content stream op] --> B{state.tm / font_size available at emit point?}
    B -->|no| N[NO-GO: refactor-first]
    B -->|yes| C[derive per-span bbox]
    C --> D[ink = glyph area / page area]
    D --> E{separates scan from digital?}
    E -->|no| N2[NO-GO for ink coverage only]
    E -->|yes| G[GO + recorded numbers]
```

## Decisions

- Spike code lives **only** under `.workflows/spikes/pdf-geometry/**`; it is never merged.
- The candidate interpreter state is fixed by the scout: `state.tm`, `state.font_size`
  (`src/formats/pdf.rs:925`, `:1005`, `:1024`, `:1059`), `effective_font_size()` (`:291`),
  `estimate_text_advance()` (`:471`). The spike may not introduce a new text path to find them.
- Fixtures are fixed: `tests/fixtures/pdf_scanned.pdf` (image-only, the scan case) and
  `tests/fixtures/pdf_mixed.pdf` (born-digital, the digital case).
- `render-pdfium` must not be enabled for the measurement — the claim under test is
  "without pdfium".
- Boxes are in PDF points; ink coverage is glyph box area ÷ page area, page from MediaBox.
- Every number in the findings must be reproducible by one command recorded in the README.

## Boundaries

### Allowed Changes
- `.workflows/spikes/pdf-geometry/**`

### Forbidden
- `src/**` — no production code, not even a helper
- `cli/**`, `tests/**`, `Cargo.toml`, `Cargo.lock`
- Do not modify or "fix" the interpreter; report what it already provides
- Do not enable `render-pdfium`, `render-hayro`, or any OCR feature

## Acceptance Criteria

Scenario: per-span geometry is derivable from existing interpreter state
  Test: `.workflows/spikes/pdf-geometry/README.md` — section "Box dump" with a go/no-go on the seam
  Given the spike reproduces a page of `tests/fixtures/pdf_mixed.pdf` through the existing interpreter
  When the spike dumps every emitted span as `x0 y0 x1 y1 font_name font_size "text"`
  Then the README shows the first 20 spans for at least one page with all seven fields populated
  And the README names the exact interpreter lines the values come from (`src/formats/pdf.rs:NNN`)

Scenario: ink coverage separates scan from born-digital
  Test: `.workflows/spikes/pdf-geometry/README.md` — section "Ink coverage table"
  Given the spike computes glyph box area ÷ page area per page
  When it runs over `pdf_scanned.pdf` and `pdf_mixed.pdf` without `render-pdfium`
  Then the README holds a per-page table with numeric coverage for both fixtures
  And the scan's coverage is numerically different from the digital fixture's, with the two ranges
    and the separating threshold stated explicitly
  And a one-line GO or NO-GO on "ink coverage identifies a scan without pdfium"

Scenario: NO-GO is recorded, not worked around
  Test: `.workflows/spikes/pdf-geometry/README.md` — section "Verdict"
  Given either question above fails
  When the spike owner writes the verdict
  Then the README states NO-GO, names which of the two questions failed, and lists the refactor that
    would be required
  And the README does not propose a new text path inside the spike

Scenario: production tree is untouched
  Test: `test -z "$(git diff --name-only src/)"`
  Given the spike ran to completion
  When `git diff --name-only src/` is executed in the repo root
  Then the output is empty
