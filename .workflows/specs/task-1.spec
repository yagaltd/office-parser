spec: task
name: "pdf-scout-interpreter-seams"
tags: [scout, pdf, baseline]
verify: "grep -q '## Open Questions' .workflows/CONTEXT.md && cargo test --features pdf"
max-rounds: 1
---

## Intent

Confirm, by reading `src/formats/pdf.rs`, exactly which interpreter state is available at the
point each text block is emitted, and produce a per-fixture baseline table: page count,
`pdf_text_quality`, per-page `page_char_counts`, whether `render-pdfium` changes the trust
verdict, and current reading-order output for `pdf_mixed.pdf` (the multi-column case).

## Acceptance Criteria

- Findings appended to `.workflows/CONTEXT.md` §Open Questions where they resolve or
  sharpen a question.
- Baseline table lists all five fixtures (`pdf_scanned`, `pdf_table_scan`, `pdf_rot90_scan`,
  `pdf_cjk_scan`, `pdf_mixed`) with non-empty rows.
- `cargo test --features pdf` still green (baseline is green).

## Constraints (Allowed Changes)

- `.workflows/CONTEXT.md`
- `.workflows/spikes/scout/**` (scratch notes)

Forbidden: everything else — especially `src/**`.
