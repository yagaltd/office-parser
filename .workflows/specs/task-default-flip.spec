spec: task
name: "default-flip"
tags: [render, default, conditional]
verify: "cargo test --features pdf,ocr-indra-native,render-hayro,render-pdfium"
max-rounds: 3
---

## Intent

CONDITIONAL contract (finalized after TASK 4's human decision): execute the bake-off verdict.
Flip branch: hayro becomes the DEFAULT OCR renderer (pdfium the fallback), OFFICE_PARSER_RENDERER
still overrides both ways. Stay branch: pdfium stays default; the ruling + numbers land in
CONTEXT.md §Rulings. The task's final shape is confirmed against the TASK 4 decision before dispatch.

## Decisions

- Flip branch (only on the TASK 4 verdict): render-pages default order = hayro then pdfium;
  OFFICE_PARSER_RENDERER={hayro,pdfium} keeps explicit control; the default change is stated in
  --help and docs (the docs task owns final wording).
- Stay branch: no code change; CONTEXT.md §Rulings gains the ruling with the WER numbers and a
  revisit-when condition (e.g. upstream fixes knockout groups, or the WER gap closes).
- Either branch: no Indra/model changes; the OCR transcript corpus stays green as re-baselined
  in the upgrade task.

## Boundaries

### Allowed Changes
- Flip: `src/render.rs`, `src/formats/pdf.rs` (default order + render log strings),
  `cli/src/main.rs` (--help text), `README.md` (one paragraph — final wording owned by the docs task)
- Stay: `.workflows/CONTEXT.md` only

### Forbidden
- Removing the pdfium feature or the fallback path in either branch
- Changing quality/routing/table logic

## Acceptance Criteria

- Scenario: the decision is executed as recorded in TASK 4
  Test: the branch's behavior observable — flip: hayro used without OFFICE_PARSER_RENDERER set
  (render log names hayro); stay: pdfium still default
- Scenario: renderer override works both ways
  Test: OFFICE_PARSER_RENDERER=hayro and OFFICE_PARSER_RENDERER=pdfium each take effect
  (render log names the engine)
- Scenario: suite green on the union of render features
  Test: full pdf,ocr-indra-native,render-hayro,render-pdfium matrix passes
