# AGENTS.md — office-parser

## CORE (binding)

- Containment: work only within your task's Allowed Changes; anything else is reported, never fixed.
- Commits: workers never commit, merge, or push; a gate commits after review passes.
- Refactor first: extend existing structure before adding new surface (`extend` verdict this plan).
- TDD: write the failing test with the change, not after; keep the suite green.
- Subtract first: prefer deleting or simplifying over adding.
- Root causes: fix causes, not symptoms; never patch around a defect.
- Docs discipline: docs describe current state; history lands only in CHANGELOG.md.
- No model marks its own work done: a task's ✅ requires a reviewer verdict `ok:true`.
- Escalate, never absorb: a discovered contradiction with the plan's Refactor verdict stops the task.

## POINTERS

- Project overview: `README.md` — 17 office/data formats → one `Document` AST; MIT; consumers are
  CognitiveOS V3 and MorphEditor.
- Build and test: `cli/README.md` (CLI, all flags); `cargo test` (default features, no pdfium);
  `cargo test --features pdf,ocr-indra-native` (in-process Indra OCR).
- Architecture and docs index: `docs/formats.md` (per-format behaviour) · `docs/spreadsheets.md` ·
  `docs/charts-diagrams.md` · `docs/library-usage.md` · `docs/release.md`.
- Domain memory and plan: `.workflows/CONTEXT.md` (glossary, domain rules, **§Rulings**) ·
  `.workflows/plan.md` (the live plan) · `.workflows/proposals/` · `.workflows/specs/`.
- Dispatch policy: `agents/registry.md` — model, thinking, tiers, parallel groups, blocked states.

## PROJECT

- PDF is a feature-gated module (`pdf` = `dep:lopdf`); every `format` feature has a
  `#[cfg(not(feature = ...))]` stub returning `Error::FeatureDisabled`. Preserve that symmetry.
- `pdf_text_quality` is `None` for every non-PDF format — never leak format-specific metadata across
  formats (`tests/docx_odt_parity.rs` enforces the discipline).
- Geometry comes from the text layer, not the raster: quality work must never add a pdfium
  dependency, and `render-pdfium`-absent builds are first-class.
- `render-pages` defaults to pdfium then the vendored patched hayro; OCR needs
  `clamp_budget = false` (full-resolution source pixels). Never make rasterization optional-but-silent.
- Trust checks that degrade to an empty vector when a feature is off are a defect class, not a
  graceful default — assert, do not silently pass.
- `tests/fixtures/` is the only corpus: `pdf_scanned`, `pdf_mixed`, `pdf_table_scan`,
  `pdf_rot90_scan`, `pdf_cjk_scan`. Do not download fixtures.
- `Cargo.lock` and `roadmap.md`, `v3-ingestion.md`, `.bug-hunter/` are gitignored — never declare
  them in a Docs-impact list; prove dependency changes with `cargo tree`, not the lockfile.
- No dev-dependencies exist. Do not introduce a test framework without the human's sign-off.
- MorphEditor markdown output is a byte-exact round-trip contract; `cli/tests/cli.rs` and
  `tests/docx_marks.rs` are the locks.
