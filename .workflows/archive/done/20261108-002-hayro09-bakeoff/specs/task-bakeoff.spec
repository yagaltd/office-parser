spec: task
name: "bakeoff"
tags: [render, benchmark, wer, measurement]
verify: "test -f .workflows/spikes/render-bakeoff/report.md && test -d .workflows/spikes/render-bakeoff/real-scans && python3 .workflows/spikes/render-bakeoff/run_bakeoff.py --selfcheck"
max-rounds: 3
---

## Intent

Measurement-only bake-off harness (human decision 4b): a seeded dense-scan generator with EXACT
ground truth (WER), a pdfium-vs-hayro A/B runner over synthetic + fixtures + a `real-scans/`
drop-in slot, and a markdown report. Lives under `.workflows/spikes/render-bakeoff/` — no src/
changes, no CI wiring. Its numbers feed TASK 4's human flip decision (rule: WER(hayro) <=
WER(pdfium) + 1pt).

## Decisions

- Determinism: the generator is seeded (fixed default seed); two runs produce byte-identical
  corpora and identical words.json — WER is exactly reproducible.
- Ground truth: words.json (per-page ordered word list the generator drew) is the WER reference.
  WER = standard edit distance over the page transcript word sequence, normalized by reference length.
- Metrics per renderer per corpus item: WER (synthetic only), char delta, word delta, block count,
  low-confidence page count, wall time. Report = one markdown table per corpus section.
- `real-scans/` drop-in: any .pdf placed there joins the run, scored with char/word delta ONLY
  (no WER claim — no ground truth). The runner must not fail when the dir is empty.
- Renderers: the runner uses the office-parser CLI built with render-pdfium + render-hayro and
  switches via OFFICE_PARSER_RENDERER; CARGO_BUILD_JOBS=4; heavy artifacts stay in the harness
  dir or /tmp.
- Implementation: python3 stdlib (generator/scorer/runner glue) + the CLI binary; no new crates,
  no new runtime dependencies.

## Boundaries

### Allowed Changes
- `.workflows/spikes/render-bakeoff/**` (new: generator, scorer, runner, README, real-scans/)

### Forbidden
- Anything outside `.workflows/spikes/render-bakeoff/` (in particular src/**, tests/**, Cargo.toml)
- Downloading fixtures (AGENTS.md: corpus is generated or human-provided only)
- Wiring the harness into CI or the test suite

## Acceptance Criteria

- Scenario: deterministic generation
  Test: two consecutive generator runs produce identical corpora (checksum match) and identical words.json
- Scenario: end-to-end A/B run
  Test: runner executes pdfium and hayro arms over the synthetic corpus + the 5 fixtures and
  emits report.md with per-corpus metric tables for BOTH renderers
- Scenario: WER scorer sanity
  Test: scorer on a hand-built micro case (known reference vs transcript) returns the hand-computed WER
- Scenario: real-scans drop-in
  Test: with at least one PDF in real-scans/, the runner scores it (char/word delta row, no WER
  column) and the run succeeds; with the dir empty, the run still succeeds
