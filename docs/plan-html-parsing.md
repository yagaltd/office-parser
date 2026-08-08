# Plan: HTML format support in office-parser

Status: planning. Owner: yagaltd/office-parser. Consumed by CognitiveOS v3
(coding-agent-loop-v3.md, WS1.5 — `web_fetch` tool: webpage = document).

## Why

v3's agent loop ingests documents from bytes (docx, pdf, epub, ...) into a
normalized `Document` AST, then renders to markdown/chunks for the store.
A webpage is a document like any other: fetch a URL once, parse it with the
same pipeline, persist a `research/reference` node, and never re-fetch (the
store is the cache). HTML is the only missing format for that vision.

## Current state (verified 2026-08-07)

- `src/formats/epub.rs:292` — `parse_xhtml_to_blocks(xhtml_bytes, xhtml_path)`
  already converts XHTML into AST blocks (headings, paragraphs, images with
  `resolve_path`, links). This is the seed: HTML support is extracting and
  generalizing it, not new ground.
- `src/formats/mod.rs` — format registry + hint detection (by extension /
  content sniffing) that dispatches bytes to the right parser.
- `src/document_ast.rs` — the normalized AST (`Document`, blocks, metadata,
  embedded assets).
- `src/render.rs` — `to_markdown`, `to_chunks`, `to_json_value` (format-agnostic
  once the AST is built).

## Goal

Add `html` (and `htm`) to the supported format list:

- Full HTML documents (not just EPUB-scoped XHTML): `<!DOCTYPE html>` pages,
  malformed / real-world HTML5, relative URLs for images/links.
- Same AST contract as every other format: ordered blocks (headings,
  paragraphs, lists, tables, links, images), metadata (`format: "html"`,
  title from `<title>`), embedded image assets.
- No HTTP, no JS rendering, no SSRF concerns inside the crate — input is
  bytes; transport is the caller's job (v3's `web_fetch` tool).

## Approach

1. **Extract, don't duplicate.** Move `parse_xhtml_to_blocks` out of
   `epub.rs` into a new `src/formats/html.rs`, generalized:
   - Accept a full HTML document, not only EPUB chapter fragments.
   - Parse with an HTML5-tolerant parser. EPUB XHTML is well-formed XML; real
     web HTML is not. Evaluate `html5ever` (spec-compliant, tolerant) against
     the current XML-based approach — if the existing parser handles
     well-formed XHTML only, HTML5 tolerance is a hard requirement for the
     web case. Decision point (see Open questions 1).
   - Extract `<title>` → metadata; keep the block mapping
     (h1-h6 → heading, p → paragraph, ul/ol → list, table → table,
     img → image asset, a[href] → link with resolved URL).
   - Resolve relative URLs against a base URL. `epub.rs` has
     `resolve_path` for in-EPUB paths; html.rs needs the general form
     (base-URL resolution) — caller supplies the base (v3 passes the
     fetched URL).
   - Strip by default: `<script>`, `<style>`, `<nav>`, `<noscript>`,
     `<iframe>`, hidden elements — same spirit as the EPUB path.
2. **Wire the registry.** `src/formats/mod.rs`: hint `html` / `htm` by
   extension + content sniff (`<!DOCTYPE html`, `<html`, `text/html`
   content-type hint via the caller's mime, when provided).
3. **Epub keeps working.** `epub.rs` delegates to the shared html.rs module
   for its chapter XHTML (same behavior, one parser). EPUB tests must stay
   green unchanged.
4. **Metadata + render.** `format: "html"`, title, language if present.
   No render.rs changes (format-agnostic already).

## Non-goals

- No web fetching / transport (v3 `web_fetch` owns HTTP + SSRF guard).
- No JS-rendered pages (no headless browser; static HTML only — matches the
  deterministic-first posture).
- No sanitization policy beyond the strip list above — raw content is the
  caller's decision (v3 stores content in `research/reference` nodes).

## Tests (no mocks — real fixture files)

- `tests/fixtures/html/` — real-world-shaped pages:
  - article page (headings, paragraphs, links, one image)
  - table-heavy page (data table → table blocks)
  - minimal page (`<title>` only)
  - malformed HTML (unclosed tags, missing doctype)
- Unit (in `html.rs`): block mapping per fixture; title metadata; link
  extraction; relative-URL resolution against a base; strip list honored;
  hint detection for `html`/`htm`.
- Regression: existing EPUB tests untouched and green (shared module).
- e2e: `parse(bytes, "page.html")` → `render::to_markdown` → structural
  asserts (headings present, script content absent).

## Docs + release

- README format list: add HTML.
- `docs/formats.md`: new HTML section (block mapping, strip list, base-URL
  handling).
- CHANGELOG: `v0.3.0 — feat: html format` (feature, new format).
- v3 ingestion doc (`docs/v3-ingestion.md`): note web_fetch consumes it.

## Commits (office-parser repo)

1. `refactor(epub): extract XHTML-to-blocks into shared html module` — pure
   move, EPUB tests green.
2. `feat(html): full-document HTML parsing (title, links, base-URL, strip
   list)` — html.rs generalized + registry wiring.
3. `test(html): fixture-based coverage (article/table/minimal/malformed)`.
4. `docs: README + formats.md + CHANGELOG for html format`.

## Open questions

1. Parser choice: extend the existing XML-based XHTML parser vs adopt
   `html5ever` for real-world HTML5 tolerance. Bias: html5ever — the web
   case demands tolerance the XML path cannot give; verify epub regression
   cost before committing.
2. Base-URL resolution: caller-supplied base only, or sniff `<base href>`?
   (draft: caller-supplied wins, `<base href>` honored when present).
3. Strip list: is `<nav>` always noise, or context for some documents?
   (draft: strip by default; revisit with real v3 retrieval data.)
