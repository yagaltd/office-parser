# office-parser roadmap

## Purpose

This roadmap is intentionally narrow.

The current priority for `office-parser` is to improve its usefulness for **real ingestion into V3**, not to turn the crate into a general product surface.

Today V3 already consumes `office-parser` directly from bytes and derives:

- extracted plain text from blocks
- sections from headings / paragraphs
- image hashes from extracted images
- document metadata / hierarchy

Reference:
- V3 adapter path: `/home/aurel/Documents/current/CognitiveOS/v3/crates/sdk/src/adapters.rs:523`

That means the most valuable work is the work that improves:

- extraction quality
- stable structure
- metadata quality
- text/chunk quality for retrieval
- image handling for ingestion

Not the work that mostly improves CLI cosmetics.

## Active Work

### 1. Improve parser output for V3 ingestion quality

This is the highest-value work.

Focus areas:

- Better document structure fidelity in the `Document` AST.
- Better section boundaries from headings, tables, and semantic blocks.
- Better metadata population in `doc.metadata.extra`.
- Better consistency across equivalent formats (`docx` vs `odt`, `pptx` vs `odp`, etc.).
- Better extracted text quality for chunking and retrieval.

Why:

- V3 currently uses the parser output directly, not a separate office canonical JSON contract.
- Better AST and text quality immediately improves ingestion results.

Recent improvements:

- `PdfTextQuality` detection — always-on, zero-dep classification of extracted text richness.
- `ExtractedImage.description` — field exists for external captioning (agent sets it via vision model).

### 2. Improve metadata that V3 can consume immediately

Prioritize metadata that is cheap to extract and useful in retrieval / UI / downstream domain logic.

Examples:

- author
- title
- page count / slide count
- chart metadata
- diagram metadata
- spreadsheet sheet names / table names
- stronger hierarchy information

Why:

- V3 already copies some of this into document metadata and hierarchy.
- Metadata improvements add value immediately without inventing a new contract.

### 3. Tighten tests around ingestion-relevant behavior

Tests should focus on what matters for V3 outcomes:

- stable extracted text
- stable section extraction
- stable hierarchy
- stable image extraction IDs / hashes
- parity across equivalent file formats
- spreadsheet segmentation behavior

This is more useful than expanding CLI-only snapshot behavior.

## Settled Decisions

### Parser crate and ingest concerns are separated

office-parser is extraction only. V3 adapter lives in its own crate at
`/home/aurel/Documents/current/CognitiveOS/v3/crates/sdk/src/adapters.rs`.
CLI is export/inspection only. No store-ingest logic will move into this crate.

### V3 file ingest is generic, parsers are specialized

V3 has separate adapter paths:

- `DocumentAdapter` — office-parser (office/pdf/epub-like files)
- `TextAdapter` — plain `.txt` / `.md`
- `EmailAdapter` — mailbox-parser (email-specific path)
- Chat — separate path for conversation/turn models

office-parser stays focused on office/document formats.

### JSON is optional debug output, not a core architecture bet

`render::to_json()` exists for debugging and manual inspection. V3 does not consume JSON.
`JsonRenderOptions` (with `include_image_bytes: false`) keeps JSON light when needed.

## Stale References

If the roadmap references an old path under `CognitiveOS/parsers/office-parser/`,
that crate has moved to `/home/aurel/Documents/github/office-parser/`.
