# V3 Ingestion Notes

`office-parser` is used by CognitiveOS V3 as a structured extraction layer for office-like documents.

The current flow is:

1. V3 receives raw document bytes.
2. `office-parser` parses the bytes into a normalized `Document`.
3. V3 derives extracted text, sections, hierarchy metadata, and image hashes from that `Document`.
4. V3 handles chunking, embedding, indexing, and storage.

This means the highest-value parser improvements are:

- better hierarchy preservation
- better extracted text quality
- better section boundaries
- better metadata
- better parity across equivalent source formats

## What V3 Actually Uses

From the parsed `Document`, V3 uses:

- normalized text from blocks
- hierarchy/sections from headings and block structure
- metadata such as title, page count, slide count, and format extras
- extracted images for hashing and downstream asset handling

The parser does not prevent truncation by itself. Its job is to preserve enough structure that V3 can retrieve the right slices of context instead of flattening the whole document poorly.

## What Belongs In `office-parser`

- deterministic parsing and extraction
- semantic document structure
- extracted assets and metadata
- export/inspection helpers such as Markdown and JSON rendering

## What Does Not Belong In `office-parser`

- store ingestion
- chunk/index/vector logic
- email parsing
- chat/conversation parsing
- fuzzy comparison or product workflow logic
