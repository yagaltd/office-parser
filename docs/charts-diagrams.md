# Charts and Diagrams

## Charts

When cached chart data exists in the document:

- emit `### Chart`
- emit a note paragraph (`chart type`, optional `units`)
- emit a `Table` for categories and series values
- store chart metadata in `doc.metadata.extra["charts"]`

If charts reference external workbooks without cache, metadata may still be emitted but table data can be absent.

## Diagrams

For simple connector-style diagrams in PPTX/ODP/DOCX/ODT:

- emit `### Diagram`
- emit Mermaid `flowchart LR` block text
- store graph JSON in `doc.metadata.extra["diagram_graphs"]`

## Exclusions

- SmartArt and complex drawn visuals are not fully interpreted.
- Slide snapshots are not generated; the parser emits semantics, not rendered slide images.
