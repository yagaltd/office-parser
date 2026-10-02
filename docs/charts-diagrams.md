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

Mermaid hardening details:

- **Arrow directions**: connector `headEnd`/`tailEnd` arrowheads are
  read; bidirectional connectors emit `<-->`, reversed connectors swap
  endpoints so the arrowhead lands on the original source shape.
- **Grouped shapes**: `grpSp` members are flattened recursively — every
  grouped shape becomes a node and can participate in edges.
- **Preset geometry**: ~30 OOXML presets map to mermaid shapes
  (`flowChartTerminator` → `stadium`, `flowChartDocument` → `document`,
  `flowChartManualOperation` → `trap-b`, storage/magnetic → `das`/`cyl`,
  decision → `diamond`, …). Unknown presets fall back to `rect`.

## Exclusions

- SmartArt and complex drawn visuals are not fully interpreted.
- Slide snapshots are not generated; the parser emits semantics, not rendered slide images.
- Diagram slides export topology (what connects to what), not visual
  layout: positions, z-order and decorative unconnected shapes are not
  represented in the mermaid block.
