//! rusto-core table-region enrichment for OCR'd pages (feature `ocr-rusto-core`).
//!
//! Contract: `.workflows/specs/task-11.spec` (+ human amendment: `rusto-core`
//! ships as a pinned git+rev dependency of the public rusto-rs repo — the
//! original "no rusto crates" line predates the local-checkout discovery and
//! is superseded; the `rusto-rs` facade crate itself stays forbidden).
//!
//! Scope is deliberately narrow: for pages OCR'd by ANY [`PageOcrEngine`],
//! this module fills the optional [`OcrPage`] rich fields:
//! - `layout_regions`: `LayoutType::Table` regions only (rusto-core
//!   `LayoutDetector`, raster-pixel coords, label "Table" — mirrors Indra's
//!   block label),
//! - `table_html`: rusto-core `TableStructure::to_html()` markup for the
//!   page's `Block::Table` grids, gated on a detected table region,
//! - `word_boxes`: NEVER populated — Indra's seam exposes no per-word data
//!   (tokens live inside indra_native.rs, which this task may not touch) and
//!   rusto-core's detector yields no text; documented decision.
//!
//! Transcript safety: enrichment runs AFTER the engine produced
//! blocks + confidence and only fills fields that are still `None`;
//! `blocks`/`confidence` are moved through untouched, so transcripts are
//! byte-identical with the feature on or off. The same det/rec/order models
//! run in both cases — no Indra re-tuning anywhere.
//!
//! Model resolution: `OFFICE_PARSER_LAYOUT_MODEL` (exact file path) wins,
//! else `models/DocOCR/layout.onnx` relative to the process CWD (rusto's
//! example convention). Without a model file the enrichment is a no-op:
//! every rich field stays `None` — the feature-off payload shape — and the
//! tests skip loudly (quality_routing.rs pdftoppm pattern). The .onnx loads
//! through rusto-nn's pure-Rust rten/rten-onnx path (no ort, no pdfium).

use std::path::{Path, PathBuf};

use super::pdf::{OcrLayoutRegion, OcrPage};
use crate::document_ast::{Block, Cell};

/// Default DocOCR layout model path (rusto example convention), relative to
/// the process working directory.
pub const DEFAULT_LAYOUT_MODEL: &str = "models/DocOCR/layout.onnx";

/// Environment override: exact path to the layout model file
/// (`.onnx` via rusto-nn's rten-onnx importer; `.rten` also accepted).
pub const LAYOUT_MODEL_ENV: &str = "OFFICE_PARSER_LAYOUT_MODEL";

/// Model resolution rule (documented contract): [`LAYOUT_MODEL_ENV`] wins
/// when it names an existing file, else [`DEFAULT_LAYOUT_MODEL`]. `None`
/// means enrichment is a no-op and every rich [`OcrPage`] field stays
/// `None` — the feature-off payload shape.
pub fn layout_model_path() -> Option<PathBuf> {
    if let Some(dir) = std::env::var_os(LAYOUT_MODEL_ENV) {
        let p = PathBuf::from(dir);
        return p.is_file().then_some(p);
    }
    let p = PathBuf::from(DEFAULT_LAYOUT_MODEL);
    p.is_file().then_some(p)
}

/// Enrich one OCR'd page. Called from `parse_with_page_ocr_inner` after
/// `engine.ocr_page(png)`; `page_png` is the same rendered PNG the engine
/// received, so region coords share the [`OcrWordBox`] raster space.
pub(crate) fn enrich_ocr_page(page: OcrPage, page_png: &[u8]) -> OcrPage {
    let model = layout_model_path();
    enrich_with_model(page, page_png, model.as_deref())
}

/// Testable core: `model_path = None` models the no-model environment
/// deterministically (no env access).
fn enrich_with_model(page: OcrPage, page_png: &[u8], model_path: Option<&Path>) -> OcrPage {
    // Opt-in seam: an engine that already supplied rich fields wins.
    if page.layout_regions.is_some() || page.word_boxes.is_some() || page.table_html.is_some() {
        return page;
    }
    let Some(model_path) = model_path else {
        warn_layout_model_missing();
        return page;
    };
    let regions = match detect_table_regions(page_png, model_path) {
        Ok(regions) => regions,
        Err(e) => {
            eprintln!("office-parser: rusto-core layout enrichment failed: {e}");
            return page;
        }
    };
    let has_table_region = !regions.is_empty();
    // Table markup: rusto confirms table layout; the cell content is Indra's
    // own `Block::Table` grid (indra_native.rs build_table_grid) — no
    // re-recognition, transcript untouched. Multiple table blocks render as
    // consecutive `<table>` elements (blocks carry no geometry at this seam,
    // so per-region assignment is not possible; the region acts as the
    // layout verdict).
    let table_html = tables_html(&page.blocks, has_table_region);
    let layout_regions = has_table_region.then_some(regions);
    OcrPage {
        layout_regions,
        table_html,
        ..page // blocks, confidence, word_boxes pass through untouched
    }
}

/// Run rusto-core's `LayoutDetector` over the rendered page and keep
/// `LayoutType::Table` regions only.
fn detect_table_regions(
    page_png: &[u8],
    model_path: &Path,
) -> Result<Vec<OcrLayoutRegion>, String> {
    let img = image::load_from_memory(page_png).map_err(|e| format!("page png decode: {e}"))?;
    let rgb = img.to_rgb8();
    let (w, h) = (rgb.width(), rgb.height());
    let mat = rusto_core::image::Mat::from_rgb8(w, h, rgb.into_raw())
        .map_err(|e| format!("page mat: {e}"))?;
    let config = rusto_core::LayoutConfig::default(model_path.to_path_buf());
    let mut detector = rusto_core::engine::LayoutDetector::new(config)
        .map_err(|e| format!("layout model load ({}): {e}", model_path.display()))?;
    let output = detector
        .detect(&mat)
        .map_err(|e| format!("layout detect: {e}"))?;
    Ok(table_regions_to_layout(&output.regions))
}

/// `LayoutType::Table` regions only — the feature's whole scope.
fn table_regions_to_layout(regions: &[rusto_core::LayoutRegion]) -> Vec<OcrLayoutRegion> {
    regions
        .iter()
        .filter(|r| r.layout_type == rusto_core::LayoutType::Table)
        .map(|r| OcrLayoutRegion {
            // mirrors Indra's "Table" block label (document_ast parity)
            label: "Table".to_string(),
            x0: r.bbox.x_min as f32,
            y0: r.bbox.y_min as f32,
            x1: r.bbox.x_max as f32,
            y1: r.bbox.y_max as f32,
        })
        .collect()
}

/// Render the page's `Block::Table` grids through rusto-core's
/// `TableStructure::to_html()` when a table region was detected.
fn tables_html(blocks: &[Block], has_table_region: bool) -> Option<String> {
    if !has_table_region {
        return None;
    }
    let mut html = String::new();
    for block in blocks {
        if let Block::Table { rows, .. } = block {
            if !html.is_empty() {
                html.push('\n');
            }
            html.push_str(&table_structure_html(rows));
        }
    }
    (!html.is_empty()).then_some(html)
}

/// AST rows → rusto-core `TableStructure` → HTML. Cell bbox is unused by
/// `to_html`; spans flow through (`colspan`/`rowspan` attributes), text is
/// HTML-escaped by rusto-core.
fn table_structure_html(rows: &[Vec<Cell>]) -> String {
    let num_cols = rows.iter().map(Vec::len).max().unwrap_or(0);
    let rusto_rows: Vec<rusto_core::TableRow> = rows
        .iter()
        .enumerate()
        .map(|(r, cells)| rusto_core::TableRow {
            cells: cells
                .iter()
                .enumerate()
                .map(|(c, cell)| rusto_core::TableCell {
                    bbox: rusto_core::BBox::new(0, 0, 0, 0),
                    row: r,
                    col: c,
                    row_span: cell.rowspan.max(1),
                    col_span: cell.colspan.max(1),
                    text: cell.text.clone(),
                })
                .collect(),
        })
        .collect();
    rusto_core::TableStructure::new(rusto_rows, rows.len(), num_cols).to_html()
}

/// One warning per process — operators must learn why the fields are absent,
/// but not once per page.
fn warn_layout_model_missing() {
    static WARNED: std::sync::atomic::AtomicBool = std::sync::atomic::AtomicBool::new(false);
    if !WARNED.swap(true, std::sync::atomic::Ordering::Relaxed) {
        eprintln!(
            "office-parser: ocr-rusto-core enabled but no layout model found \
             (set {LAYOUT_MODEL_ENV} to a layout model file, default \
             {DEFAULT_LAYOUT_MODEL}); layout_regions/table_html stay absent"
        );
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::document_ast::SourceSpan;

    fn para(text: &str) -> Block {
        Block::Paragraph {
            block_index: 0,
            text: text.to_string(),
            marks: Vec::new(),
            source: SourceSpan::default(),
        }
    }

    fn table_block(rows: &[&[&str]]) -> Block {
        Block::Table {
            block_index: 1,
            rows: rows
                .iter()
                .map(|r| {
                    r.iter()
                        .map(|t| Cell {
                            text: (*t).to_string(),
                            colspan: 1,
                            rowspan: 1,
                        })
                        .collect()
                })
                .collect(),
            widths: Vec::new(),
            source: SourceSpan::default(),
        }
    }

    fn page() -> OcrPage {
        OcrPage {
            blocks: vec![para("hello"), para("world")],
            confidence: Some(0.9),
            layout_regions: None,
            word_boxes: None,
            table_html: None,
        }
    }

    fn region(
        layout_type: rusto_core::LayoutType,
        x0: i32,
        y0: i32,
        x1: i32,
        y1: i32,
    ) -> rusto_core::LayoutRegion {
        rusto_core::LayoutRegion {
            bbox: rusto_core::BBox::new(x0, y0, x1, y1),
            layout_type,
            confidence: 0.9,
        }
    }

    /// No model ⇒ no-op: transcript + confidence untouched, rich fields None.
    /// `&[]` PNG proves no decode happens before the model resolves.
    #[test]
    fn enrich_no_model_leaves_page_untouched() {
        let page = page();
        let enriched = enrich_with_model(page, &[], None);
        assert_eq!(enriched.blocks.len(), 2);
        assert_eq!(enriched.confidence, Some(0.9));
        assert!(enriched.layout_regions.is_none());
        assert!(enriched.word_boxes.is_none());
        assert!(enriched.table_html.is_none());
    }

    /// Bogus model path ⇒ detector load fails, page returned unchanged.
    #[test]
    fn enrich_bogus_model_leaves_page_untouched() {
        let page = page();
        let enriched = enrich_with_model(
            page,
            &[],
            Some(Path::new("/nonexistent/office-parser-layout.onnx")),
        );
        assert_eq!(enriched.blocks.len(), 2);
        assert_eq!(enriched.confidence, Some(0.9));
        assert!(enriched.layout_regions.is_none());
        assert!(enriched.table_html.is_none());
    }

    /// Region filter keeps `LayoutType::Table` only, maps coords to f32.
    #[test]
    fn table_region_filter_keeps_table_only() {
        let regions = vec![
            region(rusto_core::LayoutType::Text, 0, 0, 10, 10),
            region(rusto_core::LayoutType::Table, 5, 6, 105, 206),
            region(rusto_core::LayoutType::Title, 1, 1, 2, 2),
            region(rusto_core::LayoutType::Table, 10, 20, 30, 40),
            region(rusto_core::LayoutType::Figure, 2, 2, 3, 3),
        ];
        let tables = table_regions_to_layout(&regions);
        assert_eq!(tables.len(), 2);
        assert_eq!(tables[0].label, "Table");
        assert_eq!((tables[0].x0, tables[0].y0), (5.0, 6.0));
        assert_eq!((tables[0].x1, tables[0].y1), (105.0, 206.0));
        assert_eq!((tables[1].x0, tables[1].y1), (10.0, 40.0));
    }

    /// Markup gate: needs BOTH a table region and a Table block.
    #[test]
    fn tables_html_requires_region_and_block() {
        let blocks = vec![para("p"), table_block(&[&["a", "b"]])];
        assert!(tables_html(&blocks, false).is_none(), "no region ⇒ None");
        assert!(tables_html(&[para("p")], true).is_none(), "no Table block ⇒ None");
        let html = tables_html(&blocks, true).expect("region + block ⇒ Some");
        assert!(html.contains("<table"));
        assert!(html.contains("<td>a</td>"));
        assert!(html.contains("</table>"));
    }

    /// rusto-core renders spans and escapes HTML text.
    #[test]
    fn table_structure_html_renders_spans_and_escapes() {
        let rows = vec![
            vec![Cell {
                text: "a<b&c".to_string(),
                colspan: 2,
                rowspan: 1,
            }],
            vec![
                Cell {
                    text: "x".to_string(),
                    colspan: 1,
                    rowspan: 1,
                },
                Cell {
                    text: "y".to_string(),
                    colspan: 1,
                    rowspan: 1,
                },
            ],
        ];
        let html = table_structure_html(&rows);
        assert!(html.contains("<td colspan=\"2\">a&lt;b&amp;c</td>"), "{html}");
        assert!(html.contains("<td>x</td>"), "{html}");
    }
}
