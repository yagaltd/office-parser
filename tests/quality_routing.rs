#![cfg(feature = "pdf")]
//! TASK 7 (plan 20261108-001): per-page quality population + routing verdicts.
//!
//! Routing-table contract: `.workflows/specs/task-7.spec` (RE-SCOPED) ×
//! `.workflows/spikes/pdf-complexity/README.md` §Score table — the four
//! image-only fixtures route to `Ocr` (image-only gate), `pdf_cjk_scan`'s
//! thin mojibake layer routes to `Ocr` via the 50-char threshold, and
//! `pdf_mixed` p1 (digital: rich, single-column, upright) routes `Trust`.
//!
//! Verdicts describe the TEXT LAYER, so the plain text-layer parse
//! (`pdf::parse`) exercises them deterministically — no OCR engine, no
//! pdfium needed.

use office_parser::document_ast::{Block, SourceSpan};
use office_parser::formats::pdf::{
    parse_pdf_split, PageGeometry, PageOcrEngine, ParsedPdfDocument, PdfPageBlockRange, SpanBox,
};
use office_parser::quality::{count_columns, page_verdict, PageTextQuality, PageVerdict};

const MIXED: &str = "tests/fixtures/pdf_mixed.pdf";
const SCANNED: &str = "tests/fixtures/pdf_scanned.pdf";
const TABLE_SCAN: &str = "tests/fixtures/pdf_table_scan.pdf";
const ROT90: &str = "tests/fixtures/pdf_rot90_scan.pdf";
const CJK: &str = "tests/fixtures/pdf_cjk_scan.pdf";

fn parse_pages(path: &str) -> Vec<PageTextQuality> {
    let bytes = std::fs::read(path).unwrap_or_else(|e| panic!("fixture {path}: {e}"));
    let doc = office_parser::pdf::parse(&bytes).expect("parses");
    doc.metadata
        .page_quality
        .unwrap_or_else(|| panic!("pdf parse must populate metadata.page_quality ({path})"))
}

fn expect_ocr_image_only(q: &PageTextQuality) {
    assert_eq!(
        q.text_richness, 0,
        "image-only page has no text layer: {q:?}"
    );
    assert_eq!(q.columns, 0, "no spans to measure on {q:?}");
    match &q.verdict {
        PageVerdict::Ocr { reason } => assert_eq!(reason, "image-only page"),
        other => panic!("expected Ocr (image-only), got {other:?}"),
    }
}

/// The five-fixture routing table (spec × spike §Score table).
#[test]
fn five_fixture_routing_table() {
    let pages = parse_pages(SCANNED);
    assert_eq!(pages.len(), 2);
    assert_eq!(pages[0].page, 1);
    assert_eq!(pages[1].page, 2);
    expect_ocr_image_only(&pages[0]);
    expect_ocr_image_only(&pages[1]);

    let pages = parse_pages(TABLE_SCAN);
    assert_eq!(pages.len(), 1);
    expect_ocr_image_only(&pages[0]);

    let pages = parse_pages(ROT90);
    assert_eq!(pages.len(), 1);
    expect_ocr_image_only(&pages[0]);

    // pdf_cjk_scan p1 — a real but thin text layer: routes Ocr via the char
    // threshold, NOT via image-only.
    let pages = parse_pages(CJK);
    assert_eq!(pages.len(), 1);
    let p1 = &pages[0];
    assert!(
        p1.text_richness > 0,
        "cjk fixture carries a real text layer, got {}",
        p1.text_richness
    );
    assert!(
        p1.text_richness < 50,
        "cjk layer sits below the 50-char gate, got {}",
        p1.text_richness
    );
    match &p1.verdict {
        PageVerdict::Ocr { reason } => assert!(
            reason.contains("chars <"),
            "char-threshold reason, got {reason:?}"
        ),
        other => panic!("expected Ocr via char threshold, got {other:?}"),
    }

    // pdf_mixed p1 — digital page: rich text layer, single column, upright.
    let pages = parse_pages(MIXED);
    assert_eq!(pages.len(), 2);
    let p1 = &pages[0];
    assert_eq!(
        p1.verdict,
        PageVerdict::Trust,
        "digital page trusts its text layer, got {p1:?}"
    );
    assert!(
        p1.text_richness >= 50,
        "rich layer above the gate, got {}",
        p1.text_richness
    );
    assert_eq!(p1.columns, 1, "single-column (spike §Score table)");
    assert!(!p1.rotated);
    expect_ocr_image_only(&pages[1]);
}

/// Split path without an engine: the combined doc carries the full table,
/// each page doc carries exactly its own record.
#[test]
fn split_pages_carry_their_single_page_verdict() {
    let bytes = std::fs::read(MIXED).expect("fixture");
    let (combined, pages) =
        parse_pdf_split(&bytes, None, 150, 50, false).expect("split parse");

    let combined_quality = combined.metadata.page_quality.expect("combined quality");
    assert_eq!(combined_quality.len(), 2);
    assert_eq!(combined_quality[0].verdict, PageVerdict::Trust);
    assert!(matches!(combined_quality[1].verdict, PageVerdict::Ocr { .. }));

    assert_eq!(pages.len(), 2);
    for (i, (page_number, page_doc)) in pages.iter().enumerate() {
        assert_eq!(*page_number, (i + 1) as u32);
        let pq = page_doc
            .metadata
            .page_quality
            .as_ref()
            .expect("page doc carries its quality record");
        assert_eq!(pq.len(), 1);
        assert_eq!(pq[0], combined_quality[i]);
    }
}

/// Verdict gate order: image-only wins over rotation; sparse layer states the
/// threshold; rich upright layer trusts; rich rotated layer escalates.
#[test]
fn page_verdict_gate_order() {
    assert_eq!(
        page_verdict(0, 50, false),
        PageVerdict::Ocr {
            reason: "image-only page".to_string()
        }
    );
    assert_eq!(
        page_verdict(0, 50, true),
        PageVerdict::Ocr {
            reason: "image-only page".to_string()
        },
        "image-only routes to Ocr even when rotated"
    );
    assert!(matches!(
        page_verdict(16, 50, false),
        PageVerdict::Ocr { .. }
    ));
    assert!(matches!(
        page_verdict(16, 50, true),
        PageVerdict::Ocr { .. }
    ));
    assert_eq!(page_verdict(305, 50, false), PageVerdict::Trust);
    assert_eq!(
        page_verdict(305, 50, true),
        PageVerdict::Escalate {
            complexity: Some("rotated page".to_string())
        }
    );
}

fn span(x0: f32, y0: f32, x1: f32, y1: f32, font_size: f32) -> SpanBox {
    SpanBox {
        seq: 0,
        x0,
        y0,
        x1,
        y1,
        font_name: "F1".to_string(),
        font_size,
        text: "x".to_string(),
        char_count: 1,
    }
}

/// Column counting: single column, two columns, full-width heading must not
/// mask the gutter, empty page measures 0.
#[test]
fn count_columns_over_synthetic_layouts() {
    assert_eq!(count_columns(&[]), 0, "image-only page: nothing to measure");

    let one_col: Vec<SpanBox> = (0..4)
        .map(|i| {
            if i % 2 == 0 {
                span(60.0, 100.0 + i as f32 * 20.0, 460.0, 112.0 + i as f32 * 20.0, 12.0)
            } else {
                span(60.0, 100.0 + i as f32 * 20.0, 400.0, 112.0 + i as f32 * 20.0, 12.0)
            }
        })
        .collect();
    assert_eq!(count_columns(&one_col), 1);

    // Two columns + one full-width heading. Body lines sit left [60,290] /
    // right [320,550]; the 30pt gap >= max(2x median, 12pt) is one gutter.
    let mut two_col: Vec<SpanBox> = vec![span(60.0, 60.0, 550.0, 74.0, 14.0)];
    for row in 0..4 {
        let y = 100.0 + row as f32 * 20.0;
        two_col.push(span(60.0, y, 290.0, y + 12.0, 12.0));
        two_col.push(span(320.0, y, 550.0, y + 12.0, 12.0));
    }
    assert_eq!(count_columns(&two_col), 2, "heading must not mask the gutter");

    let single_line = vec![span(60.0, 100.0, 550.0, 112.0, 12.0)];
    assert_eq!(count_columns(&single_line), 1);
}

/// `recompute_page_block_ranges` follows the merged layout and keeps the
/// metadata_json copy in agreement (the stale-ranges scout defect).
#[test]
fn recompute_page_block_ranges_follows_merged_layout() {
    let mut parsed = ParsedPdfDocument {
        blocks: Vec::new(),
        images: Vec::new(),
        metadata_json: serde_json::json!({
            "page_block_ranges": [
                {"page_number": 1, "block_start": 0, "block_end": 3},
                {"page_number": 2, "block_start": 3, "block_end": 5},
            ]
        }),
        page_block_ranges: vec![
            PdfPageBlockRange { page_number: 1, block_start: 0, block_end: 3 },
            PdfPageBlockRange { page_number: 2, block_start: 3, block_end: 5 },
        ],
        pages_geometry: Vec::new(),
    };
    parsed.recompute_page_block_ranges(&[2, 3]);
    assert_eq!(parsed.page_block_ranges[0].block_start, 0);
    assert_eq!(parsed.page_block_ranges[0].block_end, 2);
    assert_eq!(parsed.page_block_ranges[1].block_start, 2);
    assert_eq!(parsed.page_block_ranges[1].block_end, 5);
    assert_eq!(
        parsed.metadata_json["page_block_ranges"][1]["block_start"],
        serde_json::json!(2)
    );
    assert_eq!(
        parsed.metadata_json["page_block_ranges"][1]["block_end"],
        serde_json::json!(5)
    );
}

/// Geometry pass sanity on the routing path: every parsed page appears in
/// pages_geometry in page order.
#[test]
fn pages_geometry_covers_every_page_in_order() {
    for path in [MIXED, SCANNED, TABLE_SCAN, ROT90, CJK] {
        let bytes = std::fs::read(path).expect("fixture");
        let parsed = office_parser::formats::pdf::parse_pdf_full(&bytes).expect("parses");
        let numbers: Vec<u32> = parsed
            .pages_geometry
            .iter()
            .map(|g: &PageGeometry| g.page_number)
            .collect();
        let expected: Vec<u32> = (1..=parsed.page_block_ranges.len() as u32).collect();
        assert_eq!(numbers, expected, "{path}");
    }
}

/// End-to-end wiring of the stale-ranges fix (needs pdftoppm; skips loudly).
#[test]
fn merged_combined_doc_page_block_ranges_stay_fresh() {
    if !pdftoppm_available() {
        eprintln!("skipping: pdftoppm unavailable (merge path needs a rasterizer)");
        return;
    }
    struct EchoEngine;
    impl PageOcrEngine for EchoEngine {
        fn ocr_page_png(
            &mut self,
            _png: &[u8],
        ) -> Result<(Vec<Block>, Option<f64>), String> {
            Ok((
                vec![Block::Paragraph {
                    block_index: 0,
                    text: "ocr echo".to_string(),
                    marks: Vec::new(),
                    source: SourceSpan::default(),
                }],
                Some(0.9),
            ))
        }
        fn name(&self) -> &'static str {
            "echo"
        }
    }

    let bytes = std::fs::read(SCANNED).expect("fixture");
    let engine: &mut dyn PageOcrEngine = &mut EchoEngine;
    let (combined, pages) =
        parse_pdf_split(&bytes, Some(engine), 150, 50, false).expect("split parse");

    let ranges = combined.metadata.extra["page_block_ranges"]
        .as_array()
        .expect("ranges array");
    assert_eq!(ranges.len(), 2);
    assert_eq!(ranges[0]["block_start"], serde_json::json!(0));
    assert_eq!(
        ranges.last().unwrap()["block_end"],
        serde_json::json!(combined.blocks.len()),
        "ranges must cover the MERGED block vec (stale before the fix)"
    );
    for w in ranges.windows(2) {
        assert_eq!(w[0]["block_end"], w[1]["block_start"], "ranges contiguous");
    }

    assert_eq!(pages.len(), 2);
    for (i, (_, page_doc)) in pages.iter().enumerate() {
        let pq = page_doc
            .metadata
            .page_quality
            .as_ref()
            .expect("page doc quality");
        assert_eq!(pq.len(), 1);
        assert_eq!(pq[0].page, (i + 1) as u32);
        assert!(matches!(pq[0].verdict, PageVerdict::Ocr { .. }));
    }
}

fn pdftoppm_available() -> bool {
    std::env::var_os("PATH")
        .map(|p| {
            std::env::split_paths(&p)
                .any(|dir| dir.join("pdftoppm").is_file() || dir.join("pdftoppm.exe").is_file())
        })
        .unwrap_or(false)
}
