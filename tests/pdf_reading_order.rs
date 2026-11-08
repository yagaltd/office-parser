#![cfg(feature = "pdf")]
//! TASK 6 — XY-Cut reading order (`tests/pdf_reading_order.rs`).
//!
//! Contract: `.workflows/specs/task-6.spec` (plan 20261108-001).
//!
//! Scope honesty: the fixture corpus is single-column everywhere (TASK 3 measured
//! gutters=0 / columns=1 on every text-layer page), and `tests/fixtures/pdf_mixed.pdf`
//! is single-column too — the spec scenario text saying its page "holds more than one
//! column" is contradicted by that measurement and cannot be exercised by any fixture.
//! The multi-column path is therefore verified with synthetic geometry: a synthetic
//! two-column PDF (built with lopdf, already the `pdf`-feature dependency — no new
//! dependency) parsed end-to-end, plus pure `SpanBox` layouts through the public
//! `xy_cut_reading_regions`. On the real corpus, XY-Cut is a NO-OP guard:
//!
//! - every corpus page's x-projection merges into ONE cluster (asserted below), so no
//!   vertical cut can fire and the parser provably keeps the original stream path;
//! - a horizontal band cut cannot change output order (bands are y-disjoint, and the
//!   per-region concatenation equals the global top-to-bottom order);
//! - the byte-guard pins below verify both empirically: identical block output and
//!   identical `render::to_markdown` bytes on every fixture.
//!
//! Byte-identity strategy follows TASK 4's precedent (`tests/pdf_geometry.rs`): no
//! in-repo golden for the library markdown exists, so blocks are pinned exactly
//! (texts/levels/order from the byte-verified scout baseline) — both renderers are
//! pure functions of `Document.blocks`, so identical blocks pin identical bytes —
//! plus exact `to_markdown` pins where the full bytes are known.

use office_parser::document_ast::Block;
use office_parser::formats::pdf::{parse_pdf_full, xy_cut_reading_regions, SpanBox};
use office_parser::{pdf, render};

const MIXED: &str = "tests/fixtures/pdf_mixed.pdf";
const SCANNED: &str = "tests/fixtures/pdf_scanned.pdf";
const TABLE_SCAN: &str = "tests/fixtures/pdf_table_scan.pdf";
const ROT90: &str = "tests/fixtures/pdf_rot90_scan.pdf";
const CJK: &str = "tests/fixtures/pdf_cjk_scan.pdf";

/// The `pdf_mixed` page-1 paragraph texts (byte-verified scout baseline, pinned by
/// `tests/pdf_geometry.rs` at HEAD).
const MIXED_P1: &str = "During the second quarter the field team replaced 34 fiber patch panels across the northern ring. Downtime totalled eleven hours, split over four maintenance windows.";
const MIXED_P2: &str = "Customer impact was limited to two buildings on the Riverside campus. Affected tenants received credits according to the standard SLA schedule.";

/// Test-side replica of the gutter predicate's projection step: merge the page's
/// spans into x-projection clusters. `1` => no vertical cut can fire on the page
/// (a cut needs two clusters separated by a gap above threshold).
fn x_projection_clusters(spans: &[SpanBox]) -> usize {
    let mut iv: Vec<(f32, f32)> = spans.iter().map(|s| (s.x0, s.x1)).collect();
    iv.sort_by(|a, b| a.0.partial_cmp(&b.0).unwrap_or(std::cmp::Ordering::Equal));
    let mut clusters = 1usize;
    let mut hi = iv[0].1;
    for &(lo, r) in iv[1..].iter() {
        if lo > hi {
            clusters += 1;
        }
        if r > hi {
            hi = r;
        }
    }
    clusters
}

/// Synthetic `SpanBox` with a 12pt body font (thresholds: gutter >= 24pt, band >= 28pt).
fn span(seq: u32, x0: f32, y0: f32, x1: f32, y1: f32) -> SpanBox {
    SpanBox {
        seq,
        x0,
        y0,
        x1,
        y1,
        font_name: "F1".to_string(),
        font_size: 12.0,
        text: format!("s{seq}"),
        char_count: 2,
    }
}

/// Scenario: single-column output is byte-identical (spec scenario 1).
#[test]
fn test_single_column_output_unchanged() {
    // Corpus premise (TASK 3): every text-layer page merges into a single
    // x-projection cluster — no vertical gutter exists, so XY-Cut cannot split any
    // corpus page and the parser provably keeps the original stream path.
    for path in [MIXED, SCANNED, TABLE_SCAN, ROT90, CJK] {
        let bytes = std::fs::read(path).expect("fixture");
        let parsed = parse_pdf_full(&bytes).expect("parses");
        for g in &parsed.pages_geometry {
            if g.spans.is_empty() {
                continue;
            }
            assert_eq!(
                x_projection_clusters(&g.spans),
                1,
                "{path} page {}: a column gutter exists — XY-Cut would reorder this page and the byte-guard below would be meaningless",
                g.page_number
            );
        }
    }

    // pdf_scanned / pdf_table_scan: image-only, no blocks at all.
    for path in [SCANNED, TABLE_SCAN] {
        let bytes = std::fs::read(path).expect("fixture");
        let doc = pdf::parse(&bytes).expect("parses");
        assert!(doc.blocks.is_empty(), "{path}: expected no blocks");
        assert_eq!(render::to_markdown(&doc), "", "{path}: markdown");
    }

    // pdf_cjk_scan: exactly one paragraph, byte-exact markdown.
    {
        let bytes = std::fs::read(CJK).expect("fixture");
        let doc = pdf::parse(&bytes).expect("parses");
        assert_eq!(doc.blocks.len(), 1);
        assert_eq!(render::to_markdown(&doc), "‰º†ÊÑüÂô®ËÆ¢ÂçïÊëòË¶Å\n\n");
    }

    // pdf_rot90_scan: one image block (markdown prefix; full id is 64 hex chars).
    {
        let bytes = std::fs::read(ROT90).expect("fixture");
        let doc = pdf::parse(&bytes).expect("parses");
        assert_eq!(doc.blocks.len(), 1);
        assert!(render::to_markdown(&doc).starts_with("[image:sha256:ac2179740601"));
    }

    // pdf_mixed: L1 + L2 heading, two paragraphs, one image (page 1); page 2 empty.
    // Blocks pinned exactly (order + kind + level + full text) — identical blocks
    // pin identical rendered bytes because both renderers are pure functions.
    {
        let bytes = std::fs::read(MIXED).expect("fixture");
        let doc = pdf::parse(&bytes).expect("parses");
        assert_eq!(doc.blocks.len(), 5);
        let expectations: [(&str, Option<u8>, &str); 5] = [
            ("heading", Some(1), "Quarterly Field Report"),
            ("heading", Some(2), "Network Maintenance Summary"),
            ("paragraph", None, MIXED_P1),
            ("paragraph", None, MIXED_P2),
            ("image", None, "sha256:1f5353555f1f"),
        ];
        for (i, (kind, level, expected)) in expectations.iter().enumerate() {
            match (&doc.blocks[i], *kind) {
                (Block::Heading { level: l, text, .. }, "heading") => {
                    assert_eq!(*l, level.unwrap());
                    assert_eq!(text, *expected);
                }
                (Block::Paragraph { text, .. }, "paragraph") => {
                    assert_eq!(text, *expected);
                }
                (Block::Image { id, .. }, "image") => {
                    assert!(id.starts_with(*expected), "id: {id}");
                }
                (b, k) => panic!("block {i}: expected {k}, got {b:?}"),
            }
        }
        // Markdown byte-pin: everything up to the image id's first 12 hex chars is
        // known exactly; the remaining 52 hex chars are the sha256 itself.
        let md = render::to_markdown(&doc);
        let prefix = format!(
            "# Quarterly Field Report\n\n## Network Maintenance Summary\n\n{MIXED_P1}\n\n{MIXED_P2}\n\n[image:sha256:1f5353555f1f"
        );
        assert!(md.starts_with(&prefix), "markdown drift: {md}");
        assert!(md.ends_with(" alt=\"\"]\n\n"), "markdown image tail drift: {md}");
    }
}

/// Scenario: multi-column content is not interleaved (spec scenario 2).
#[test]
fn test_mixed_fixture_columns_not_interleaved() {
    // SCOPE HONESTY: the spec premise ("pdf_mixed ... holds more than one column")
    // is contradicted by the TASK 3 measurement — the fixture's digital page is
    // single-column (gutters=0; asserted in test_single_column_output_unchanged).
    // The known-sentence assertions below are still meaningful: no text may be
    // injected between a sentence and its own continuation, and the reading order
    // of the four known blocks is pinned. The genuine multi-column case (which no
    // fixture can exercise) is verified end-to-end by
    // test_synthetic_two_column_pdf_reads_columns_in_order below.
    let bytes = std::fs::read(MIXED).expect("fixture");
    let doc = pdf::parse(&bytes).expect("parses");
    let md = render::to_markdown(&doc);

    // Each known paragraph appears CONTIGUOUSLY — no column (or any other) text
    // may be injected inside a sentence's continuation.
    assert!(md.contains(MIXED_P1), "P1 not contiguous: {md}");
    assert!(md.contains(MIXED_P2), "P2 not contiguous: {md}");
    // The known sentence start is immediately followed by its own continuation.
    assert!(
        md.contains("across the northern ring. Downtime totalled"),
        "sentence continuation interrupted: {md}"
    );

    // Reading order: H1 < H2 < P1 < P2 — on known words, not a snapshot hash.
    let h1 = md.find("# Quarterly Field Report").expect("H1 present");
    let h2 = md.find("## Network Maintenance Summary").expect("H2 present");
    let p1 = md.find(MIXED_P1).expect("P1 present");
    let p2 = md.find(MIXED_P2).expect("P2 present");
    assert!(h1 < h2 && h2 < p1 && p1 < p2, "block reading order drifted");
}

/// Scenario: heading levels survive reordering (spec scenario 3, negative).
#[test]
fn test_heading_levels_preserved_after_reorder() {
    let bytes = std::fs::read(MIXED).expect("fixture");
    let doc = pdf::parse(&bytes).expect("parses");
    let mut headings: Vec<(u8, String)> = Vec::new();
    let mut paragraphs = 0usize;
    for b in &doc.blocks {
        match b {
            Block::Heading { level, text, .. } => headings.push((*level, text.clone())),
            Block::Paragraph { .. } => paragraphs += 1,
            _ => {}
        }
    }
    assert_eq!(
        headings,
        vec![
            (1u8, "Quarterly Field Report".to_string()),
            (2u8, "Network Maintenance Summary".to_string()),
        ],
        "headings must survive the (corpus no-op) reordering unchanged"
    );
    assert_eq!(paragraphs, 2, "a heading was demoted to Paragraph");
}

/// Scenario: empty page does not panic (spec scenario 4).
#[test]
fn test_empty_page_ordering_returns_no_blocks() {
    assert!(xy_cut_reading_regions(&[]).is_empty());

    let bytes = std::fs::read(SCANNED).expect("fixture");
    let parsed = parse_pdf_full(&bytes).expect("parses");
    assert_eq!(parsed.pages_geometry.len(), 2);
    for g in &parsed.pages_geometry {
        assert!(g.spans.is_empty(), "page {}: unexpected spans", g.page_number);
        assert!(xy_cut_reading_regions(&g.spans).is_empty());
    }
    assert!(parsed.blocks.is_empty());

    let bytes = std::fs::read(MIXED).expect("fixture");
    let parsed = parse_pdf_full(&bytes).expect("parses");
    assert_eq!(parsed.page_block_ranges.len(), 2);
    let p2 = &parsed.page_block_ranges[1];
    assert_eq!(p2.page_number, 2);
    // "Empty" means no TEXT spans — the page's embedded scan is still extracted
    // as exactly one Image block (ExtractedImage feature, rot90 precedent).
    assert_eq!(p2.block_end - p2.block_start, 1, "one Image block for the scan");
    match &parsed.blocks[p2.block_start] {
        Block::Image { id, .. } => assert!(id.starts_with("sha256:"), "id: {id}"),
        b => panic!("expected the page-scan Image block, got {b:?}"),
    }
    let g2 = parsed
        .pages_geometry
        .iter()
        .find(|g| g.page_number == 2)
        .expect("PageGeometry for page 2");
    assert!(g2.spans.is_empty());
    assert!(xy_cut_reading_regions(&g2.spans).is_empty());
}

/// Minimal one-page PDF builder: MediaBox [0 0 612 792], one Helvetica font,
/// `content` as the raw page content stream (lopdf is already the `pdf`-feature
/// dependency — no new dependency; same approach as tests/pdf_geometry.rs).
fn synth_one_page_pdf(content: &[u8]) -> Vec<u8> {
    use lopdf::{Dictionary, Document, Object, Stream};

    let nm = |s: &str| Object::Name(s.as_bytes().to_vec());
    let int = |v: i64| Object::Integer(v);
    let real = |v: f32| Object::Real(v);

    let mut doc = Document::new();
    let content_id = doc.add_object(Stream::new(Dictionary::new(), content.to_vec()));

    let mut font = Dictionary::new();
    font.set("Type", nm("Font"));
    font.set("Subtype", nm("Type1"));
    font.set("BaseFont", nm("Helvetica"));
    let font_id = doc.add_object(font);

    let pages_id = doc.add_object(Dictionary::new());
    let mut resources = Dictionary::new();
    let mut font_map = Dictionary::new();
    font_map.set("F1", Object::Reference(font_id));
    resources.set("Font", Object::Dictionary(font_map));

    let mut page = Dictionary::new();
    page.set("Type", nm("Page"));
    page.set("Parent", Object::Reference(pages_id));
    page.set(
        "MediaBox",
        Object::Array(vec![int(0), int(0), real(612.0), real(792.0)]),
    );
    page.set("Resources", Object::Dictionary(resources));
    page.set("Contents", Object::Reference(content_id));
    let page_id = doc.add_object(page);

    let mut pages_dict = Dictionary::new();
    pages_dict.set("Type", nm("Pages"));
    pages_dict.set("Kids", Object::Array(vec![Object::Reference(page_id)]));
    pages_dict.set("Count", int(1));
    doc.set_object(pages_id, pages_dict);

    let mut catalog = Dictionary::new();
    catalog.set("Type", nm("Catalog"));
    catalog.set("Pages", Object::Reference(pages_id));
    let catalog_id = doc.add_object(catalog);
    doc.trailer.set("Root", Object::Reference(catalog_id));

    let mut buf = Vec::new();
    doc.save_to(&mut buf).expect("lopdf save");
    buf
}

/// Scenario: multi-column content is not interleaved — the path no fixture can
/// exercise, verified end-to-end with a synthetic two-column PDF.
#[test]
fn test_synthetic_two_column_pdf_reads_columns_in_order() {
    // Stream order is deliberately interleaved (L1, R1, L2, R2 at the same two
    // y-bands): pre-TASK-6 this interleaves into "L1 R1" / "L2 R2" lines, so this
    // test is non-vacuous — it fails without the XY-Cut reorder.
    let content = b"BT /F1 12 Tf 72 720 Td (Alpha one two three.) Tj ET\n\
BT /F1 12 Tf 312 720 Td (Beta seven eight nine.) Tj ET\n\
BT /F1 12 Tf 72 700 Td (Alpha four five six.) Tj ET\n\
BT /F1 12 Tf 312 700 Td (Beta ten eleven twelve.) Tj ET";
    let bytes = synth_one_page_pdf(content);

    let parsed = parse_pdf_full(&bytes).expect("parses");
    assert_eq!(parsed.pages_geometry.len(), 1);
    let g = &parsed.pages_geometry[0];
    assert_eq!(g.spans.len(), 4);
    let stream: Vec<&str> = g.spans.iter().map(|s| s.text.as_str()).collect();
    assert_eq!(
        stream,
        [
            "Alpha one two three.",
            "Beta seven eight nine.",
            "Alpha four five six.",
            "Beta ten eleven twelve.",
        ],
        "fixture self-check: stream order must be interleaved or the test proves nothing"
    );

    // The vertical gutter (x: [72..~201] vs [312..~455], gap ~111pt at a 12pt font)
    // splits left/right; each column is one leaf region, top-to-bottom.
    assert_eq!(
        xy_cut_reading_regions(&g.spans),
        [vec![0u32, 2], vec![1u32, 3]],
        "expected exactly one vertical cut: left column then right column"
    );

    let doc = pdf::parse(&bytes).expect("parses");
    let expected = [
        "Alpha one two three.",
        "Alpha four five six.",
        "Beta seven eight nine.",
        "Beta ten eleven twelve.",
    ];
    assert_eq!(doc.blocks.len(), expected.len(), "four paragraphs, no cover heading");
    for (i, want) in expected.iter().enumerate() {
        match &doc.blocks[i] {
            Block::Paragraph { text, .. } => assert_eq!(text, want, "block {i}"),
            other => panic!("block {i}: expected Paragraph, got {other:?}"),
        }
    }
    // Full left column precedes the right column; markdown is byte-pinned.
    assert_eq!(
        render::to_markdown(&doc),
        "Alpha one two three.\n\nAlpha four five six.\n\nBeta seven eight nine.\n\nBeta ten eleven twelve.\n\n"
    );
}

#[test]
fn test_xy_cut_full_width_header_above_two_columns() {
    // The classic layout a vertical-only cut cannot handle: a full-width row
    // BRIDGES the gutter, so the top-level x-projection has no gap. The
    // horizontal band cut peels the header first, then the vertical cut
    // separates the columns (task-6 spec diagram: vertical fails => horizontal).
    let spans = [
        span(0, 72.0, 60.0, 540.0, 76.0),   // full-width header (bridges gutter)
        span(1, 72.0, 120.0, 200.0, 132.0), // left col line 1
        span(2, 72.0, 150.0, 200.0, 162.0), // left col line 2
        span(3, 312.0, 120.0, 440.0, 132.0), // right col line 1
        span(4, 312.0, 150.0, 440.0, 162.0), // right col line 2
    ];
    let regions = xy_cut_reading_regions(&spans);
    assert_eq!(regions, [vec![0u32], vec![1u32, 2], vec![3u32, 4]]);
    // Regions partition the page: every seq exactly once, header first.
    let mut flat = regions.iter().flatten().copied().collect::<Vec<_>>();
    flat.sort_unstable();
    assert_eq!(flat, vec![0, 1, 2, 3, 4]);
}

#[test]
fn test_xy_cut_narrow_gap_is_not_a_gutter() {
    // A 20pt x-gap at a 12pt body font is below the gutter threshold
    // (max(2x fs, 12) = 24pt): no column split — e.g. wide table cells or an
    // indented block must not be read as columns.
    let spans = [
        span(0, 72.0, 100.0, 200.0, 112.0),
        span(1, 220.0, 100.0, 360.0, 112.0),
    ];
    let regions = xy_cut_reading_regions(&spans);
    assert_eq!(regions.len(), 1, "20pt < 24pt threshold: no split");

    // Just above the threshold the split fires, left region first.
    let spans = [
        span(0, 72.0, 100.0, 200.0, 112.0),
        span(1, 260.0, 100.0, 420.0, 112.0),
    ];
    let regions = xy_cut_reading_regions(&spans);
    assert_eq!(regions, [vec![0u32], vec![1u32]]);
}

#[test]
fn test_xy_cut_single_column_stays_one_region() {
    // Stacked full-width blocks with paragraph-sized gaps: no vertical gutter,
    // and paragraph gaps are below the band threshold — one region, stream path.
    let spans = [
        span(0, 72.0, 60.0, 540.0, 76.0),
        span(1, 72.0, 100.0, 540.0, 112.0),
        span(2, 72.0, 130.0, 540.0, 142.0),
    ];
    let regions = xy_cut_reading_regions(&spans);
    assert_eq!(regions.len(), 1);
    assert_eq!(regions[0], vec![0u32, 1, 2]);
}

#[test]
fn test_xy_cut_leaf_orders_top_to_bottom_not_by_stream_seq() {
    // Inside a region the order is geometric (ascending y0 = downwards), not
    // stream order: seq 0 sits BELOW seq 1 and must be emitted second.
    let spans = [
        span(0, 60.0, 84.0, 180.0, 96.0),  // left col, lower line
        span(1, 60.0, 60.0, 180.0, 72.0),  // left col, upper line
        span(2, 320.0, 60.0, 460.0, 72.0), // right col
    ];
    let regions = xy_cut_reading_regions(&spans);
    assert_eq!(regions, [vec![1u32, 0], vec![2u32]]);
}
