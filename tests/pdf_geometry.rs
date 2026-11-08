#![cfg(feature = "pdf")]
//! TASK 4 — `PageGeometry`/`SpanBox` emission from the PDF content interpreter.
//!
//! Contract: `.workflows/specs/task-4.spec` (plan 20261108-001).
//!
//! Byte-identity strategy (`test_existing_block_output_unchanged`): no in-repo
//! pre-change golden exists for `render::to_markdown` / `to_json_value`, so the test
//! pins the full block output of all five corpus fixtures (counts, exact texts,
//! heading levels, image-id prefixes). Both renderers are pure functions of
//! `Document.blocks`, so identical blocks pin identical rendered bytes. Text pins
//! come from the pre-change scout baseline
//! (`.workflows/spikes/scout/baseline/`, byte-verified); image ids are pinned by
//! their first 12 sha256 chars (cli truncation, cli/src/main.rs:380-381). The gate
//! additionally runs the CLI on `tests/fixtures/pdf_scanned.pdf` before and after
//! this change — the markdown must be byte-identical.

use office_parser::document_ast::Block;
use office_parser::formats::pdf::{parse_pdf_full, CoordinateSystem, PageGeometry};
use office_parser::{pdf, render};

const MIXED: &str = "tests/fixtures/pdf_mixed.pdf";
const SCANNED: &str = "tests/fixtures/pdf_scanned.pdf";
const TABLE_SCAN: &str = "tests/fixtures/pdf_table_scan.pdf";
const ROT90: &str = "tests/fixtures/pdf_rot90_scan.pdf";
const CJK: &str = "tests/fixtures/pdf_cjk_scan.pdf";

fn approx(a: f32, b: f32) -> bool {
    (a - b).abs() < 1e-2
}

fn page_geo(pages: &[PageGeometry], n: u32) -> &PageGeometry {
    pages
        .iter()
        .find(|g| g.page_number == n)
        .unwrap_or_else(|| panic!("no PageGeometry for page {n}"))
}

/// Build a minimal one-page PDF with lopdf (already the `pdf`-feature dependency —
/// no new dependency): MediaBox [0 0 612 792], CropBox [100 100 500 700], one
/// Helvetica text op at pen (150, 650), optional /Rotate.
fn synth_pdf(rotate: Option<i64>) -> Vec<u8> {
    use lopdf::{dictionary, Dictionary, Document, Object, Stream};

    let nm = |s: &str| Object::Name(s.as_bytes().to_vec());
    let int = |v: i64| Object::Integer(v);
    let real = |v: f32| Object::Real(v);

    let mut doc = Document::new();

    let content_id = doc.add_object(Stream::new(
        Dictionary::new(),
        b"BT /F1 18 Tf 150 650 Td (Crop me) Tj ET".to_vec(),
    ));

    let mut font = Dictionary::new();
    font.set("Type", nm("Font"));
    font.set("Subtype", nm("Type1"));
    font.set("BaseFont", nm("Helvetica"));
    let font_id = doc.add_object(font);

    let mut pages = Dictionary::new();
    pages.set("Type", nm("Pages"));
    let pages_id = doc.add_object(pages);

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
    page.set(
        "CropBox",
        Object::Array(vec![int(100), int(100), real(500.0), real(700.0)]),
    );
    page.set("Resources", Object::Dictionary(resources));
    page.set("Contents", Object::Reference(content_id));
    if let Some(rot) = rotate {
        page.set("Rotate", int(rot));
    }
    let page_id = doc.add_object(page);

    let mut pages_dict = Dictionary::new();
    pages_dict.set("Type", nm("Pages"));
    pages_dict.set(
        "Kids",
        Object::Array(vec![Object::Reference(page_id)]),
    );
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

/// Scenario: every span carries a box, font name and font size.
#[test]
fn test_spans_carry_bbox_font_name_and_size() {
    let bytes = std::fs::read(MIXED).expect("fixture");
    let parsed = parse_pdf_full(&bytes).expect("parses");

    assert_eq!(parsed.pages_geometry.len(), parsed.page_block_ranges.len());
    let p1 = page_geo(&parsed.pages_geometry, 1);
    assert!(!p1.spans.is_empty(), "page 1 is the born-digital text page");
    for s in &p1.spans {
        assert!(s.x1 > s.x0, "span {}: x1 ({}) <= x0 ({})", s.seq, s.x1, s.x0);
        assert!(s.y1 > s.y0, "span {}: y1 ({}) <= y0 ({})", s.seq, s.y1, s.y0);
        assert!(!s.font_name.is_empty(), "span {}: empty font_name", s.seq);
        assert!(s.font_size > 0.0, "span {}: font_size <= 0", s.seq);
        assert_eq!(s.char_count, s.text.chars().count());
    }
    // Page 2 is the image-only half of the fixture: no text ops, no spans.
    let p2 = page_geo(&parsed.pages_geometry, 2);
    assert!(p2.spans.is_empty(), "image-only page must have no spans");
}

/// Scenario: crop-box origin is subtracted.
///
/// Expected span for this fixture (deterministic): pen (150, 650), fs 18,
/// advance = (0.55*6 + 0.33)*18 = 65.34 → normalized box x0=50, x1≈115.34,
/// y0=32 (700-668), y1=50 (700-650).
#[test]
fn test_cropbox_origin_is_subtracted() {
    let bytes = synth_pdf(None);
    let parsed = parse_pdf_full(&bytes).expect("parses");
    assert_eq!(parsed.pages_geometry.len(), 1);
    let g = &parsed.pages_geometry[0];

    // The ORIGINAL crop box is recorded, so the subtraction stays reversible.
    assert_eq!(g.media_box, Some([0.0, 0.0, 612.0, 792.0]));
    assert_eq!(g.crop_box, Some([100.0, 100.0, 500.0, 700.0]));
    assert_eq!(g.coordinate_system, CoordinateSystem::Source);

    assert_eq!(g.spans.len(), 1);
    let s = &g.spans[0];
    assert!(s.x0 >= 0.0, "crop origin not subtracted: x0 = {}", s.x0);
    assert_eq!(s.x0, 50.0, "x0 must be pen 150 - crop origin 100");
    assert_eq!(s.y0, 32.0, "top edge = max_y 700 - raw y1 668");
    assert_eq!(s.y1, 50.0, "bottom edge = max_y 700 - raw y0 650");
    assert!(s.x1 > 115.0 && s.x1 < 116.0, "advance-bounded x1, got {}", s.x1);
    assert!(s.x1 > s.x0 && s.y1 > s.y0);
    // Reversible: adding the recorded origin recovers the raw pen x.
    let origin_x = g.crop_box.unwrap()[0];
    assert!((s.x0 + origin_x - 150.0).abs() < 1e-4);
    assert_eq!(s.seq, 0);
    assert_eq!(s.font_name, "F1");
    assert_eq!(s.font_size, 18.0);
    assert_eq!(s.text, "Crop me");
    assert_eq!(s.char_count, 7);
}

/// Scenario: native rotation is preserved.
#[test]
fn test_initial_rotation_preserved() {
    // /Rotate 90 and 270 survive as `initial_rotation` even though boxes are NOT
    // rotated (coordinate_system stays "source" in this task).
    for (rotate, expected) in [(Some(90i64), 90u16), (Some(270), 270), (None, 0)] {
        let bytes = synth_pdf(rotate);
        let parsed = parse_pdf_full(&bytes).expect("parses");
        let g = &parsed.pages_geometry[0];
        assert_eq!(g.initial_rotation, expected, "rotate={rotate:?}");
        assert_eq!(g.coordinate_system, CoordinateSystem::Source);
    }

    // The corpus fixture does NOT actually carry /Rotate (no `Rotate` key anywhere in
    // the file — its rot90-ness is the swapped 792x612 landscape MediaBox), so its
    // initial_rotation is the inherited default 0; page boxes must still resolve.
    let bytes = std::fs::read(ROT90).expect("fixture");
    let parsed = parse_pdf_full(&bytes).expect("parses");
    let g = &parsed.pages_geometry[0];
    assert_eq!(g.initial_rotation, 0);
    assert_eq!(g.coordinate_system, CoordinateSystem::Source);
    assert!(g.spans.is_empty());
    let media = g.media_box.expect("MediaBox resolved");
    assert!(approx(media[2], 792.05) && approx(media[3], 612.04), "media: {media:?}");
}

/// Scenario: existing block output is byte-identical (negative).
#[test]
fn test_existing_block_output_unchanged() {
    // pdf_scanned / pdf_table_scan: image-only, no blocks at all.
    for path in [SCANNED, TABLE_SCAN] {
        let bytes = std::fs::read(path).expect("fixture");
        let doc = pdf::parse(&bytes).expect("parses");
        assert!(doc.blocks.is_empty(), "{path}: expected no blocks");
        assert_eq!(render::to_markdown(&doc), "", "{path}: markdown");
        let j = render::to_json_value(&doc);
        assert_eq!(j["blocks"].as_array().unwrap().len(), 0, "{path}: json blocks");
    }

    // pdf_cjk_scan: exactly one paragraph (the 16-char latin-1 mojibake layer).
    {
        let bytes = std::fs::read(CJK).expect("fixture");
        let doc = pdf::parse(&bytes).expect("parses");
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Paragraph { text, .. } => assert_eq!(text, "‰º†ÊÑüÂô®ËÆ¢ÂçïÊëòË¶Å"),
            other => panic!("expected Paragraph, got {other:?}"),
        }
        assert_eq!(render::to_markdown(&doc), "‰º†ÊÑüÂô®ËÆ¢ÂçïÊëòË¶Å\n\n");
        let j = render::to_json_value(&doc);
        assert_eq!(j["blocks"][0]["type"], "paragraph");
        assert_eq!(j["blocks"][0]["text"], "‰º†ÊÑüÂô®ËÆ¢ÂçïÊëòË¶Å");
    }

    // pdf_rot90_scan: one image block (id prefix = first 12 sha256 chars).
    {
        let bytes = std::fs::read(ROT90).expect("fixture");
        let doc = pdf::parse(&bytes).expect("parses");
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Image { id, .. } => {
                assert!(id.starts_with("sha256:ac2179740601"), "id: {id}");
            }
            other => panic!("expected Image, got {other:?}"),
        }
        let md = render::to_markdown(&doc);
        assert!(md.starts_with("[image:sha256:ac2179740601"), "md: {md}");
        let j = render::to_json_value(&doc);
        assert_eq!(j["blocks"][0]["type"], "image");
        assert!(
            j["blocks"][0]["id"].as_str().unwrap().starts_with("sha256:ac2179740601"),
            "json id: {}", j["blocks"][0]["id"]
        );
    }

    // pdf_mixed: L1 + L2 heading, two paragraphs, one image (page 1); page 2 empty.
    {
        let bytes = std::fs::read(MIXED).expect("fixture");
        let doc = pdf::parse(&bytes).expect("parses");
        assert_eq!(doc.blocks.len(), 5);
        let expectations: [(&str, Option<u8>, &str); 5] = [
            ("heading", Some(1), "Quarterly Field Report"),
            ("heading", Some(2), "Network Maintenance Summary"),
            ("paragraph", None, "During the second quarter the field team replaced 34 fiber patch panels across the northern ring. Downtime totalled eleven hours, split over four maintenance windows."),
            ("paragraph", None, "Customer impact was limited to two buildings on the Riverside campus. Affected tenants received credits according to the standard SLA schedule."),
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
        let j = render::to_json_value(&doc);
        assert_eq!(j["blocks"].as_array().unwrap().len(), 5);
    }
}

/// Scenario: an empty page does not panic — one PageGeometry per page, empty
/// spans, page boxes still resolved.
#[test]
fn test_page_without_spans_yields_empty_geometry() {
    let bytes = std::fs::read(SCANNED).expect("fixture");
    let parsed = parse_pdf_full(&bytes).expect("parses");
    assert_eq!(parsed.pages_geometry.len(), 2);
    for (i, g) in parsed.pages_geometry.iter().enumerate() {
        assert_eq!(g.page_number, (i + 1) as u32);
        assert!(g.spans.is_empty(), "page {} must have no spans", g.page_number);
        // the scanner fixture declares a CropBox — it must resolve, not vanish
        let crop = g.crop_box.expect("CropBox resolved");
        assert!(approx(crop[2], 612.041) && approx(crop[3], 792.053), "crop: {crop:?}");
        let media = g.media_box.expect("MediaBox resolved");
        assert!(
            approx(media[2], 612.041) && approx(media[3], 792.053),
            "media: {media:?}"
        );
        assert_eq!(g.initial_rotation, 0);
        assert_eq!(g.coordinate_system, CoordinateSystem::Source);
    }
}
