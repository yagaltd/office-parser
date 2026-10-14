#![cfg(all(feature = "pdf", feature = "ocr-indra-native"))]
//! Indra OCR pipeline tests (native in-process engine).
//!
//! Fixture provenance: `pdf_scanned.pdf` was produced by rasterizing a
//! two-page text PDF at 150 dpi (pdftoppm) and re-wrapping the PNGs into an
//! image-only PDF — a genuine scanned document with no text layer.

use office_parser::{Document, PdfTextQuality};

const SCAN: &str = "tests/fixtures/pdf_scanned.pdf";

/// The scanned fixture must classify as ImageOnly — the OCR trigger.
#[test]
fn scanned_fixture_classifies_as_image_only() {
    let bytes = std::fs::read(SCAN).expect("fixture");
    let doc = office_parser::pdf::parse(&bytes).expect("parses");
    assert_eq!(doc.metadata.pdf_text_quality, Some(PdfTextQuality::ImageOnly));
    // nearly no text without OCR
    let chars: usize = doc
        .blocks
        .iter()
        .map(|b| match b {
            office_parser::document_ast::Block::Paragraph { text, .. } => text.len(),
            office_parser::document_ast::Block::Heading { text, .. } => text.len(),
            _ => 0,
        })
        .sum();
    assert!(chars < 50, "scanned fixture should carry almost no text layer, got {chars}");
}


/// Empty-text and page-furniture blocks are dropped; consecutive List-items

/// Full pipeline: render → indra subprocess → blocks.

fn which_available(bin: &str) -> bool {
    std::env::var_os("PATH")
        .map(|p| {
            std::env::split_paths(&p).any(|dir| {
                dir.join(bin).is_file() || dir.join(format!("{bin}.exe")).is_file()
            })
        })
        .unwrap_or(false)
}

/// Mode B: the native (in-process ort) engine on the same fixture must
/// reproduce the Python engine's content. Skips without model files.
#[test]
fn native_engine_end_to_end() {
    let Ok(mut engine) = office_parser::IndraNative::load() else {
        eprintln!(
            "skipping: Indra models unavailable ({})",
            office_parser::IndraNative::load().err().unwrap_or_default()
        );
        return;
    };
    let bytes = std::fs::read(SCAN).expect("fixture");
    let doc = engine
        .parse_document(&bytes, 0.60)
        .expect("native ocr parse");
    use office_parser::document_ast::Block;

    let headings: Vec<&str> = doc
        .blocks
        .iter()
        .filter_map(|b| match b {
            Block::Heading { text, .. } => Some(text.as_str()),
            _ => None,
        })
        .collect();
    // The Python engine (golden JSONL) finds all three section headers.
    for expected in ["Quarterly Field Report", "Network Maintenance Summary", "Recommendations"] {
        assert!(
            headings.iter().any(|h| h.contains(expected)),
            "missing heading {expected:?}; got {headings:?}"
        );
    }
    let paragraphs: Vec<&str> = doc
        .blocks
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph { text, .. } => Some(text.as_str()),
            _ => None,
        })
        .collect();
    assert!(
        paragraphs
            .iter()
            .any(|t| t.contains("eleven hours, split over four maintenance windows")),
        "got {paragraphs:?}"
    );

    let ocr = doc.metadata.extra.get("ocr").expect("ocr metadata");
    assert_eq!(ocr["engine"], "indra-native");
    assert_eq!(ocr["pages_ocr"], 2);
}

/// Per-page routing: a mixed PDF (digital page + scanned page) gets BOTH —
/// text layer for the rich page, OCR for the empty one.
#[test]
fn mixed_pdf_routes_per_page() {
    let Ok(mut engine) = office_parser::IndraNative::load() else {
        eprintln!(
            "skipping: Indra models unavailable ({})",
            office_parser::IndraNative::load().err().unwrap_or_default()
        );
        return;
    };
    let bytes = std::fs::read("tests/fixtures/pdf_mixed.pdf").expect("fixture");
    let doc = engine.parse_document(&bytes, 0.60).expect("mixed parse");

    let texts: Vec<String> = doc
        .blocks
        .iter()
        .filter_map(|b| match b {
            office_parser::document_ast::Block::Paragraph { text, .. } => Some(text.clone()),
            office_parser::document_ast::Block::Heading { text, .. } => Some(text.clone()),
            _ => None,
        })
        .collect();
    let joined = texts.join("\n");
    // text-layer page content survived…
    assert!(joined.contains("Quarterly Field Report"), "got {joined:?}");
    // …AND the scanned page was OCR'd (its content is the same document, rasterized)
    assert!(joined.contains("eleven hours"), "scanned page missing; got {joined:?}");

    let ocr = doc.metadata.extra.get("ocr").expect("ocr metadata");
    assert_eq!(ocr["pages_ocr"], 1, "exactly one page routed to OCR");
    assert_eq!(doc.metadata.pdf_text_quality, Some(PdfTextQuality::TextSparse));
}

/// Table reconstruction: a scanned table page comes back as a real Table
/// block with rows and columns (not flattened paragraph text).
#[test]
fn native_table_reconstruction() {
    let Ok(mut engine) = office_parser::IndraNative::load() else {
        eprintln!(
            "skipping: Indra models unavailable ({})",
            office_parser::IndraNative::load().err().unwrap_or_default()
        );
        return;
    };
    let bytes = std::fs::read("tests/fixtures/pdf_table_scan.pdf").expect("fixture");
    let doc = engine.parse_document(&bytes, 0.60).expect("table ocr");
    use office_parser::document_ast::Block;

    let tables: Vec<&Vec<Vec<office_parser::document_ast::Cell>>> = doc
        .blocks
        .iter()
        .filter_map(|b| match b {
            Block::Table { rows, .. } => Some(rows),
            _ => None,
        })
        .collect();
    assert!(!tables.is_empty(), "no Table block; blocks: {:#?}", doc.blocks.iter().map(|b| match b {
        Block::Table { rows, .. } => format!("Table {}x{}", rows.len(), rows.first().map_or(0, |r| r.len())),
        Block::Paragraph { text, .. } => format!("Para({text})"),
        Block::Heading { text, .. } => format!("Head({text})"),
        _ => "other".into(),
    }).collect::<Vec<_>>());
    let rows = tables[0];
    assert!(rows.len() >= 3, "expected >=3 rows, got {rows:?}");
    let flat: Vec<String> = rows
        .iter()
        .flat_map(|r| r.iter().map(|c| c.text.clone()))
        .collect();
    let joined = flat.join(" ");
    for expected in ["DP Sensor", "12", "49.90", "Temp Probe", "Valve Kit", "120.00"] {
        assert!(joined.contains(expected), "missing {expected:?} in {joined:?}");
    }
    // Qty digits land in a numeric cell (column index may shift ±1 when the
    // renderer's sharpness moves a word-gap across the cell-gap threshold —
    // best-effort alignment; exact grids are the vision-tier's job)
    assert!(
        rows.iter()
            .skip(1)
            .all(|r| r.iter().any(|c| c.text.chars().all(|ch| ch.is_ascii_digit())
                && !c.text.is_empty())),
        "Qty digits missing: {rows:?}"
    );
}

/// Session B: orient.onnx turns a sideways scan upright before OCR.
#[test]
fn native_orientation_recovers_rotated_scan() {
    let Ok(mut engine) = office_parser::IndraNative::load() else {
        eprintln!("skipping: Indra models unavailable");
        return;
    };
    let bytes = std::fs::read("tests/fixtures/pdf_rot90_scan.pdf").expect("fixture");
    let doc = engine.parse_document(&bytes, 0.60).expect("rot90 ocr");
    use office_parser::document_ast::Block;
    let joined = doc
        .blocks
        .iter()
        .map(|b| match b {
            Block::Heading { text, .. } | Block::Paragraph { text, .. } => text.clone(),
            _ => String::new(),
        })
        .collect::<Vec<_>>()
        .join(" ");
    assert!(
        joined.contains("Quarterly Field Report"),
        "rotated scan not recovered: {joined:?}"
    );
}

/// Session C: the script lane — CJK lines route through rec_cjk.onnx and
/// come back as Chinese text, not mojibake.
#[test]
fn native_cjk_script_lane() {
    let Ok(mut engine) = office_parser::IndraNative::load() else {
        eprintln!("skipping: Indra models unavailable");
        return;
    };
    let bytes = std::fs::read("tests/fixtures/pdf_cjk_scan.pdf").expect("fixture");
    let doc = engine.parse_document(&bytes, 0.30).expect("cjk ocr");
    let joined: String = doc
        .blocks
        .iter()
        .map(|b| match b {
            office_parser::document_ast::Block::Paragraph { text, .. }
            | office_parser::document_ast::Block::Heading { text, .. } => text.clone(),
            _ => String::new(),
        })
        .collect::<Vec<_>>()
        .join(" ");
    let cjk_chars = joined
        .chars()
        .filter(|c| ('\u{4E00}'..='\u{9FFF}').contains(c))
        .count();
    assert!(
        cjk_chars >= 6,
        "expected CJK text recovered, got {cjk_chars} hanzi in {joined:?}"
    );
}
