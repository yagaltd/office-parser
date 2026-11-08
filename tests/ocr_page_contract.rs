#![cfg(all(feature = "pdf", feature = "ocr-indra-native"))]
//! TASK 10 (plan 20261108-001): the `OcrPage` + defaulted
//! `PageOcrEngine::ocr_page()` contract.
//!
//! Contract: `.workflows/specs/task-10.spec` — the four scenarios:
//! 1. a legacy engine (only `ocr_page_png`) compiles and runs through the
//!    defaulted `ocr_page`; the three optional fields come back `None` and
//!    blocks/confidence match `ocr_page_png`'s output;
//! 2. the trait stays object-safe: `&mut dyn PageOcrEngine` + `ocr_page()`
//!    compiles and runs (a `where Self: Sized` bound would make it not);
//! 3. engine-supplied optional fields thread into `metadata.extra["ocr"]`,
//!    and engines returning `None` leave those keys absent, never null;
//! 4. `tests/pdf_ocr.rs` (6 tests) stays green and byte-identical — covered
//!    by the suite run + `git diff --name-only tests/pdf_ocr.rs`, not a
//!    test in this file.

use office_parser::document_ast::{Block, SourceSpan};
use office_parser::formats::pdf::{
    parse_pdf_split, OcrLayoutRegion, OcrPage, OcrWordBox, PageOcrEngine,
};

const SCANNED: &str = "tests/fixtures/pdf_scanned.pdf";

fn one_paragraph(text: &str) -> Vec<Block> {
    vec![Block::Paragraph {
        block_index: 0,
        text: text.to_string(),
        marks: Vec::new(),
        source: SourceSpan::default(),
    }]
}

/// Scenario 1: a legacy implementor needs no change — the defaulted
/// `ocr_page` wraps `ocr_page_png` and returns `None` for the rich fields.
#[test]
fn test_legacy_engine_implementing_only_ocr_page_png_still_works() {
    struct LegacyEngine;
    impl PageOcrEngine for LegacyEngine {
        fn ocr_page_png(&mut self, _png: &[u8]) -> Result<(Vec<Block>, Option<f64>), String> {
            Ok((one_paragraph("legacy block"), Some(0.75)))
        }
        fn name(&self) -> &'static str {
            "legacy"
        }
    }

    let (png_blocks, png_conf) = LegacyEngine.ocr_page_png(b"\x89PNG-fake").expect("ocr_page_png");

    let mut engine = LegacyEngine;
    let dyn_engine: &mut dyn PageOcrEngine = &mut engine;
    let page = dyn_engine
        .ocr_page(b"\x89PNG-fake")
        .expect("defaulted ocr_page wraps ocr_page_png");

    assert_eq!(page.blocks, png_blocks, "blocks match ocr_page_png");
    assert_eq!(page.confidence, png_conf, "confidence matches ocr_page_png");
    assert!(page.layout_regions.is_none(), "legacy engine: no layout regions");
    assert!(page.word_boxes.is_none(), "legacy engine: no word boxes");
    assert!(page.table_html.is_none(), "legacy engine: no table html");
}

/// Scenario 2: object safety — the defaulted method carries no `where
/// Self: Sized` escape hatch, so it is callable through `&mut dyn
/// PageOcrEngine`. Compilation of this test is the proof; the run pins it.
#[test]
fn test_page_ocr_engine_is_object_safe() {
    struct MinimalEngine;
    impl PageOcrEngine for MinimalEngine {
        fn ocr_page_png(&mut self, _png: &[u8]) -> Result<(Vec<Block>, Option<f64>), String> {
            Ok((Vec::new(), None))
        }
        fn name(&self) -> &'static str {
            "minimal"
        }
    }

    fn ocr_page_through_dyn(engine: &mut dyn PageOcrEngine) -> OcrPage {
        engine.ocr_page(b"\x89PNG-fake").expect("ocr_page on dyn")
    }

    let mut engine = MinimalEngine;
    let page = ocr_page_through_dyn(&mut engine);
    assert!(page.blocks.is_empty());
    assert!(page.confidence.is_none());
    assert!(page.layout_regions.is_none());
    assert!(page.word_boxes.is_none());
    assert!(page.table_html.is_none());
}

/// Scenario 3: an overriding engine threads its rich fields into
/// `metadata.extra["ocr"]`; a legacy engine leaves the keys absent (null
/// would break the absent-not-null metadata discipline).
#[test]
fn test_optional_fields_threaded_into_ocr_metadata() {
    struct RichEngine;
    impl PageOcrEngine for RichEngine {
        fn ocr_page_png(&mut self, _png: &[u8]) -> Result<(Vec<Block>, Option<f64>), String> {
            Ok((one_paragraph("ocr text"), Some(0.9)))
        }
        fn ocr_page(&mut self, png: &[u8]) -> Result<OcrPage, String> {
            let (blocks, confidence) = self.ocr_page_png(png)?;
            Ok(OcrPage {
                blocks,
                confidence,
                layout_regions: Some(vec![OcrLayoutRegion {
                    label: "Table".to_string(),
                    x0: 40.0,
                    y0: 60.0,
                    x1: 560.0,
                    y1: 300.0,
                }]),
                word_boxes: Some(vec![OcrWordBox {
                    text: "Quarterly".to_string(),
                    x0: 40.0,
                    y0: 62.0,
                    x1: 120.0,
                    y1: 80.0,
                }]),
                table_html: Some("<table><tr><td>Q1</td></tr></table>".to_string()),
            })
        }
        fn name(&self) -> &'static str {
            "rich-test"
        }
    }

    struct NoneEngine;
    impl PageOcrEngine for NoneEngine {
        fn ocr_page_png(&mut self, _png: &[u8]) -> Result<(Vec<Block>, Option<f64>), String> {
            Ok((Vec::new(), None))
        }
        fn name(&self) -> &'static str {
            "none-test"
        }
    }

    if !pdftoppm_available() {
        eprintln!("skipping: pdftoppm unavailable (routing path needs a rasterizer)");
        return;
    }
    let bytes = std::fs::read(SCANNED).expect("fixture");

    // Positive path: both scanned pages forced through the engine; every
    // optional field reaches the ocr metadata mirror, one record per page.
    let mut rich = RichEngine;
    let engine: &mut dyn PageOcrEngine = &mut rich;
    let (doc, _pages) = parse_pdf_split(&bytes, Some(engine), 150, 50, true).expect("split parse");
    let ocr = doc.metadata.extra.get("ocr").expect("ocr metadata present");
    assert_eq!(ocr["engine"], "rich-test");
    assert_eq!(ocr["pages_ocr"], 2, "both scanned pages routed to the engine");

    let table_html = ocr["table_html"].as_array().expect("table_html threaded");
    assert_eq!(table_html.len(), 2, "one record per OCR'd page");
    assert_eq!(table_html[0]["page"], serde_json::json!(1));
    assert_eq!(
        table_html[0]["html"],
        "<table><tr><td>Q1</td></tr></table>"
    );

    let word_boxes = ocr["word_boxes"].as_array().expect("word_boxes threaded");
    assert_eq!(word_boxes[0]["page"], serde_json::json!(1));
    assert_eq!(word_boxes[0]["words"][0]["text"], "Quarterly");
    assert_eq!(word_boxes[0]["words"][0]["x0"].as_f64(), Some(40.0));

    let regions = ocr["layout_regions"]
        .as_array()
        .expect("layout_regions threaded");
    assert_eq!(regions[0]["page"], serde_json::json!(1));
    assert_eq!(regions[0]["regions"][0]["label"], "Table");
    assert_eq!(regions[0]["regions"][0]["y1"].as_f64(), Some(300.0));

    // Negative path: the default wrapper yields None for all three fields,
    // so the metadata keys must be ABSENT — never null.
    let mut none = NoneEngine;
    let engine: &mut dyn PageOcrEngine = &mut none;
    let (doc, _pages) = parse_pdf_split(&bytes, Some(engine), 150, 50, true).expect("split parse");
    let ocr = doc.metadata.extra.get("ocr").expect("ocr metadata present");
    assert_eq!(ocr["pages_ocr"], 2);
    assert!(ocr.get("table_html").is_none(), "table_html absent, not null");
    assert!(ocr.get("word_boxes").is_none(), "word_boxes absent, not null");
    assert!(
        ocr.get("layout_regions").is_none(),
        "layout_regions absent, not null"
    );
}

fn pdftoppm_available() -> bool {
    std::env::var_os("PATH")
        .map(|p| {
            std::env::split_paths(&p)
                .any(|dir| dir.join("pdftoppm").is_file() || dir.join("pdftoppm.exe").is_file())
        })
        .unwrap_or(false)
}
