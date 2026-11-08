//! TASK 5 (plan 20261108-001): `PageTextQuality` / `PageVerdict` types and
//! their JSON shape through `render::to_json_value`.

use office_parser::{Document, DocumentMetadata, Format};
use office_parser::quality::{PageTextQuality, PageVerdict};
use office_parser::render::to_json_value;

fn doc_with_page_quality(page_quality: Option<Vec<PageTextQuality>>) -> Document {
    Document {
        blocks: vec![],
        images: vec![],
        metadata: DocumentMetadata {
            format: Format::Pdf,
            title: None,
            page_count: None,
            slide_count: None,
            pdf_text_quality: None,
            page_quality,
            extra: serde_json::json!({}),
        },
    }
}

#[test]
fn test_page_quality_is_none_for_docx() {
    let bytes = include_bytes!("fixtures/docx_headings_lists_tables_images.docx");
    let doc = office_parser::docx::parse(bytes).expect("parse docx");

    assert_eq!(doc.metadata.page_quality, None);
    assert_eq!(doc.metadata.pdf_text_quality, None);

    // The renderer nulls the field when absent (same convention as
    // page_count / slide_count) and must never invent per-page data.
    let v = to_json_value(&doc);
    assert_eq!(v["metadata"]["page_quality"], serde_json::Value::Null);
    assert_eq!(v["metadata"]["pdf_text_quality"], serde_json::Value::Null);
}

#[test]
fn test_page_quality_serde_roundtrip() {
    let pages = vec![
        PageTextQuality {
            page: 1,
            text_richness: 310,
            columns: 1,
            rotated: false,
            verdict: PageVerdict::Trust,
        },
        PageTextQuality {
            page: 2,
            text_richness: 0,
            columns: 1,
            rotated: false,
            verdict: PageVerdict::Ocr {
                reason: "image-only page".to_string(),
            },
        },
        PageTextQuality {
            page: 3,
            text_richness: 21,
            columns: 2,
            rotated: true,
            verdict: PageVerdict::Escalate {
                complexity: Some("column gutter".to_string()),
            },
        },
    ];
    let doc = doc_with_page_quality(Some(pages.clone()));

    let v = to_json_value(&doc);
    let rendered = v["metadata"]["page_quality"]
        .as_array()
        .expect("page_quality renders as an array");
    assert_eq!(rendered.len(), 3);
    assert_eq!(rendered[0]["verdict"]["kind"], "trust");
    assert_eq!(rendered[1]["verdict"]["kind"], "ocr");
    assert_eq!(rendered[2]["verdict"]["kind"], "escalate");
    assert_eq!(rendered[2]["columns"], 2);
    assert_eq!(rendered[2]["rotated"], true);

    let back: Vec<PageTextQuality> =
        serde_json::from_value(v["metadata"]["page_quality"].clone())
            .expect("page_quality deserialises");
    assert_eq!(back, pages);
}

#[test]
fn test_escalate_and_ocr_are_distinct_in_json() {
    let ocr = serde_json::to_value(PageVerdict::Ocr {
        reason: "scan".to_string(),
    })
    .expect("serialise ocr");
    let escalate =
        serde_json::to_value(PageVerdict::Escalate { complexity: None }).expect("serialise escalate");

    // Internally tagged, self-describing: the discriminators differ and each
    // variant carries only its own fields.
    assert_eq!(ocr["kind"], "ocr");
    assert_eq!(escalate["kind"], "escalate");
    assert!(ocr.get("complexity").is_none());
    assert!(escalate.get("reason").is_none());
    assert_eq!(ocr["reason"], "scan");
    assert_eq!(escalate["complexity"], serde_json::Value::Null);
}

#[test]
fn test_pdf_text_quality_variants_unchanged() {
    let quality = office_parser::PdfTextQuality::TextSparse;
    // Exhaustive match: the three pre-existing variants must keep compiling
    // with their original meaning (regression lock for 0.2.0 consumers).
    match quality {
        office_parser::PdfTextQuality::TextRich => {}
        office_parser::PdfTextQuality::TextSparse => {}
        office_parser::PdfTextQuality::ImageOnly => {}
    }
}
