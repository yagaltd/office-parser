#![cfg(feature = "pdf")]
//! TASK 8 (plan 20261108-001): verdict-driven routing, the worst-page-wins
//! rollup, and the `metadata.extra["pdf_quality"]` mirror.
//!
//! Contract: `.workflows/specs/task-8.spec`. TASK 7 shipped verdict
//! population (`metadata.page_quality`); this file locks the remainder:
//! the router consumes verdicts — `Escalate` NEVER adds a page to `ocr_pages`
//! (it is a vision-LLM signal, because OCR cannot repair reading order) —
//! the rollup routes the worst page while the document-level
//! `PdfTextQuality` enum stays the plain summary, the per-page array is
//! mirrored into `metadata.extra["pdf_quality"]`, and `parse_pdf_split` /
//! `parse_with_page_ocr_opts` report identical per-page verdicts.

use std::cell::Cell;

use office_parser::document_ast::{Block, SourceSpan};
use office_parser::formats::pdf::{
    parse_pdf_split, parse_with_page_ocr_opts, routing_rollup, PageOcrEngine,
    PAGE_TEXT_THRESHOLD,
};
use office_parser::quality::{PageTextQuality, PageVerdict};

const MIXED: &str = "tests/fixtures/pdf_mixed.pdf";

/// 65 alphanumeric chars per sentence; repeated 5x = 325, comfortably above
/// the 50-char gate (`PAGE_TEXT_THRESHOLD`) and ~300 chars as the task asks.
const BODY_SENTENCE: &str =
    "The quality router reads the text layer of every page before any merge happens. ";

/// One-page PDF with a rich text layer and an optional native `/Rotate` —
/// TASK 4's `synth_pdf` pattern (lopdf, already the `pdf`-feature
/// dependency; no new dependency, no pdfium requirement).
fn synth_pdf(rotate: Option<i64>) -> Vec<u8> {
    use lopdf::{Dictionary, Document, Object, Stream};

    let nm = |s: &str| Object::Name(s.as_bytes().to_vec());
    let int = |v: i64| Object::Integer(v);
    let real = |v: f32| Object::Real(v);

    let mut doc = Document::new();

    let content = BODY_SENTENCE.repeat(5);
    let content_id = doc.add_object(Stream::new(
        Dictionary::new(),
        format!("BT /F1 12 Tf 72 700 Td ({content}) Tj ET").into_bytes(),
    ));

    let mut font = Dictionary::new();
    font.set("Type", nm("Font"));
    font.set("Subtype", nm("Type1"));
    font.set("BaseFont", nm("Helvetica"));
    let font_id = doc.add_object(font);

    let mut resources = Dictionary::new();
    let mut font_map = Dictionary::new();
    font_map.set("F1", Object::Reference(font_id));
    resources.set("Font", Object::Dictionary(font_map));

    let mut pages = Dictionary::new();
    pages.set("Type", nm("Pages"));
    let pages_id = doc.add_object(pages);

    let mut page = Dictionary::new();
    page.set("Type", nm("Page"));
    page.set("Parent", Object::Reference(pages_id));
    page.set(
        "MediaBox",
        Object::Array(vec![int(0), int(0), real(612.0), real(792.0)]),
    );
    page.set("Resources", Object::Dictionary(resources));
    page.set("Contents", Object::Reference(content_id));
    if let Some(rot) = rotate {
        page.set("Rotate", int(rot));
    }
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

/// Engine that counts every call and would replace a routed page with an
/// unmistakable "ocr echo" block — a routed page cannot hide.
#[derive(Default)]
struct CountingEngine {
    calls: Cell<usize>,
}

impl PageOcrEngine for CountingEngine {
    fn ocr_page_png(&mut self, _png: &[u8]) -> Result<(Vec<Block>, Option<f64>), String> {
        self.calls.set(self.calls.get() + 1);
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
        "counting"
    }
}

/// Plain text of a block, for containment assertions.
fn block_text(b: &Block) -> String {
    match b {
        Block::Heading { text, .. } | Block::Paragraph { text, .. } | Block::Note { text, .. } => {
            text.clone()
        }
        Block::List { items, .. } => items
            .iter()
            .map(|i| i.text.as_str())
            .collect::<Vec<_>>()
            .join(" "),
        Block::Table { rows, .. } => rows
            .iter()
            .flat_map(|r| r.iter())
            .map(|c| c.text.as_str())
            .collect::<Vec<_>>()
            .join(" "),
        Block::Image { alt, .. } | Block::Link { text: alt, .. } => alt.clone().unwrap_or_default(),
    }
}

/// `metadata.extra["ocr"]`'s OCR'd-page count (0 when the key is absent —
/// the parser only writes it once a page actually routed).
fn pages_ocr_of(doc: &office_parser::Document) -> u64 {
    doc.metadata
        .extra
        .get("ocr")
        .and_then(|o| o.get("pages_ocr"))
        .and_then(|p| p.as_u64())
        .unwrap_or(0)
}

fn pdftoppm_available() -> bool {
    std::env::var_os("PATH")
        .map(|p| {
            std::env::split_paths(&p)
                .any(|dir| dir.join("pdftoppm").is_file() || dir.join("pdftoppm.exe").is_file())
        })
        .unwrap_or(false)
}

/// Renderer detection for the merge-path tests (render_pages_png_opts order:
/// pdfium first when the feature + library are present, then poppler).
#[cfg(feature = "render-pdfium")]
fn pdfium_renderer_available() -> bool {
    office_parser::formats::pdf::pdfium_lib_path().is_some()
}

#[cfg(not(feature = "render-pdfium"))]
fn pdfium_renderer_available() -> bool {
    false
}

fn renderer_available() -> bool {
    pdftoppm_available() || pdfium_renderer_available()
}

/// Scenario: one quality entry per page (spec).
#[test]
fn test_page_quality_len_matches_page_count() {
    let bytes = std::fs::read(MIXED).expect("fixture");
    let doc = office_parser::pdf::parse(&bytes).expect("parses");
    let pq = doc.metadata.page_quality.as_ref().expect("page_quality");
    assert_eq!(pq.len(), doc.metadata.page_count.expect("page count"));
    let numbers: Vec<u32> = pq.iter().map(|q| q.page).collect();
    let expected: Vec<u32> = (1..=pq.len() as u32).collect();
    assert_eq!(numbers, expected, "pages are 1..=count in order");
}

/// Scenario: an Escalate page is never sent to OCR (spec) — the dedicated
/// negative test. A rich (~325 chars >= gate) ROTATED text-layer page must
/// keep its text layer: the verdict is Escalate, the engine is never called,
/// and pages_ocr does not count it. Escalate is a vision-LLM signal only —
/// OCR cannot repair reading order.
#[test]
fn test_escalate_page_is_not_routed_to_ocr() {
    let bytes = synth_pdf(Some(90));
    let plain = office_parser::pdf::parse(&bytes).expect("plain parse");

    // Text-layer truth: rich, rotated, escalates.
    let pq = plain.metadata.page_quality.as_ref().expect("page_quality");
    assert_eq!(pq.len(), 1);
    assert_eq!(pq[0].page, 1);
    assert!(pq[0].rotated, "/Rotate 90 must surface as rotated");
    assert!(
        pq[0].text_richness >= 300,
        "rich layer above the task's ~300 chars, got {}",
        pq[0].text_richness
    );
    assert_eq!(
        pq[0].verdict,
        PageVerdict::Escalate {
            complexity: Some("rotated page".to_string())
        }
    );

    // Router, WITHOUT force: the Escalate page stays out of ocr_pages.
    let mut engine = CountingEngine::default();
    let doc = parse_with_page_ocr_opts(&bytes, &mut engine, 150, PAGE_TEXT_THRESHOLD, false)
        .expect("routed parse");
    assert_eq!(engine.calls.get(), 0, "the engine must never be called");
    // The text layer stayed: identical blocks to the plain text-layer parse
    // (an OCR merge would have replaced them with the engine's "ocr echo").
    assert_eq!(doc.blocks, plain.blocks, "text layer survived un-OCR'd");
    assert_eq!(
        pages_ocr_of(&doc),
        0,
        "pages_ocr must not count an Escalate page"
    );
    // The verdict still describes the pre-routing text layer.
    assert_eq!(doc.metadata.page_quality.as_ref(), Some(pq));
}

/// Scenario: a trusted simple page stays Trust (spec negative). Single
/// column, text-rich, unrotated — no OCR reason, no escalation.
#[test]
fn test_trusted_simple_page_has_no_ocr_reason() {
    let bytes = synth_pdf(None);
    let doc = office_parser::pdf::parse(&bytes).expect("plain parse");
    let pq = doc.metadata.page_quality.as_ref().expect("page_quality");
    assert_eq!(pq.len(), 1);
    assert_eq!(
        pq[0].verdict,
        PageVerdict::Trust,
        "no OCR reason, no escalation"
    );
    assert_eq!(pq[0].columns, 1, "single column");
    assert!(!pq[0].rotated);
    assert!(pq[0].text_richness >= 50, "text-rich above the gate");

    // And the router agrees: a Trust page is never routed either.
    let mut engine = CountingEngine::default();
    let routed = parse_with_page_ocr_opts(&bytes, &mut engine, 150, PAGE_TEXT_THRESHOLD, false)
        .expect("routed parse");
    assert_eq!(engine.calls.get(), 0);
    assert_eq!(routed.blocks, doc.blocks, "text layer survived un-OCR'd");
    assert_eq!(pages_ocr_of(&routed), 0);
}

/// Scenario: split-pages and combined parsing agree (spec).
#[test]
fn test_split_pages_reports_same_verdicts() {
    // Always-on parity: same input, both entry points, identical verdicts.
    let bytes = synth_pdf(None);
    let mut a = CountingEngine::default();
    let a_engine: &mut dyn PageOcrEngine = &mut a;
    let (split_combined, split_pages) =
        parse_pdf_split(&bytes, Some(a_engine), 150, PAGE_TEXT_THRESHOLD, false)
            .expect("split parse");
    let mut b = CountingEngine::default();
    let opts_combined = parse_with_page_ocr_opts(&bytes, &mut b, 150, PAGE_TEXT_THRESHOLD, false)
        .expect("opts parse");
    assert_eq!(a.calls.get(), 0);
    assert_eq!(b.calls.get(), 0);
    assert_eq!(
        split_combined.metadata.page_quality,
        opts_combined.metadata.page_quality,
        "combined verdicts identical"
    );
    assert_eq!(split_pages.len(), 1);
    assert_eq!(
        split_pages[0].1.metadata.page_quality,
        opts_combined.metadata.page_quality,
        "the single page doc carries the same verdict"
    );

    // Fixture parity: pdf_mixed through both paths with the engine attached
    // (page 2 routes, so this needs a renderer).
    if !renderer_available() {
        eprintln!("skipping pdf_mixed merge-path parity: no PDF renderer (pdftoppm/pdfium)");
        return;
    }
    let bytes = std::fs::read(MIXED).expect("fixture");
    let mut split_engine = CountingEngine::default();
    let split_dyn: &mut dyn PageOcrEngine = &mut split_engine;
    let (split_combined, split_pages) =
        parse_pdf_split(&bytes, Some(split_dyn), 150, PAGE_TEXT_THRESHOLD, false)
            .expect("split parse");
    let mut opts_engine = CountingEngine::default();
    let opts_combined = parse_with_page_ocr_opts(&bytes, &mut opts_engine, 150, PAGE_TEXT_THRESHOLD, false)
        .expect("opts parse");
    assert_eq!(split_engine.calls.get(), 1, "page 2 (image-only) routed");
    assert_eq!(opts_engine.calls.get(), 1, "page 2 (image-only) routed");
    let a = split_combined
        .metadata
        .page_quality
        .as_ref()
        .expect("split quality");
    let b = opts_combined
        .metadata
        .page_quality
        .as_ref()
        .expect("opts quality");
    assert_eq!(a, b, "per-page verdicts identical element for element");
    assert_eq!(a.len(), split_pages.len());
    for (i, (_, page_doc)) in split_pages.iter().enumerate() {
        let pq = page_doc
            .metadata
            .page_quality
            .as_ref()
            .expect("page doc quality");
        assert_eq!(pq.len(), 1);
        assert_eq!(&pq[0], &a[i], "page doc carries its own verdict");
    }
    // And they match the text-layer truth from the plain parse.
    let plain = office_parser::pdf::parse(&bytes).expect("plain parse");
    assert_eq!(
        a,
        plain.metadata.page_quality.as_ref().expect("plain quality"),
        "verdicts are the text-layer truth on every path"
    );
}

/// Scenario: extra mirror matches the struct (spec).
#[test]
fn test_pdf_quality_mirrored_in_metadata_extra() {
    let bytes = std::fs::read(MIXED).expect("fixture");

    // Plain combined parse: the mirror describes the same pages with the
    // same verdict kinds as `metadata.page_quality`.
    let doc = office_parser::pdf::parse(&bytes).expect("parses");
    let pq = doc.metadata.page_quality.as_ref().expect("page_quality");
    let mirror = doc.metadata.extra["pdf_quality"]
        .as_array()
        .expect("mirror array");
    assert_eq!(mirror.len(), pq.len());
    for (m, q) in mirror.iter().zip(pq) {
        assert_eq!(m["page"], serde_json::json!(q.page));
        assert_eq!(m["text_richness"], serde_json::json!(q.text_richness));
        assert_eq!(m["columns"], serde_json::json!(q.columns));
        assert_eq!(m["rotated"], serde_json::json!(q.rotated));
        assert_eq!(m["verdict"], serde_json::to_value(&q.verdict).unwrap());
    }

    // Split path: the combined doc mirrors the full array, each page doc
    // mirrors exactly its own single record.
    let (combined, pages) =
        parse_pdf_split(&bytes, None, 150, PAGE_TEXT_THRESHOLD, false).expect("split parse");
    let cq = combined
        .metadata
        .page_quality
        .as_ref()
        .expect("combined quality");
    let cm = combined.metadata.extra["pdf_quality"]
        .as_array()
        .expect("combined mirror");
    assert_eq!(cm.len(), cq.len());
    for (m, q) in cm.iter().zip(cq) {
        assert_eq!(m["verdict"], serde_json::to_value(&q.verdict).unwrap());
    }
    for (i, (_, page_doc)) in pages.iter().enumerate() {
        let s = page_doc
            .metadata
            .page_quality
            .as_ref()
            .expect("page doc quality");
        assert_eq!(s.len(), 1);
        let m = page_doc.metadata.extra["pdf_quality"]
            .as_array()
            .expect("page mirror");
        assert_eq!(m.len(), 1);
        assert_eq!(m[0]["page"], serde_json::json!(s[0].page));
        assert_eq!(m[0]["verdict"], serde_json::to_value(&s[0].verdict).unwrap());
        assert_eq!(&s[0], &cq[i], "page doc record matches combined position");
    }
}

/// Scenario: document rollup prefers the worst page (spec) — worst-page-wins
/// routing while the public document-level enum stays the summary.
#[test]
fn test_document_rollup_uses_worst_page_for_routing() {
    let mk = |page: u32, verdict: PageVerdict| PageTextQuality {
        page,
        text_richness: 100,
        columns: 1,
        rotated: false,
        verdict,
    };

    // Page 1 TextRich/Trust, page 2 ImageOnly/Ocr: the rollup routes page 2.
    let pages = vec![
        mk(1, PageVerdict::Trust),
        mk(
            2,
            PageVerdict::Ocr {
                reason: "image-only page".to_string(),
            },
        ),
    ];
    assert_eq!(routing_rollup(&pages), vec![1usize]);

    // Every Ocr page routes (worst-page-wins, not doc-majority).
    assert_eq!(
        routing_rollup(&[
            mk(
                1,
                PageVerdict::Ocr {
                    reason: "sparse text layer (16 chars < 50)".to_string()
                }
            ),
            mk(2, PageVerdict::Trust),
            mk(
                3,
                PageVerdict::Ocr {
                    reason: "image-only page".to_string()
                }
            ),
        ]),
        vec![0usize, 2usize]
    );

    // Escalate NEVER adds a page — vision-LLM signal only.
    assert!(routing_rollup(&[
        mk(1, PageVerdict::Trust),
        mk(
            2,
            PageVerdict::Escalate {
                complexity: Some("rotated page".to_string())
            }
        ),
    ])
    .is_empty());

    // End-to-end: pdf_mixed's doc-level value stays whatever its text layer
    // says, yet the image-only page 2 still routes.
    if !renderer_available() {
        eprintln!("skipping end-to-end rollup: no PDF renderer (pdftoppm/pdfium)");
        return;
    }
    let bytes = std::fs::read(MIXED).expect("fixture");
    let plain = office_parser::pdf::parse(&bytes).expect("plain parse");
    let mut engine = CountingEngine::default();
    let doc = parse_with_page_ocr_opts(&bytes, &mut engine, 150, PAGE_TEXT_THRESHOLD, false)
        .expect("routed parse");
    // The public pdf_text_quality enum keeps its document-level value.
    assert_eq!(
        doc.metadata.pdf_text_quality,
        plain.metadata.pdf_text_quality,
        "doc-level enum is the summary, unchanged by routing"
    );
    assert_eq!(engine.calls.get(), 1, "only the worst page routed");
    let ocr = doc.metadata.extra.get("ocr").expect("ocr metadata");
    assert_eq!(ocr["pages_ocr"], serde_json::json!(1));
    assert!(ocr["low_confidence_pages"].is_array());
    // Page 2 came back from the engine…
    assert!(
        doc.blocks
            .iter()
            .any(|b| block_text(b).contains("ocr echo")),
        "the image-only page went through the engine"
    );
    // …while page 1's text layer stayed, block for block.
    let ranges = plain.metadata.extra["page_block_ranges"]
        .as_array()
        .expect("ranges");
    let p1_start = ranges[0]["block_start"].as_u64().unwrap() as usize;
    let p1_end = ranges[0]["block_end"].as_u64().unwrap() as usize;
    for b in &plain.blocks[p1_start..p1_end] {
        let t = block_text(b);
        if t.is_empty() {
            continue;
        }
        assert!(
            doc.blocks.iter().any(|db| block_text(db) == t),
            "text-layer block survived: {t:?}"
        );
    }
    // Verdicts describe the text layer, not the merged result.
    assert_eq!(doc.metadata.page_quality, plain.metadata.page_quality);
}
