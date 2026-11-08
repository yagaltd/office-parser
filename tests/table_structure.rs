#![cfg(all(feature = "pdf", feature = "ocr-indra-native", feature = "ocr-table-structure"))]
//! TASK 13 (own-table-structure): the contract transfers from
//! tests/ocr_rusto_tables.rs (deleted with vendored rusto-core).
//!
//! Scenario split (spec Acceptance Criteria):
//! - test_transcripts_identical_with_model_on_and_off — transcripts +
//!   confidence identical with the SLANet model pinned vs hidden
//!   (pdf_scanned.pdf, both arms in one binary via env toggling).
//! - test_no_model_feature_off_payload_shape — on a fixture without tables,
//!   feature-on + no-model produces the exact feature-off payload shape
//!   (no rich keys at all).
//! - test_geometry_passthrough_and_no_html_without_model — the engine's
//!   geometry (layout_regions/word_boxes) threads regardless of the model;
//!   table_html stays absent without it.
//! - test_table_html_populated_for_table_scan_fixture — deterministic stub
//!   engine + real model ⇒ table_html present and structure-correct.
//! - test_real_engine_produces_table_html_for_table_scan — IndraNative +
//!   pdf_table_scan + real model (the gate's validated scenario; skips
//!   loudly without the model).

use office_parser::document_ast::{Block, Cell, SourceSpan};
use office_parser::formats::pdf::{parse_pdf_split, OcrLayoutRegion, OcrPage, OcrWordBox, PageOcrEngine};
use office_parser::formats::table_structure::TABLE_MODEL_ENV;

const SCAN: &str = "tests/fixtures/pdf_scanned.pdf";
const TABLE_SCAN: &str = "tests/fixtures/pdf_table_scan.pdf";

/// Env toggling must not race across tests: model-path tests take this lock
/// for their whole scope. Field order drops env restore before mutex release.
struct EnvScope {
    _mutex: std::sync::MutexGuard<'static, ()>,
    prior: Option<std::ffi::OsString>,
}

static ENV_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

impl EnvScope {
    fn acquire() -> Self {
        Self {
            _mutex: ENV_LOCK.lock().unwrap_or_else(|e| e.into_inner()),
            prior: std::env::var_os(TABLE_MODEL_ENV),
        }
    }
}

impl Drop for EnvScope {
    fn drop(&mut self) {
        match self.prior.take() {
            Some(v) => unsafe { std::env::set_var(TABLE_MODEL_ENV, v) },
            None => unsafe { std::env::remove_var(TABLE_MODEL_ENV) },
        }
    }
}

fn pin_model_env(model: &std::path::Path) {
    // SAFETY: tests in this binary touching OFFICE_PARSER_TABLE_MODEL hold
    // ENV_LOCK for their whole scope (set_var/remove_var are unsafe in
    // edition 2024).
    unsafe { std::env::set_var(TABLE_MODEL_ENV, model) };
}

fn hide_model_env() {
    unsafe { std::env::set_var(TABLE_MODEL_ENV, "/nonexistent/office-parser-slanet.onnx") };
}

fn pdftoppm_available() -> bool {
    std::env::var_os("PATH").map_or(false, |p| {
        std::env::split_paths(&p)
            .any(|dir| dir.join("pdftoppm").is_file() || dir.join("pdftoppm.exe").is_file())
    })
}

fn table_model_path() -> Option<std::path::PathBuf> {
    office_parser::formats::table_structure::table_model_path()
}

/// Flattened block text in order — the transcript.
fn transcript(doc: &office_parser::Document) -> Vec<String> {
    doc.blocks
        .iter()
        .map(|b| match b {
            Block::Heading { text, .. } | Block::Paragraph { text, .. } => text.clone(),
            Block::List { items, .. } => items
                .iter()
                .map(|i| i.text.clone())
                .collect::<Vec<_>>()
                .join("\n"),
            Block::Table { rows, .. } => rows
                .iter()
                .map(|r| {
                    r.iter()
                        .map(|c| c.text.clone())
                        .collect::<Vec<_>>()
                        .join("|")
                })
                .collect::<Vec<_>>()
                .join("\n"),
            _ => String::new(),
        })
        .collect()
}

fn cell(text: &str) -> Cell {
    Cell {
        text: text.to_string(),
        colspan: 1,
        rowspan: 1,
    }
}

/// Deterministic engine exposing Indra-style table geometry (the task-13
/// geometry contract) with cells known a priori; overrides `ocr_page` exactly
/// like indra_native.rs does under the feature. Object-safety scenario: it is
/// used through `&mut dyn PageOcrEngine`.
struct GeometryEngine;
impl PageOcrEngine for GeometryEngine {
    fn ocr_page_png(&mut self, _png: &[u8]) -> Result<(Vec<Block>, Option<f64>), String> {
        Ok((
            vec![Block::Table {
                block_index: 0,
                rows: vec![
                    vec![cell("alpha"), cell("beta")],
                    vec![cell("gamma"), cell("delta")],
                ],
                widths: Vec::new(),
                source: SourceSpan::default(),
            }],
            Some(0.9),
        ))
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
            word_boxes: Some(vec![
                OcrWordBox { text: "alpha".into(), x0: 40.0, y0: 62.0, x1: 280.0, y1: 80.0 },
                OcrWordBox { text: "beta".into(), x0: 300.0, y0: 62.0, x1: 560.0, y1: 80.0 },
                OcrWordBox { text: "gamma".into(), x0: 40.0, y0: 90.0, x1: 280.0, y1: 110.0 },
                OcrWordBox { text: "delta".into(), x0: 300.0, y0: 90.0, x1: 560.0, y1: 110.0 },
            ]),
            table_html: None,
        })
    }

    fn name(&self) -> &'static str {
        "geometry-stub"
    }
}

/// Spec scenario: transcripts identical with the model pinned vs hidden.
/// Both arms run in one binary; the hidden arm fails before any decode
/// (enrichment provably no-op). pdf_scanned has no tables, so this locks the
/// strongest statement: the feature changes NOTHING on table-free pages.
#[test]
fn test_transcripts_identical_with_model_on_and_off() {
    let Some(model) = table_model_path() else {
        eprintln!(
            "skipping: no SLANet-plus model (set {TABLE_MODEL_ENV} to a model file) — \
             transcript identity with the model active needs it"
        );
        return;
    };
    let Ok(mut engine) = office_parser::IndraNative::load() else {
        eprintln!(
            "skipping: Indra models unavailable ({})",
            office_parser::IndraNative::load().err().unwrap_or_default()
        );
        return;
    };
    let bytes = std::fs::read(SCAN).expect("fixture");

    let _scope = EnvScope::acquire();
    pin_model_env(&model);
    let doc_on = engine.parse_document(&bytes, 0.60).expect("parse (model on)");
    hide_model_env();
    let doc_off = engine.parse_document(&bytes, 0.60).expect("parse (model off)");
    drop(_scope);

    assert_eq!(
        transcript(&doc_on),
        transcript(&doc_off),
        "transcript must be identical with the model on/off"
    );
    assert_eq!(doc_on.blocks.len(), doc_off.blocks.len());
    let ocr_on = doc_on.metadata.extra.get("ocr").expect("ocr metadata (on)");
    let ocr_off = doc_off.metadata.extra.get("ocr").expect("ocr metadata (off)");
    assert_eq!(ocr_on["pages_ocr"], ocr_off["pages_ocr"]);
    assert_eq!(
        ocr_on["low_confidence_pages"],
        ocr_off["low_confidence_pages"],
        "confidence-derived rollup identical"
    );
}

/// Feature-off payload shape: table-free pages + no model ⇒ no rich key at
/// all — exactly what a build without the feature produces. Always runnable
/// where Indra + a rasterizer exist; skips loudly otherwise.
#[test]
fn test_no_model_feature_off_payload_shape() {
    if !pdftoppm_available() {
        eprintln!("skipping: pdftoppm unavailable (OCR merge path needs a rasterizer)");
        return;
    }
    let Ok(mut engine) = office_parser::IndraNative::load() else {
        eprintln!(
            "skipping: Indra models unavailable ({})",
            office_parser::IndraNative::load().err().unwrap_or_default()
        );
        return;
    };
    let _scope = EnvScope::acquire();
    hide_model_env();
    let bytes = std::fs::read(SCAN).expect("fixture");
    let doc = engine
        .parse_document(&bytes, 0.60)
        .expect("parse with hidden model");
    drop(_scope);

    let ocr = doc.metadata.extra.get("ocr").expect("ocr metadata present");
    assert_eq!(ocr["pages_ocr"], 2);
    for key in ["layout_regions", "word_boxes", "table_html"] {
        assert!(
            ocr.get(key).is_none(),
            "{key} must stay absent without the model on a table-free page \
             (feature-off payload shape)"
        );
    }
}

/// The engine's geometry threads regardless of the model; the enrichment
/// adds nothing without it. Stub engine (deterministic), object-safety
/// scenario included (`&mut dyn PageOcrEngine`).
#[test]
fn test_geometry_passthrough_and_no_html_without_model() {
    if !pdftoppm_available() {
        eprintln!("skipping: pdftoppm unavailable (OCR merge path needs a rasterizer)");
        return;
    }
    let bytes = std::fs::read(TABLE_SCAN).expect("fixture");
    let _scope = EnvScope::acquire();
    hide_model_env();
    let engine: &mut dyn PageOcrEngine = &mut GeometryEngine;
    let (combined, _pages) =
        parse_pdf_split(&bytes, Some(engine), 150, 50, false).expect("split parse");
    drop(_scope);

    let ocr = combined.metadata.extra.get("ocr").expect("ocr metadata");
    assert!(
        ocr.get("layout_regions").is_some(),
        "engine geometry must thread without the model"
    );
    assert!(
        ocr.get("word_boxes").is_some(),
        "engine cell rects must thread without the model"
    );
    assert!(
        ocr.get("table_html").is_none(),
        "table_html stays absent without the model (documented no-op)"
    );
    // Transcript passthrough at the seam: only block_index is renumbered.
    match &combined.blocks[0] {
        Block::Table { rows, .. } => {
            let texts: Vec<&str> = rows
                .iter()
                .flat_map(|r| r.iter())
                .map(|c| c.text.as_str())
                .collect();
            assert_eq!(texts, ["alpha", "beta", "gamma", "delta"]);
        }
        other => panic!("expected the stub's Table block, got {other:?}"),
    }
}

/// Positive arm — stub geometry engine + real model: table_html must be
/// present and structure-correct (one `<table>` wrapper, td open/close
/// parity). Cell text assertions live in the module unit tests; which tokens
/// SLANet emits over a synthetic raster is the model's business, not this
/// test's. Skips loudly without the model.
#[test]
fn test_table_html_populated_for_table_scan_fixture() {
    if !pdftoppm_available() {
        eprintln!("skipping: pdftoppm unavailable (OCR merge path needs a rasterizer)");
        return;
    }
    let Some(model) = table_model_path() else {
        eprintln!(
            "skipping: no SLANet-plus model (set {TABLE_MODEL_ENV}) — \
             table_html stays None by contract without it"
        );
        return;
    };
    let bytes = std::fs::read(TABLE_SCAN).expect("fixture");
    let _scope = EnvScope::acquire();
    pin_model_env(&model);
    let engine: &mut dyn PageOcrEngine = &mut GeometryEngine;
    let (combined, _pages) =
        parse_pdf_split(&bytes, Some(engine), 150, 50, false).expect("split parse");
    drop(_scope);

    let ocr = combined.metadata.extra.get("ocr").expect("ocr metadata");
    let html = ocr
        .get("table_html")
        .and_then(|t| t.as_array())
        .and_then(|a| a.first())
        .and_then(|e| e["html"].as_str())
        .expect(
            "model present + table geometry ⇒ table_html must be Some \
             (SLANet produced no td tokens?)",
        );
    assert!(html.starts_with("<table>"), "html must open a table element: {html}");
    assert!(html.ends_with("</table>"), "html must close the table element: {html}");
    let opens = html.matches("<td").count();
    let closes = html.matches("</td>").count();
    if opens > 0 {
        assert_eq!(opens, closes, "every opened td must close: {html}");
    }
}

/// Spec scenario: real model + real engine — pdf_table_scan produces
/// table_html. Deterministic given the fixture (Indra det/rec are fixed);
/// skips loudly without the model so the default suite stays green.
#[test]
fn test_real_engine_produces_table_html_for_table_scan() {
    if !pdftoppm_available() {
        eprintln!("skipping: pdftoppm unavailable (OCR merge path needs a rasterizer)");
        return;
    }
    let Some(model) = table_model_path() else {
        eprintln!(
            "skipping: no SLANet-plus model (set {TABLE_MODEL_ENV} to a model file) — \
             the real-model scenario needs it"
        );
        return;
    };
    let Ok(mut engine) = office_parser::IndraNative::load() else {
        eprintln!(
            "skipping: Indra models unavailable ({})",
            office_parser::IndraNative::load().err().unwrap_or_default()
        );
        return;
    };
    let bytes = std::fs::read(TABLE_SCAN).expect("fixture");
    let _scope = EnvScope::acquire();
    pin_model_env(&model);
    let doc = engine
        .parse_document(&bytes, 0.60)
        .expect("parse table scan with model");
    drop(_scope);

    let ocr = doc.metadata.extra.get("ocr").expect("ocr metadata present");
    assert!(
        ocr.get("layout_regions").is_some(),
        "Indra must emit a Table block on pdf_table_scan (fixture contract)"
    );
    let html = ocr
        .get("table_html")
        .and_then(|t| t.as_array())
        .and_then(|a| a.first())
        .and_then(|e| e["html"].as_str())
        .expect("real model + Indra table region ⇒ table_html must be Some");
    assert!(html.contains("<table"), "wrapper missing: {html}");
    assert!(html.contains("<td"), "no cells decoded: {html}");
}
