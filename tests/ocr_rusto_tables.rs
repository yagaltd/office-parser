#![cfg(all(feature = "pdf", feature = "ocr-indra-native", feature = "ocr-rusto-core"))]
//! TASK 11 (plan 20261108-001): rusto-core table-region enrichment.
//! Contract: `.workflows/specs/task-11.spec` + human amendment (rusto-core as
//! pinned git+rev dep; `rusto-rs` facade still forbidden).
//!
//! Model reality: no DocOCR layout model exists in this checkout or HF cache,
//! so the enriched-path tests skip loudly (quality_routing.rs pdftoppm
//! pattern) when `OFFICE_PARSER_LAYOUT_MODEL`/`models/DocOCR/layout.onnx` is
//! absent. The no-model arm is always runnable and asserts the feature-off
//! payload shape instead of silently passing.
//!
//! Scenario split:
//! - test_transcripts_identical_with_feature_on_and_off — spec scenario 2
//!   verbatim (pdf_scanned.pdf, both arms in one binary via env toggling).
//! - test_table_html_populated_for_table_scan_fixture — spec scenario 3 with
//!   a deterministic stub engine (cells known a priori; Indra's Table-block
//!   emission is not this task's contract).
//! - test_no_model_feature_off_payload_shape — feature-off payload shape.
//! - test_indra_recognition_source_unchanged — spec scenario 4 (git diff).

use office_parser::document_ast::{Block, Cell, SourceSpan};
use office_parser::formats::pdf::{parse_pdf_split, PageOcrEngine};
use office_parser::formats::rusto_tables::LAYOUT_MODEL_ENV;

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
            prior: std::env::var_os(LAYOUT_MODEL_ENV),
        }
    }
}

impl Drop for EnvScope {
    fn drop(&mut self) {
        match self.prior.take() {
            Some(v) => unsafe { std::env::set_var(LAYOUT_MODEL_ENV, v) },
            None => unsafe { std::env::remove_var(LAYOUT_MODEL_ENV) },
        }
    }
}

fn pin_model_env(model: &std::path::Path) {
    // SAFETY: tests in this binary touching OFFICE_PARSER_LAYOUT_MODEL hold
    // ENV_LOCK for their whole scope (set_var/remove_var are unsafe in
    // edition 2024).
    unsafe { std::env::set_var(LAYOUT_MODEL_ENV, model) };
}

fn hide_model_env() {
    unsafe { std::env::set_var(LAYOUT_MODEL_ENV, "/nonexistent/office-parser-layout.onnx") };
}

fn pdftoppm_available() -> bool {
    std::env::var_os("PATH").map_or(false, |p| {
        std::env::split_paths(&p)
            .any(|dir| dir.join("pdftoppm").is_file() || dir.join("pdftoppm.exe").is_file())
    })
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

/// Spec scenario 2. Both arms run in one binary: ON pins the resolved layout
/// model (detector active), OFF hides it (enrichment provably no-op — the
/// bogus path fails before any decode). Confidence has no per-page accessor;
/// `low_confidence_pages` is the confidence-derived observable, and the
/// confidence VALUE is moved through enrich_ocr_page untouched by
/// construction (no code path writes it).
#[test]
fn test_transcripts_identical_with_feature_on_and_off() {
    let Some(model) = office_parser::formats::rusto_tables::layout_model_path() else {
        eprintln!(
            "skipping: no DocOCR layout model (set {LAYOUT_MODEL_ENV} to a layout \
             model file) — transcript identity with the detector active needs it"
        );
        return;
    };
    let _scope = EnvScope::acquire();
    let Ok(mut engine) = office_parser::IndraNative::load() else {
        eprintln!(
            "skipping: Indra models unavailable ({})",
            office_parser::IndraNative::load().err().unwrap_or_default()
        );
        return;
    };
    let bytes = std::fs::read(SCAN).expect("fixture");

    pin_model_env(&model);
    let doc_on = engine.parse_document(&bytes, 0.60).expect("parse (enrichment on)");
    hide_model_env();
    let doc_off = engine.parse_document(&bytes, 0.60).expect("parse (enrichment off)");

    assert_eq!(
        transcript(&doc_on),
        transcript(&doc_off),
        "transcript must be identical with enrichment on/off"
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
    assert!(ocr_off.get("layout_regions").is_none());
    assert!(ocr_off.get("word_boxes").is_none());
    assert!(ocr_off.get("table_html").is_none());
}

/// Spec scenario 3, positive arm — deterministic stub engine (the cells are
/// known a priori; whether Indra itself labels a block "Table" is indra_
/// native.rs territory, not this feature's contract). Needs pdftoppm (or
/// pdfium) + the layout model; skips loudly otherwise. The negative arm
/// ("disabled ⇒ None") is covered by test_no_model_feature_off_payload_shape.
#[test]
fn test_table_html_populated_for_table_scan_fixture() {
    if !pdftoppm_available() {
        eprintln!("skipping: pdftoppm unavailable (OCR merge path needs a rasterizer)");
        return;
    }
    let Some(model) = office_parser::formats::rusto_tables::layout_model_path() else {
        eprintln!(
            "skipping: no DocOCR layout model (set {LAYOUT_MODEL_ENV}) — \
             table_html stays None by contract without it"
        );
        return;
    };
    struct TableEngine;
    impl PageOcrEngine for TableEngine {
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
        fn name(&self) -> &'static str {
            "table-stub"
        }
    }

    let bytes = std::fs::read(TABLE_SCAN).expect("fixture");
    let _scope = EnvScope::acquire();
    pin_model_env(&model);
    let engine: &mut dyn PageOcrEngine = &mut TableEngine;
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
            "layout model present + table fixture + Table block ⇒ table_html must be Some \
             (rusto found no LayoutType::Table region?)",
        );
    assert!(html.contains("<table"), "html must open a table element: {html}");
    assert!(html.contains("<td>alpha</td>"), "cell text must render: {html}");
    assert!(html.contains("</table>"), "html must close the table: {html}");
    // Transcript identity at the seam: the stub payload passes through
    // byte-for-byte (only block_index is renumbered by the merge).
    assert_eq!(combined.blocks.len(), 1, "stub payload passes through");
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
    // Documented decision: word boxes are never populated by this feature.
    assert!(ocr.get("word_boxes").is_none());
}

/// Feature-off payload shape: with the layout model unresolvable, the ocr
/// metadata carries none of the rich keys — exactly what a build without the
/// feature produces (spec scenario 3, negative arm). Always runnable where
/// Indra + a rasterizer exist; skips loudly otherwise.
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
        .expect("parse with hidden layout model");
    drop(_scope);

    let ocr = doc.metadata.extra.get("ocr").expect("ocr metadata present");
    assert_eq!(ocr["pages_ocr"], 2);
    for key in ["layout_regions", "word_boxes", "table_html"] {
        assert!(
            ocr.get(key).is_none(),
            "{key} must stay absent without the layout model (feature-off payload shape)"
        );
    }
}

/// Spec scenario 4: Indra's recognition source is untouched by this task.
#[test]
fn test_indra_recognition_source_unchanged() {
    let out = std::process::Command::new("git")
        .args(["diff", "--name-only", "--", "src/formats/indra_native.rs"])
        .current_dir(env!("CARGO_MANIFEST_DIR"))
        .output()
        .expect("run git");
    assert!(
        out.status.success(),
        "git diff failed: {}",
        String::from_utf8_lossy(&out.stderr)
    );
    let changed = String::from_utf8_lossy(&out.stdout);
    assert!(
        changed.trim().is_empty(),
        "src/formats/indra_native.rs must stay untouched; diff lists: {changed}"
    );
}
