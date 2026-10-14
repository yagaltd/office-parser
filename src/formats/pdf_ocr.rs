//! Indra OCR pipeline for scanned PDFs (Mode A: subprocess).
//!
//! When `parse_pdf` classifies a document as [`PdfTextQuality::ImageOnly`]
//! (fewer than 50 alphanumeric characters — likely a scan), the bytes can be
//! pushed through the [Indra OCR](https://huggingface.co/datalab-to/indra-ocr)
//! engine instead of returning an empty shell:
//!
//! 1. render every page to PNG (`render_pages_png`: pdfium feature or the
//!    system `pdftoppm`)
//! 2. run the `indra-ocr` CLI over the rendered pages (subprocess — requires
//!    `pip install indra-ocr`, override the binary with `INDRA_OCR_BIN`)
//! 3. map Indra's layout JSON (`blocks[{label, bbox, text}]`, per-line
//!    `conf`) onto the office-parser [`Block`] model
//! 4. re-run the normal render path, so scanned documents come out as plain
//!    headings/paragraphs like every other format
//!
//! Page headers/footers are dropped (consistent with DOCX header/footer
//! policy); unknown labels fall back to paragraphs, so engine version drift
//! degrades instead of failing. Aggregate per-page confidence is recorded in
//! `metadata.extra["ocr"]["low_confidence_pages"]` for downstream fallback
//! (e.g. routing hard pages to a vision LLM).

use std::path::PathBuf;

use serde::Deserialize;

use crate::document_ast::{Block, ListItem, SourceSpan};

pub const DEFAULT_DPI: u32 = 200;
pub const DEFAULT_MIN_PAGE_CONFIDENCE: f64 = 0.60;

/// Options for the Indra OCR pass.
#[derive(Debug, Clone)]
pub struct PdfOcrOptions {
    /// Render DPI fed to Indra (it natively renders at ~200 dpi).
    pub dpi: u32,
    /// Pages whose mean line confidence falls below this are recorded in
    /// `metadata.extra["ocr"]["low_confidence_pages"]`.
    pub min_page_confidence: f64,
    /// A page whose text layer has fewer alphanumeric chars than this is
    /// routed to OCR. Lower it when pages carry thin/garbage text layers.
    pub min_chars: usize,
    /// OCR every page regardless of its text layer (the text layer is then
    /// discarded for pages that have one).
    pub force: bool,
}

impl Default for PdfOcrOptions {
    fn default() -> Self {
        Self {
            dpi: DEFAULT_DPI,
            min_page_confidence: DEFAULT_MIN_PAGE_CONFIDENCE,
            min_chars: super::pdf::PAGE_TEXT_THRESHOLD,
            force: false,
        }
    }
}

impl PdfOcrOptions {
    fn indra_bin() -> PathBuf {
        std::env::var_os("INDRA_OCR_BIN")
            .map(PathBuf::from)
            .unwrap_or_else(|| PathBuf::from("indra-ocr"))
    }
}

// --- Indra JSONL schema (tolerant: unknown fields ignored, defaults everywhere)

#[derive(Deserialize)]
struct IndraPage {
    #[serde(default)]
    page: usize,
    #[serde(default)]
    blocks: Vec<IndraBlock>,
    #[serde(default)]
    lines: Vec<IndraLine>,
}

#[derive(Deserialize)]
struct IndraBlock {
    #[serde(default)]
    label: String,
    #[serde(default)]
    text: String,
}

#[derive(Deserialize)]
struct IndraLine {
    #[serde(default)]
    conf: f64,
}

// --- mapping

fn spans() -> SourceSpan {
    SourceSpan::default() // OCR text has no position in any extracted-text layer
}

/// Map one Indra layout block to office-parser blocks.
/// `list_acc` accumulates consecutive List-item blocks into one List block.
fn map_label(
    label: &str,
    text: String,
    index: usize,
    list_acc: &mut Vec<ListItem>,
    out: &mut Vec<Block>,
) {
    let text = text.trim_end().to_string();
    if text.trim().is_empty() {
        return;
    }
    match label {
        // Page furniture: dropped, consistent with DOCX headers/footers.
        "Page-Header" | "Page-Footer" | "Picture" => {}
        "Title" => {
            flush_list(list_acc, out);
            out.push(Block::Heading {
                block_index: index,
                level: 1,
                text,
                marks: Vec::new(),
                source: spans(),
            });
        }
        "Section-Header" => {
            flush_list(list_acc, out);
            out.push(Block::Heading {
                block_index: index,
                level: 2,
                text,
                marks: Vec::new(),
                source: spans(),
            });
        }
        "List-item" => {
            list_acc.push(ListItem {
                level: 1,
                text,
                marks: Vec::new(),
                source: spans(),
            });
        }
        // Text, Caption, Footnote, Table (markdown text in v1), unknown labels.
        _ => {
            flush_list(list_acc, out);
            out.push(Block::Paragraph {
                block_index: index,
                text,
                marks: Vec::new(),
                source: spans(),
            });
        }
    }
}

fn flush_list(list_acc: &mut Vec<ListItem>, out: &mut Vec<Block>) {
    if list_acc.is_empty() {
        return;
    }
    let items = std::mem::take(list_acc);
    out.push(Block::List {
        block_index: out.len(),
        ordered: false,
        items,
        source: spans(),
    });
}

/// Map raw Indra JSONL output (one JSON page per line) into blocks,
/// concatenating pages in order. Public for tests and direct reuse.
pub fn map_jsonl_to_blocks(raw: &[u8]) -> Vec<Block> {
    let mut blocks: Vec<Block> = parse_jsonl(raw)
        .iter()
        .flat_map(|page| map_page_to_blocks(page))
        .collect();
    // block_index is global across pages (renderer contract)
    for (i, b) in blocks.iter_mut().enumerate() {
        set_block_index(b, i);
    }
    blocks
}

fn map_page_to_blocks(page: &IndraPage) -> Vec<Block> {
    let mut out = Vec::new();
    let mut list_acc: Vec<ListItem> = Vec::new();
    for b in &page.blocks {
        map_label(&b.label, b.text.clone(), out.len(), &mut list_acc, &mut out);
    }
    flush_list(&mut list_acc, &mut out);
    // normalize block_index to array position
    for (i, b) in out.iter_mut().enumerate() {
        set_block_index(b, i);
    }
    out
}

fn set_block_index(b: &mut Block, i: usize) {
    match b {
        Block::Heading { block_index, .. }
        | Block::Paragraph { block_index, .. }
        | Block::List { block_index, .. }
        | Block::Table { block_index, .. }
        | Block::Image { block_index, .. }
        | Block::Link { block_index, .. }
        | Block::Note { block_index, .. } => *block_index = i,
    }
}

fn mean_confidence(page: &IndraPage) -> Option<f64> {
    if page.lines.is_empty() {
        return None;
    }
    let n = page.lines.len() as f64;
    Some(page.lines.iter().map(|l| l.conf).sum::<f64>() / n)
}

fn parse_jsonl(bytes: &[u8]) -> Vec<IndraPage> {
    let text = String::from_utf8_lossy(bytes);
    text.lines()
        .filter_map(|line| serde_json::from_str::<IndraPage>(line).ok())
        .collect()
}

/// [`parse_auto`] plus per-page splitting (1 page = 1 document). Without
/// `ocr` the split is text-layer only.
pub fn parse_auto_split(
    bytes: &[u8],
    opts: &PdfOcrOptions,
    ocr: bool,
) -> crate::Result<(
    crate::Document,
    Vec<(u32, crate::Document)>,
)> {
    #[cfg(feature = "ocr-indra-native")]
    if ocr {
        if let Ok(mut engine) = crate::formats::indra_native::IndraNative::load() {
            return crate::formats::pdf::parse_pdf_split(
                bytes,
                Some(&mut engine),
                opts.dpi,
                opts.min_chars,
                opts.force,
            );
        }
    }
    if ocr {
        let mut sub = SubprocessOcr::new();
        return crate::formats::pdf::parse_pdf_split(
            bytes,
            Some(&mut sub),
            opts.dpi,
            opts.min_chars,
            opts.force,
        );
    }
    crate::formats::pdf::parse_pdf_split(bytes, None, opts.dpi, opts.min_chars, opts.force)
}

/// Parse a scanned PDF with the best available Indra engine: the native
/// ONNX pipeline when compiled with `ocr-indra-native` and its models are
/// present, else the subprocess engine (Mode A).
pub fn parse_auto(bytes: &[u8], opts: &PdfOcrOptions) -> crate::Result<crate::Document> {
    #[cfg(feature = "ocr-indra-native")]
    if let Ok(mut engine) = crate::formats::indra_native::IndraNative::load() {
        return engine.parse_document_opts(bytes, opts);
    }
    parse_with_indra(bytes, opts)
}

/// Subprocess engine: one indra-ocr call per routed page.
/// (dpi/thresholds flow through `parse_with_page_ocr_opts`, not here.)
struct SubprocessOcr;

impl SubprocessOcr {
    fn new() -> Self {
        Self
    }
}

impl super::pdf::PageOcrEngine for SubprocessOcr {
    fn ocr_page_png(&mut self, page_png: &[u8]) -> Result<(Vec<Block>, Option<f64>), String> {
        static SEQ: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
        let tmp = std::env::temp_dir().join(format!(
            "office-parser-indra-{}-{}",
            std::process::id(),
            SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
        ));
        std::fs::create_dir_all(&tmp).map_err(|e| e.to_string())?;
        let png_path = tmp.join("page.png");
        let out_dir = tmp.join("out");
        std::fs::write(&png_path, page_png).map_err(|e| e.to_string())?;
        let run = run_indra(std::slice::from_ref(&png_path), &out_dir);
        let _ = std::fs::remove_dir_all(&tmp);
        let jsonl_paths = run.map_err(|e| e.to_string())?;
        let mut raw = Vec::new();
        for p in &jsonl_paths {
            raw.extend(std::fs::read(p).map_err(|e| e.to_string())?);
        }
        let mut blocks = map_jsonl_to_blocks(&raw);
        // mean line confidence for the routing metadata
        let mut confs: Vec<f64> = Vec::new();
        for line in String::from_utf8_lossy(&raw).lines() {
            if let Ok(page) = serde_json::from_str::<IndraPage>(line) {
                confs.extend(page.lines.iter().map(|l| l.conf));
            }
        }
        let conf = (!confs.is_empty()).then(|| confs.iter().sum::<f64>() / confs.len() as f64);
        for (i, b) in blocks.iter_mut().enumerate() {
            set_block_index_pub(b, i);
        }
        Ok((std::mem::take(&mut blocks), conf))
    }

    fn name(&self) -> &'static str {
        "indra"
    }
}

fn set_block_index_pub(b: &mut Block, i: usize) {
    match b {
        Block::Heading { block_index, .. }
        | Block::Paragraph { block_index, .. }
        | Block::List { block_index, .. }
        | Block::Table { block_index, .. }
        | Block::Image { block_index, .. }
        | Block::Link { block_index, .. }
        | Block::Note { block_index, .. } => *block_index = i,
    }
}

/// Parse a PDF, routing image-like pages through the Indra OCR subprocess.
pub fn parse_with_indra(bytes: &[u8], opts: &PdfOcrOptions) -> crate::Result<crate::Document> {
    super::pdf::parse_with_page_ocr_opts(
        bytes,
        &mut SubprocessOcr::new(),
        opts.dpi,
        opts.min_chars,
        opts.force,
    )
}

/// Run Indra over an already-parsed scanned PDF document and replace its
/// (nearly empty) blocks with OCR output. Public for tests and for callers
/// that want to force OCR regardless of the text-quality classification.
pub fn ocr_document(
    bytes: &[u8],
    opts: &PdfOcrOptions,
    doc: &mut crate::Document,
) -> crate::Result<()> {
    let pages_png = super::pdf::render_pages_png(bytes, opts.dpi)?;
    let tmp = std::env::temp_dir().join(format!("office-parser-indra-{}", std::process::id()));
    std::fs::create_dir_all(&tmp)
        .map_err(|e| crate::Error::from_parse(crate::Format::Pdf, anyhow::anyhow!(e.to_string())))?;
    let mut png_paths = Vec::new();
    for (i, png) in pages_png.iter().enumerate() {
        let p = tmp.join(format!("page-{i:04}.png"));
        std::fs::write(&p, png)
            .map_err(|e| crate::Error::from_parse(crate::Format::Pdf, anyhow::anyhow!(e.to_string())))?;
        png_paths.push(p);
    }
    let out_dir = tmp.join("out");
    let jsonl_paths = run_indra(&png_paths, &out_dir)?;
    let mut raw = Vec::new();
    for p in &jsonl_paths {
        raw.extend(
            std::fs::read(p)
                .map_err(|e| crate::Error::from_parse(crate::Format::Pdf, anyhow::anyhow!(e.to_string())))?,
        );
    }
    let _ = std::fs::remove_dir_all(&tmp);
    let pages = parse_jsonl(&raw);
    if pages.is_empty() {
        return Err(crate::Error::from_parse(
            crate::Format::Pdf,
            anyhow::anyhow!("indra-ocr produced no parsable page output"),
        ));
    }

    let mut blocks = Vec::new();
    let mut low_conf = Vec::new();
    for page in &pages {
        if let Some(conf) = mean_confidence(page) {
            if conf < opts.min_page_confidence {
                low_conf.push(serde_json::json!(page.page));
            }
        }
        blocks.extend(map_page_to_blocks(page));
    }

    doc.blocks = blocks;
    doc.metadata.extra["ocr"] = serde_json::json!({
        "engine": "indra",
        "dpi": opts.dpi,
        "pages_ocr": pages.len(),
        "min_page_confidence": opts.min_page_confidence,
        "low_confidence_pages": low_conf,
    });
    Ok(())
}

fn run_indra(png_paths: &[PathBuf], out_dir: &PathBuf) -> crate::Result<Vec<PathBuf>> {
    std::fs::create_dir_all(out_dir)
        .map_err(|e| crate::Error::from_parse(crate::Format::Pdf, anyhow::anyhow!(e.to_string())))?;
    let mut cmd = std::process::Command::new(PdfOcrOptions::indra_bin());
    cmd.args(png_paths)
        .arg("--format")
        .arg("json")
        .arg("--out")
        .arg(out_dir)
        .stdin(std::process::Stdio::null());
    let output = cmd
        .output()
        .map_err(|e| {
            crate::Error::from_parse(
                crate::Format::Pdf,
                anyhow::anyhow!(format!(
                    "failed to launch indra-ocr (set INDRA_OCR_BIN or `pip install indra-ocr`): {e}"
                )),
            )
        })?;
    if !output.status.success() {
        return Err(crate::Error::from_parse(
            crate::Format::Pdf,
            anyhow::anyhow!(format!(
                "indra-ocr failed ({}): {}",
                output.status,
                String::from_utf8_lossy(&output.stderr).trim()
            )),
        ));
    }
    // Indra writes one .jsonl per input; page files zero-padded sort correctly.
    let mut entries: Vec<PathBuf> = std::fs::read_dir(out_dir)
        .map_err(|e| crate::Error::from_parse(crate::Format::Pdf, anyhow::anyhow!(e.to_string())))?
        .filter_map(|e| e.ok().map(|e| e.path()))
        .filter(|p| p.extension().is_some_and(|e| e == "jsonl"))
        .collect();
    entries.sort();
    if entries.is_empty() {
        return Err(crate::Error::from_parse(
            crate::Format::Pdf,
            anyhow::anyhow!("indra-ocr wrote no .jsonl output"),
        ));
    }
    Ok(entries)
}
