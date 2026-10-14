//! Indra OCR options + engine dispatch for scanned-PDF recovery.
//!
//! Rust-only (Mode B): the native ONNX pipeline in
//! [`crate::formats::indra_native::IndraNative`] runs in-process — no
//! Python, no subprocess. Rendering goes through pdfium (preferred) or
//! pdftoppm inside `pdf::render_pages_png_opts`.

use crate::document::PdfTextQuality;

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

/// Parse a scanned PDF with the native engine. Errors when the models are
/// unavailable (set `INDRA_OCR_MODEL_DIR` or download via huggingface_hub).
pub fn parse_auto(bytes: &[u8], opts: &PdfOcrOptions) -> crate::Result<crate::Document> {
    parse_auto_split(bytes, opts, true).map(|(doc, _)| doc)
}

/// [`parse_auto`] plus per-page splitting (1 page = 1 document). Without
/// OCR routing the split is text-layer only.
pub fn parse_auto_split(
    bytes: &[u8],
    opts: &PdfOcrOptions,
    ocr: bool,
) -> crate::Result<(
    crate::Document,
    Vec<(u32, crate::Document)>,
)> {
    if !ocr {
        return crate::formats::pdf::parse_pdf_split(bytes, None, opts.dpi, opts.min_chars, opts.force);
    }
    // split + ocr routes EVERY page through the engine (the user wants the
    // engine's layout output per page, not the raw text layer)
    let force = true;
    let mut engine = crate::formats::indra_native::IndraNative::load().map_err(|e| {
        crate::Error::from_parse(
            crate::Format::Pdf,
            anyhow::anyhow!(
                "Indra models unavailable ({e}); set INDRA_OCR_MODEL_DIR or download \
                 via huggingface_hub (datalab-to/indra-ocr)"
            ),
        )
    })?;
    crate::formats::pdf::parse_pdf_split(
        bytes,
        Some(&mut engine),
        opts.dpi,
        opts.min_chars,
        force,
    )
}

/// Kept for API compatibility; the OCR quality classification itself is
/// document-level metadata from `pdf::parse`.
pub fn classify(bytes: &[u8]) -> crate::Result<Option<PdfTextQuality>> {
    Ok(super::pdf::parse(bytes)?.metadata.pdf_text_quality)
}
