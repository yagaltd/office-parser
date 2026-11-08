//! Per-page PDF text-quality types (plan 20261108-001, TASK 5).
//!
//! The document-wide [`crate::PdfTextQuality`] is unchanged; these types add
//! per-page detail. They are populated by the PDF parser (TASK 8) and are
//! `None` for every non-PDF format, mirroring the `pdf_text_quality` discipline.

use serde::{Deserialize, Serialize};

/// Text quality assessed for a single PDF page.
///
/// JSON shape (emitted by [`crate::render::to_json_value`] under
/// `metadata.page_quality`):
///
/// ```json
/// {
///   "page": 1,
///   "text_richness": 310,
///   "columns": 1,
///   "rotated": false,
///   "verdict": { "kind": "trust" }
/// }
/// ```
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct PageTextQuality {
    /// 1-based page number.
    pub page: u32,
    /// Alphanumeric characters extracted from the page's text layer
    /// (`0` for image-only pages).
    pub text_richness: u32,
    /// Detected text-column count; `1` for single-column pages.
    pub columns: usize,
    /// `true` when the page's dominant text runs at a non-zero rotation.
    pub rotated: bool,
    /// Routing decision for the page.
    pub verdict: PageVerdict,
}

/// Routing decision for one PDF page.
///
/// JSON is internally tagged on `kind` (`trust` | `ocr` | `escalate`) so the
/// shape is self-describing and extensible.
///
/// `Ocr` and `Escalate` are deliberately distinct variants: OCR recovers
/// unreadable raster text, but it cannot repair a reading order the text
/// layer already mangled — those pages escalate instead
/// (`.workflows/CONTEXT.md` §Domain Rules).
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum PageVerdict {
    /// Geometry and text layer agree; use the extracted text as-is.
    Trust,
    /// Page needs OCR; `reason` states why the text layer was not trusted.
    Ocr {
        /// Why the text layer is insufficient (e.g. `"image-only page"`).
        reason: String,
    },
    /// Layout needs more than OCR; `complexity` names the triggering signal
    /// when one is known.
    Escalate {
        /// Layout signal that triggered escalation, if any.
        complexity: Option<String>,
    },
}
