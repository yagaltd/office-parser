//! Per-page PDF text-quality types (plan 20261108-001, TASK 5).
//!
//! The document-wide [`crate::PdfTextQuality`] is unchanged; these types add
//! per-page detail. They are populated by the PDF parser (TASK 7) and are
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
    /// Detected text-column count; `1` for single-column pages, `0` when the
    /// page carries no text spans at all (image-only — nothing to measure).
    /// Counted by [`count_columns`] from the span boxes.
    pub columns: usize,
    /// `true` when the page carries a non-zero native `/Rotate` — the rotation
    /// signal the geometry pass preserves (`PageGeometry::initial_rotation`).
    /// Span boxes are axis-aligned bounding boxes, so a content-stream tilt is
    /// not recoverable from them; `/Rotate` is the honest standalone flag.
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

/// Routing decision for one page, from its text-layer richness, the OCR char
/// gate, and the page's rotation flag.
///
/// Order matters: an image-only page (`text_richness == 0`) routes to OCR
/// even when it is also rotated — the image gate comes first, exactly like
/// `PdfTextQuality::ImageOnly` gates the document-level router.
pub fn page_verdict(text_richness: u32, min_chars: usize, rotated: bool) -> PageVerdict {
    if text_richness == 0 {
        PageVerdict::Ocr {
            reason: "image-only page".to_string(),
        }
    } else if (text_richness as usize) < min_chars {
        PageVerdict::Ocr {
            reason: format!("sparse text layer ({text_richness} chars < {min_chars})"),
        }
    } else if rotated {
        PageVerdict::Escalate {
            complexity: Some("rotated page".to_string()),
        }
    } else {
        PageVerdict::Trust
    }
}

/// Horizontal text-column count for one page, derived from its span boxes.
///
/// Methodology provenance: `.workflows/spikes/pdf-complexity/README.md`
/// §Methodology — TASK 3 measured `gutters = 0` / `columns = 1` on every
/// text-layer page of the corpus with this algorithm's line-based variant;
/// this span-level union variant supersedes it after the gate found that
/// same-row two-column spans merge into one full-width line under greedy
/// y-overlap grouping (masking the gutter).
///
/// Algorithm: full-width spans (headings/rules, wider than 70% of the text
/// width) are excluded; gutters are the uncovered strips of the remaining
/// span-union x-intervals at least `max(2 × median span font size, 12pt)`
/// wide; `columns = gutters + 1`.
///
/// Returns `0` when the page has no text spans (image-only): no layout to
/// measure.
#[cfg(feature = "pdf")]
pub fn count_columns(spans: &[crate::formats::pdf::SpanBox]) -> usize {
    if spans.is_empty() {
        return 0;
    }

    // Span-level union (no line merging — same-row column spans cannot mask a
    // gutter): full-width spans (headings/rules, wider than 70% of the text
    // width) are excluded, the union of the remaining x-intervals is computed,
    // and uncovered strips at least max(2 x median font size, 12pt) wide count
    // as gutters. columns = gutters + 1.
    let text_min = spans.iter().map(|s| s.x0).fold(f32::INFINITY, f32::min);
    let text_max = spans.iter().map(|s| s.x1).fold(f32::NEG_INFINITY, f32::max);
    let text_width = (text_max - text_min).max(1.0);

    let mut iv: Vec<(f32, f32)> = spans
        .iter()
        .map(|s| (s.x0, s.x1))
        .filter(|(x0, x1)| x1 - x0 <= 0.70 * text_width)
        .collect();
    if iv.is_empty() {
        // Every span is full-width: a single column by definition.
        return 1;
    }
    iv.sort_by(|a, b| a.0.total_cmp(&b.0));

    let mut sizes: Vec<f32> = spans.iter().map(|s| s.font_size).collect();
    sizes.sort_by(f32::total_cmp);
    let mid = sizes.len() / 2;
    let median_font = if sizes.len() % 2 == 1 {
        sizes[mid]
    } else {
        (sizes[mid - 1] + sizes[mid]) / 2.0
    };
    let min_gutter = (2.0 * median_font).max(12.0);

    let mut gutters = 0usize;
    let mut prev_end = iv[0].1;
    for &(x0, x1) in iv.iter().skip(1) {
        if x0 - prev_end >= min_gutter {
            gutters += 1;
        }
        prev_end = prev_end.max(x1);
    }
    gutters + 1
}
