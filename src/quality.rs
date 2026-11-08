//! Per-page PDF text-quality types (plan 20261108-001, TASK 5) and the
//! pdfium-free no-ink trust check (TASK 9).
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
        /// Why the text layer is insufficient (`"image-only page"`, `"no-ink"`,
        /// or a pdfium trust reason).
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

// ---------------------------------------------------------------------------
// TASK 9 — pdfium-free no-ink trust check (plan 20261108-001, task-9 spec).
//
// The geometry-based trust source: does a page's own span-box geometry
// corroborate the char count its text layer claims? Runs under the `pdf`
// feature alone — no pdfium, no precondition — so trust verdicts are never
// silently empty without `render-pdfium` (AGENTS.md defect class). pdfium's
// invisible/bad-char checks remain the additional signals where present.

/// Per-character ink-box area factor of [`no_ink_reason`]: a text run's
/// advance is ≈ 0.55 × font size per non-space glyph — the same estimate the
/// PDF interpreter records into every span box, calibrated on the TASK 2
/// spike (~300 chars at 11 pt ≈ 19,800 pt² of span boxes ≈ 0.041 of A4).
pub const NO_INK_ADVANCE_FACTOR: f32 = 0.55;

/// Fallback body font size (pt) of [`no_ink_reason`] for a page that declares
/// no usable median span font size (no spans, or every size ≤ 0/NaN): a
/// typical document body font (the fixture corpus' body text is 11 pt). It
/// gives even a box-less page a positive expected ink area to fall short of.
pub const NO_INK_NOMINAL_FONT_PT: f32 = 11.0;

/// Implausibility ratio of [`no_ink_reason`]: the measured span-box area must
/// reach this fraction of the area the claimed char count needs at the
/// page's median span font size. A HEURISTIC calibrated on the five fixtures
/// — not a calibrated probability (task-9 spec Decisions). Self-consistent
/// pages measure ≥ 100% (boxes hold at least their alphanumeric chars);
/// pages claiming ink their boxes cannot hold measure ≈ 0%.
pub const NO_INK_MIN_COVERAGE_RATIO: f32 = 0.25;

/// Machine-readable reason literal carried by [`PageVerdict::Ocr`] when the
/// no-ink check fires — deliberately distinct from the pdfium-derived reason
/// strings so consumers can tell the sources apart (task-9 spec Decisions).
pub const NO_INK_REASON: &str = "no-ink";

/// TASK 9 — the pdfium-free no-ink trust check: does the page's own span
/// geometry corroborate the char count its text layer claims?
///
/// A page that LOOKS text-rich (`text_richness >= min_chars`, the same gate
/// the per-page OCR router applies) but whose [`SpanBox`]es hold implausibly
/// little ink for that char count is untrusted: the claimed text does not
/// exist where the geometry says the page put its text. Typical producers of
/// the contradiction are invisible OCR layers whose spans carry collapsed or
/// degenerate text matrices (zero-area boxes) or never entered the content
/// stream the geometry pass walked.
///
/// Implausibility rule (constants above, in one place): a plausible page
/// needs `chars × NO_INK_ADVANCE_FACTOR × font_pt²` pt² of span boxes, with
/// `font_pt` the page's median span font size — or
/// [`NO_INK_NOMINAL_FONT_PT`] when the page declares none. Flag when the
/// measured box area falls below [`NO_INK_MIN_COVERAGE_RATIO`] of that
/// expectation. Both sides are compared in pt²: dividing both by the (fixed,
/// positive) page area yields the coverage form of the spec diagram.
///
/// Design property (median-font rule): a page whose boxes are merely small
/// but self-consistent (every span at 1 pt) is NOT flagged — this check
/// detects a char-count-vs-geometry contradiction, not absolute typesetting
/// taste; pdfium's render-mode-3 invisibility signal stays the detector for
/// genuinely invisible-but-boxed layers.
///
/// Returns `Some(NO_INK_REASON)` when the page is untrusted, `None` when the
/// char gates own the page (`text_richness < min_chars`), when the page area
/// is undefined/zero/NaN (no verdict from a zero-area page), or when the
/// boxes corroborate the char count.
#[cfg(feature = "pdf")]
pub fn no_ink_reason(
    text_richness: u32,
    min_chars: usize,
    spans: &[crate::formats::pdf::SpanBox],
    page_area_pt2: f32,
) -> Option<String> {
    // Only a page that looks text-rich can lie about its ink; sparser pages
    // already route via the image-only / sparse char gates.
    if (text_richness as usize) < min_chars {
        return None;
    }
    // Undefined page area (no MediaBox/CropBox in the /Parent chain), a
    // zero-area page, or NaN: nothing to divide by, no verdict from this
    // check. `!(x > 0.0)` deliberately also rejects NaN — `x <= 0.0` alone
    // would let a poisoned area pass.
    if !(page_area_pt2 > 0.0) {
        return None;
    }
    let measured = glyph_box_area_pt2(spans);
    let font_pt = median_font_size(spans)
        .filter(|fs| fs.is_finite() && *fs > 0.0)
        .unwrap_or(NO_INK_NOMINAL_FONT_PT);
    let expected = text_richness as f32 * NO_INK_ADVANCE_FACTOR * font_pt * font_pt;
    if measured < NO_INK_MIN_COVERAGE_RATIO * expected {
        Some(NO_INK_REASON.to_string())
    } else {
        None
    }
}

/// Summed span-box area in pt²: `Σ (x1−x0)·(y1−y0)`, negative edges clamped
/// to 0. Overlapping spans double-count — the TASK 2 spike's stated caveat;
/// the overlap only inflates the measured area, i.e. errs against flagging.
#[cfg(feature = "pdf")]
fn glyph_box_area_pt2(spans: &[crate::formats::pdf::SpanBox]) -> f32 {
    spans
        .iter()
        .map(|s| (s.x1 - s.x0).max(0.0) * (s.y1 - s.y0).max(0.0))
        .sum()
}

/// Median span font size (sorted middle; even counts average the pair),
/// `None` for a span-less page — the same statistic [`count_columns`] uses
/// for its gutter width and [`no_ink_reason`] for its expected ink area.
#[cfg(feature = "pdf")]
fn median_font_size(spans: &[crate::formats::pdf::SpanBox]) -> Option<f32> {
    if spans.is_empty() {
        return None;
    }
    let mut sizes: Vec<f32> = spans.iter().map(|s| s.font_size).collect();
    sizes.sort_by(f32::total_cmp);
    let mid = sizes.len() / 2;
    Some(if sizes.len() % 2 == 1 {
        sizes[mid]
    } else {
        (sizes[mid - 1] + sizes[mid]) / 2.0
    })
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

    let min_gutter = (2.0 * median_font_size(spans).unwrap_or(0.0)).max(12.0);

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
