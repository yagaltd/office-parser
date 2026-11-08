#![cfg(feature = "pdf")]
//! TASK 9 (plan 20261108-001) — the pdfium-free no-ink trust check.
//!
//! Spec: `.workflows/specs/task-9.spec` under the RE-SCOPED geometry-only
//! reading: the no-ink check reads ONLY TASK 4 page geometry (`SpanBox`es ×
//! MediaBox/CropBox area); `render-pdfium` is an additional signal, never a
//! precondition. Green on both
//! `cargo test --features pdf --test pdf_no_ink` and
//! `cargo test --no-default-features --features pdf --test pdf_no_ink`.
//!
//! Regime map (per task contract):
//! - checker math + boundary constants → this file, against
//!   `quality::no_ink_reason` directly;
//! - `build_page_quality` wiring with a hand-built `PageGeometry` → in-crate
//!   unit tests in `src/formats/pdf.rs` (the fn is private by design);
//! - end-to-end → the existing fixture corpus here (the fixtures are
//!   self-consistent: `pdf_scanned` has zero spans AND zero chars, so the
//!   image-only gate owns it — no-ink must NOT fire).

use office_parser::formats::pdf::{parse_pdf_split, SpanBox};
use office_parser::quality::{
    no_ink_reason, NO_INK_ADVANCE_FACTOR, NO_INK_MIN_COVERAGE_RATIO, NO_INK_REASON, PageVerdict,
};

const A4_AREA_PT2: f32 = 612.0 * 792.0;

fn span(x0: f32, y0: f32, x1: f32, y1: f32, font_size: f32, chars: usize) -> SpanBox {
    SpanBox {
        seq: 0,
        x0,
        y0,
        x1,
        y1,
        font_name: "F1".to_string(),
        font_size,
        text: "x".repeat(chars),
        char_count: chars,
    }
}

/// Regime 1 — a normal rich page measures plausible ink and is not flagged.
/// Shape mirrored from the TASK 2 spike's pdf_mixed p1 box dump (8 spans,
/// body text at 11 pt, measured ≈ 27.9k pt² ≈ 131% of the alphanumeric
/// expectation — comfortably above the 25% floor).
#[test]
fn normal_rich_page_measures_plausible_ink() {
    let mut spans = vec![
        span(72.0, 720.0, 281.88, 738.0, 18.0, 22), // "Quarterly Field Report"
        span(72.0, 680.0, 259.33, 693.0, 13.0, 27), // subheading
    ];
    for row in 0..6u32 {
        let y = 650.0 - 15.0 * row as f32;
        spans.push(span(72.0, y, 400.0, y + 11.0, 11.0, 45)); // body lines
    }
    let chars: u32 = spans.iter().map(|s| s.char_count as u32).sum();
    assert_eq!(chars, 319);
    assert_eq!(no_ink_reason(chars, 50, &spans, A4_AREA_PT2), None);
}

/// Regime 2 — the contradiction the check exists for: the text layer claims a
/// rich page but the geometry pass walked zero span boxes (collapsed /
/// degenerate text matrices, or an ink-less OCR layer). With no spans the
/// expectation rests on NO_INK_NOMINAL_FONT_PT — and is still positive, so
/// the measured 0 pt² falls short of it.
#[test]
fn rich_char_count_with_no_span_boxes_is_no_ink() {
    assert_eq!(
        no_ink_reason(3_000, 50, &[], A4_AREA_PT2),
        Some(NO_INK_REASON.to_string())
    );
}

/// Regime 3 — boundary constants: the flag is strict `measured <
/// NO_INK_MIN_COVERAGE_RATIO × expected`, pinned here to one exact pt² of
/// the boundary (1000 claimed chars at median 11 pt: expected = 66,550 pt²,
/// boundary = 16,637.5 pt²).
#[test]
fn implausibility_boundary_tracks_the_constants() {
    let claimed = 1_000.0_f32;
    let fs = 11.0_f32;
    let expected = claimed * NO_INK_ADVANCE_FACTOR * fs * fs;
    let boundary = NO_INK_MIN_COVERAGE_RATIO * expected;
    // Exactly at the boundary (a boundary/fs × fs pt box): plausible.
    let at_boundary = [span(0.0, 0.0, boundary / fs, fs, fs, 1_000)];
    assert_eq!(no_ink_reason(1_000, 50, &at_boundary, A4_AREA_PT2), None);
    // One pt² less of box area: implausible → no-ink.
    let below = [span(0.0, 0.0, (boundary - 1.0) / fs, fs, fs, 1_000)];
    assert_eq!(
        no_ink_reason(1_000, 50, &below, A4_AREA_PT2),
        Some(NO_INK_REASON.to_string())
    );
}

/// The expectation rides on the page's MEDIAN span font size: identical box
/// widths and char count flip the verdict when only the median font differs.
#[test]
fn expected_ink_uses_the_median_span_font_size() {
    // Three 40-char spans; per-span box area = 90 × its own font size.
    let mk = |fs1: f32, fs2: f32, fs3: f32| {
        vec![
            span(0.0, 0.0, 90.0, fs1, fs1, 40),
            span(0.0, 20.0, 90.0, 20.0 + fs2, fs2, 40),
            span(0.0, 60.0, 90.0, 60.0 + fs3, fs3, 40),
        ]
    };
    // Median 8 pt (8, 8, 18): expected = 120 × 0.55 × 64 = 4,224 pt²;
    // measured 3,060 pt² corroborates the claim.
    assert_eq!(no_ink_reason(120, 50, &mk(8.0, 8.0, 18.0), A4_AREA_PT2), None);
    // Median 18 pt (8, 18, 18): expected = 120 × 0.55 × 324 = 21,384 pt²;
    // measured 3,960 pt² is ~18% of it — implausible → no-ink.
    assert_eq!(
        no_ink_reason(120, 50, &mk(8.0, 18.0, 18.0), A4_AREA_PT2),
        Some(NO_INK_REASON.to_string())
    );
}

/// Edge case (spec scenario): a zero/undefined page area defines the result
/// (no NaN, no panic) and derives no trust verdict.
#[test]
fn zero_page_area_does_not_panic_and_yields_no_verdict() {
    let spans = [span(0.0, 0.0, 10.0, 10.0, 11.0, 5)];
    assert_eq!(no_ink_reason(3_000, 50, &spans, 0.0), None);
    assert_eq!(no_ink_reason(3_000, 50, &spans, -5.0), None);
    assert_eq!(no_ink_reason(3_000, 50, &spans, f32::NAN), None);
}

/// The char gates own sparse pages: below `min_chars` the no-ink check
/// abstains even with zero boxes (the reason stays "sparse text layer (...)"
/// or "image-only page" — never "no-ink").
#[test]
fn sparse_pages_are_the_char_gates_business() {
    assert_eq!(no_ink_reason(30, 50, &[], A4_AREA_PT2), None);
    assert_eq!(no_ink_reason(0, 50, &[], A4_AREA_PT2), None);
}

/// End-to-end (spec, re-scoped): pdf_scanned carries zero spans AND zero
/// chars, so the IMAGE-ONLY gate owns the verdict — no-ink must not fire
/// (its richness gate requires a page that looks text-rich). Asserts the
/// reason literal stays "image-only page" on the no-pdfium build.
#[test]
fn scanned_pages_route_image_only_not_no_ink_without_pdfium() {
    let bytes = std::fs::read("tests/fixtures/pdf_scanned.pdf").expect("fixture");
    let doc = office_parser::pdf::parse(&bytes).expect("parses");
    let pages = doc.metadata.page_quality.expect("page quality");
    assert_eq!(pages.len(), 2);
    for q in &pages {
        assert_eq!(q.text_richness, 0);
        match &q.verdict {
            PageVerdict::Ocr { reason } => {
                assert_eq!(reason, "image-only page", "gate order: image-only, not {reason:?}")
            }
            other => panic!("expected Ocr, got {other:?}"),
        }
    }
}

/// Negative (spec): the corpus' real text page is not falsely untrusted.
#[test]
fn text_page_remains_trusted() {
    let bytes = std::fs::read("tests/fixtures/pdf_mixed.pdf").expect("fixture");
    let doc = office_parser::pdf::parse(&bytes).expect("parses");
    let pages = doc.metadata.page_quality.expect("page quality");
    assert_eq!(pages.len(), 2);
    assert_eq!(pages[0].verdict, PageVerdict::Trust);
    assert!(matches!(
        &pages[1].verdict,
        PageVerdict::Ocr { reason } if reason == "image-only page"
    ));
}

/// Regression (spec + AGENTS.md defect class): without render-pdfium the
/// trust-check output is not a silent empty vector — every parsed page of
/// the corpus carries a verdict record, and every untrusted page names a
/// non-empty reason derived from gates that need no pdfium.
#[test]
fn trust_verdicts_are_never_silently_empty_without_pdfium() {
    for name in [
        "pdf_scanned",
        "pdf_mixed",
        "pdf_table_scan",
        "pdf_rot90_scan",
        "pdf_cjk_scan",
    ] {
        let path = format!("tests/fixtures/{name}.pdf");
        let bytes = std::fs::read(&path).unwrap_or_else(|e| panic!("{path}: {e}"));
        let doc = office_parser::pdf::parse(&bytes).expect("parses");
        let pages = doc
            .metadata
            .page_quality
            .unwrap_or_else(|| panic!("{path}: no page_quality"));
        assert!(!pages.is_empty(), "{path}: verdict vector is empty");
        for q in &pages {
            if let PageVerdict::Ocr { reason } = &q.verdict {
                assert!(!reason.is_empty(), "{path} p{}: empty untrust reason", q.page);
            }
        }
    }
}

/// force bypasses every trust check (spec scenario): with `force = true` no
/// page carries a no-ink verdict. The corpus itself is geometry-consistent,
/// so the load-bearing force assertion lives in the in-crate wiring test
/// (`force_bypasses_the_no_ink_wiring`) on a hand-built no-ink page; this
/// end-to-end pass guards the routing path (force + engine OCRs all pages
/// regardless of geometry).
#[test]
fn force_produces_no_no_ink_verdict() {
    for name in ["pdf_scanned", "pdf_mixed"] {
        let path = format!("tests/fixtures/{name}.pdf");
        let bytes = std::fs::read(&path).unwrap_or_else(|e| panic!("{path}: {e}"));
        let (doc, _) = parse_pdf_split(&bytes, None, 150, 50, true).expect("split parse");
        let pages = doc.metadata.page_quality.expect("page quality");
        for q in &pages {
            if let PageVerdict::Ocr { reason } = &q.verdict {
                assert_ne!(reason, NO_INK_REASON, "{path} p{}", q.page);
            }
        }
    }
}
