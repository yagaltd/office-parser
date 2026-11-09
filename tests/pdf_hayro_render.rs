#![cfg(all(feature = "pdf", feature = "render-hayro"))]
//! The vendored hayro 0.9.0 path renders all five PDF fixtures non-blank.

const FIXTURES: [&str; 5] = [
    "tests/fixtures/pdf_scanned.pdf",
    "tests/fixtures/pdf_mixed.pdf",
    "tests/fixtures/pdf_table_scan.pdf",
    "tests/fixtures/pdf_rot90_scan.pdf",
    "tests/fixtures/pdf_cjk_scan.pdf",
];

/// test_hayro_renders_fixtures_non_blank: every fixture yields pages whose non-white pixel
/// ratio is > 0 through the hayro backend.
#[test]
fn test_hayro_renders_fixtures_non_blank() {
    for fixture in FIXTURES {
        let bytes = std::fs::read(fixture).expect("fixture");
        let pages = office_parser::formats::pdf::render_pages_png(&bytes, 150).expect("renders");
        assert!(!pages.is_empty(), "{fixture}: no pages");
        for (i, png) in pages.iter().enumerate() {
            let decoder = png::Decoder::new(std::io::Cursor::new(png));
            let mut reader = decoder.read_info().expect("png");
            let mut buf = vec![0; reader.output_buffer_size()];
            let info = reader.next_frame(&mut buf).expect("frame");
            let px = &buf[..info.buffer_size()];
            let non_white = px
                .chunks_exact(info.color_type.samples())
                .filter(|p| p.iter().any(|&c| c < 250))
                .count();
            assert!(
                non_white > 0,
                "{fixture} page {i}: rendered blank (non-white pixels: {non_white})"
            );
        }
    }
}
