use office_parser::document_ast::{Block, MarkKind};

#[test]
fn docx_run_styling_becomes_inline_marks() {
    let bytes = std::fs::read("tests/fixtures/docx_bold_italic.docx").unwrap();
    let doc = office_parser::parse(&bytes, "in.docx").unwrap();

    let para = doc
        .blocks
        .iter()
        .find_map(|b| match b {
            Block::Paragraph { text, marks, .. } if text.starts_with("Plain ") => {
                Some((text.as_str(), marks))
            }
            _ => None,
        })
        .expect("styled paragraph");

    let (text, marks) = para;
    assert_eq!(text, "Plain bold and italic gone");
    assert_eq!(marks.len(), 3, "marks: {marks:?}");
    assert_eq!(marks[0].kind, MarkKind::Bold);
    assert_eq!((marks[0].start, marks[0].end), (6, 10));
    assert_eq!(marks[1].kind, MarkKind::Italic);
    assert_eq!((marks[1].start, marks[1].end), (15, 21));
    assert_eq!(marks[2].kind, MarkKind::Strikethrough);
    assert_eq!((marks[2].start, marks[2].end), (21, 26));

    // hyperlink becomes a Link mark at the right range
    let linkpara = doc
        .blocks
        .iter()
        .find_map(|b| match b {
            Block::Paragraph { text, marks, .. } if text.starts_with("Visit ") => {
                Some((text.as_str(), marks))
            }
            _ => None,
        })
        .expect("hyperlink paragraph");
    assert_eq!(linkpara.0, "Visit example now");
    assert_eq!(linkpara.1.len(), 1, "{:?}", linkpara.1);
    match &linkpara.1[0].kind {
        MarkKind::Link { url } => assert_eq!(url, "https://example.com"),
        other => panic!("expected link mark, got {other:?}"),
    }
    assert_eq!((linkpara.1[0].start, linkpara.1[0].end), (6, 13));
}

#[test]
fn morph_renderer_serializes_docx_marks() {
    let bytes = std::fs::read("tests/fixtures/docx_bold_italic.docx").unwrap();
    let doc = office_parser::parse(&bytes, "in.docx").unwrap();
    let md = office_parser::render::to_morph_markdown(&doc);
    assert!(
        md.contains("Plain **bold** and *italic*~~ gone~~"),
        "morph md: {md}"
    );
    assert!(
        md.contains("Visit [example](https://example.com) now"),
        "link md: {md}"
    );
}

#[test]
fn docx_table_widths_drive_morph_wrapper() {
    let bytes = std::fs::read("tests/fixtures/docx_table_widths.docx").unwrap();
    let doc = office_parser::parse(&bytes, "in.docx").unwrap();
    let md = office_parser::render::to_morph_markdown(&doc);
    assert!(
        md.contains("::: table {widths=[3000,1000]}\n| Name | Qty |\n|---|---|\n| Widget | 4 |\n:::"),
        "wrapper md: {md}"
    );
}
