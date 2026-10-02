use office_parser::document_ast::{Block, MarkKind};

#[test]
fn odp_span_styles_become_inline_marks() {
    let bytes = std::fs::read("tests/fixtures/odp_span_marks.odp").unwrap();
    let doc = office_parser::parse(&bytes, "in.odp").unwrap();

    let para = doc
        .blocks
        .iter()
        .find_map(|b| match b {
            Block::Paragraph { text, marks, .. } if text.starts_with("body ") => {
                Some((text.as_str(), marks))
            }
            _ => None,
        })
        .expect("styled body paragraph");

    assert_eq!(para.0, "body bold bit end");
    assert_eq!(para.1.len(), 1, "{:?}", para.1);
    assert_eq!(para.1[0].kind, MarkKind::Bold);
    assert_eq!((para.1[0].start, para.1[0].end), (5, 13));

    let md = office_parser::render::to_morph_markdown(&doc);
    assert!(md.contains("body **bold bit** end"), "morph: {md}");
}
