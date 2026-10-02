use office_parser::document_ast::{Block, MarkKind};

#[test]
fn pptx_run_attributes_become_inline_marks() {
    let bytes = std::fs::read("tests/fixtures/pptx_run_marks.pptx").unwrap();
    let doc = office_parser::parse(&bytes, "in.pptx").unwrap();

    // bold slide title text
    let bold = doc.blocks.iter().find_map(|b| match b {
        Block::Paragraph { text, marks, .. } if text == "Styled Run" => Some(marks.clone()),
        _ => None,
    });
    let bm = bold.expect("bold paragraph");
    assert_eq!(bm.len(), 1);
    assert_eq!(bm[0].kind, MarkKind::Bold);
    assert_eq!((bm[0].start, bm[0].end), (0, 10));

    // italic + strike on the second shape, same range
    let fancy = doc.blocks.iter().find_map(|b| match b {
        Block::Paragraph { text, marks, .. } if text == "plain fancy" => Some(marks.clone()),
        _ => None,
    });
    let fm = fancy.expect("fancy paragraph");
    assert_eq!(fm.len(), 2, "{fm:?}");
    assert!(fm.iter().all(|m| (m.start, m.end) == (6, 11)));
    assert!(fm.iter().any(|m| m.kind == MarkKind::Italic));
    assert!(fm.iter().any(|m| m.kind == MarkKind::Strikethrough));

    let md = office_parser::render::to_morph_markdown(&doc);
    assert!(md.contains("## Styled Run") || md.contains("# Styled Run"), "{md}");
    assert!(md.contains("plain ~~*fancy*~~") || md.contains("plain *~~fancy~~*"), "{md}");
}
