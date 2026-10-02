use office_parser::document_ast::{Block, MarkKind};

#[test]
fn odt_span_styles_become_inline_marks() {
    let bytes = std::fs::read("tests/fixtures/odt_span_marks.odt").unwrap();
    let doc = office_parser::parse(&bytes, "in.odt").unwrap();

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

    assert_eq!(para.0, "Plain bold and italic.");
    assert_eq!(para.1.len(), 2, "{:?}", para.1);
    assert_eq!(para.1[0].kind, MarkKind::Bold);
    assert_eq!((para.1[0].start, para.1[0].end), (6, 10));
    assert_eq!(para.1[1].kind, MarkKind::Italic);
    assert_eq!((para.1[1].start, para.1[1].end), (15, 21));

    // parent-style inheritance: T3 inherits bold from T1, adds strike
    let stacked = doc
        .blocks
        .iter()
        .find_map(|b| match b {
            Block::Paragraph { text, marks, .. } if text.starts_with("Both ") => {
                Some(marks.clone())
            }
            _ => None,
        })
        .expect("stacked paragraph");
    let kinds: Vec<&MarkKind> = stacked.iter().map(|m| &m.kind).collect();
    assert!(kinds.contains(&&MarkKind::Bold), "{stacked:?}");
    assert!(kinds.contains(&&MarkKind::Strikethrough), "{stacked:?}");

    // morph renderer emits the markers
    let md = office_parser::render::to_morph_markdown(&doc);
    assert!(
        md.contains("Plain **bold** and *italic*."),
        "morph: {md}"
    );
    assert!(md.contains("Both **~~stacked~~**"), "stacked morph: {md}");
}
