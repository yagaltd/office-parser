#[test]
fn docx_multicolumn_section_emits_morph_columns_fence() {
    let bytes = std::fs::read("tests/fixtures/docx_columns.docx").unwrap();
    let doc = office_parser::parse(&bytes, "in.docx").unwrap();
    let md = office_parser::render::to_morph_markdown(&doc);

    // intro stays outside; the 2-col section is fenced
    assert!(md.starts_with("Intro before columns."), "{md}");
    let fence_start = md.find("::: columns").expect("fence missing");
    assert!(fence_start > 0);
    assert_eq!(md.matches("::: column\n").count(), 2, "{md}");
    // the two section paragraphs land in separate column groups
    let col_a = md.find("Left column content").unwrap();
    let col_b = md.find("Right column content").unwrap();
    let group_split = md.find(":::\n::: column\n").expect("column separator");
    assert!(col_a > fence_start && col_a < group_split, "{md}");
    assert!(col_b > group_split, "{md}");
    // no content after the fence
    let fence_end = md.rfind(":::\n").unwrap();
    assert!(md[fence_end + 4..].trim().is_empty(), "{md}");
}
