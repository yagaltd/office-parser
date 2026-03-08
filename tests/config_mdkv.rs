use anyhow::Result;

#[test]
fn json_yaml_toml_emit_markdown_kv_with_max_depth_4_and_tables() -> Result<()> {
    let json = br#"{
  "a": {"b": {"c": {"d": {"e": 1, "arr": [{"x": 1, "y": "z"}, {"x": 2, "y": "w"}]}}}},
  "items": [{"id": 1, "name": "alice"}, {"id": 2, "name": "bob"}],
  "tags": ["x", "y"]
}"#;

    let doc = office_parser::json::parse(json)?;

    // Preserve source key order at the root (JSON object order).
    let md = office_parser::render::to_markdown(&doc);
    let i_service = md.find("# a").unwrap_or(usize::MAX);
    let i_items = md.find("# items").unwrap_or(usize::MAX);
    let i_tags = md.find("# tags").unwrap_or(usize::MAX);
    assert!(i_service < i_items && i_items < i_tags);

    // No headings deeper than ####
    assert!(!doc.blocks.iter().any(|b| {
        matches!(b, office_parser::document_ast::Block::Heading { level, .. } if *level > 4)
    }));

    // Expect at least one table for arrays-of-objects.
    assert!(doc.blocks.iter().any(|b| {
        matches!(b, office_parser::document_ast::Block::Table { rows, .. } if !rows.is_empty() && rows[0].iter().any(|c| c.text == "#") )
    }));

    // Scalar arrays should become a list.
    assert!(doc.blocks.iter().any(|b| {
        matches!(b, office_parser::document_ast::Block::List { items, .. } if items.iter().any(|it| it.text == "x") && items.iter().any(|it| it.text == "y"))
    }));

    let yaml = r#"
a:
  b:
    c:
      d:
        e: 1
        arr:
          - x: 1
            y: z
          - x: 2
            y: w
items:
  - id: 1
    name: alice
  - id: 2
    name: bob
tags:
  - x
  - y
"#;
    let doc = office_parser::yaml::parse(yaml.as_bytes())?;
    assert!(!doc.blocks.is_empty());
    assert!(
        doc.blocks
            .iter()
            .any(|b| matches!(b, office_parser::document_ast::Block::Table { .. }))
    );

    let md = office_parser::render::to_markdown(&doc);
    let i_service = md.find("# a").unwrap_or(usize::MAX);
    let i_items = md.find("# items").unwrap_or(usize::MAX);
    let i_tags = md.find("# tags").unwrap_or(usize::MAX);
    assert!(i_service < i_items && i_items < i_tags);

    let toml = r#"
[a.b.c.d]
e = 1

[[a.b.c.d.arr]]
x = 1
y = "z"

[[a.b.c.d.arr]]
x = 2
y = "w"

[[items]]
id = 1
name = "alice"

[[items]]
id = 2
name = "bob"

tags = ["x", "y"]
"#;
    let doc = office_parser::toml::parse(toml.as_bytes())?;
    assert!(!doc.blocks.is_empty());
    assert!(
        doc.blocks
            .iter()
            .any(|b| matches!(b, office_parser::document_ast::Block::Table { .. }))
    );

    let md = office_parser::render::to_markdown(&doc);
    let i_service = md.find("# a").unwrap_or(usize::MAX);
    let i_items = md.find("# items").unwrap_or(usize::MAX);
    let i_tags = md.find("# tags").unwrap_or(usize::MAX);
    assert!(i_service < i_items && i_items < i_tags);

    Ok(())
}
