use anyhow::Result;

#[test]
fn xml_parses_wp_like_and_converts_cdata_html_to_markdown() -> Result<()> {
    let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<rss version="2.0">
  <channel>
    <title>example</title>
    <item>
      <title>Hello</title>
      <guid isPermaLink="false">x</guid>
      <content:encoded><![CDATA[<!-- wp:paragraph --><p>Hi <b>there</b>.</p><!-- /wp:paragraph -->]]></content:encoded>
    </item>
  </channel>
</rss>
"#;

    let doc = office_parser::xml::parse(xml)?;
    assert_eq!(doc.metadata.format, office_parser::Format::Xml);

    let md = office_parser::render::to_markdown(&doc);
    assert!(md.contains("# rss"));
    assert!(md.contains("version: 2.0") || md.contains("@version: 2.0"));
    assert!(md.contains("title: example"));
    assert!(md.contains("Hi there."), "md was:\n{md}");

    Ok(())
}
