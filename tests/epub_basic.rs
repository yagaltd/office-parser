use anyhow::Result;

fn zip_bytes(entries: &[(&str, &[u8])]) -> Result<Vec<u8>> {
    use std::io::Write;
    use zip::write::FileOptions;

    let cur = std::io::Cursor::new(Vec::<u8>::new());
    let mut w = zip::ZipWriter::new(cur);
    let opts = FileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, bytes) in entries {
        w.start_file(*name, opts)?;
        w.write_all(bytes)?;
    }
    Ok(w.finish()?.into_inner())
}

#[test]
#[cfg(feature = "epub")]
fn epub_parses_spine_xhtml_and_images() -> Result<()> {
    let container_xml = r#"<?xml version="1.0"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>
"#;

    let opf = r#"<?xml version="1.0" encoding="UTF-8"?>
<package version="3.0" xmlns="http://www.idpf.org/2007/opf" unique-identifier="BookId">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:title>Test Book</dc:title>
  </metadata>
  <manifest>
    <item id="ch1" href="ch1.xhtml" media-type="application/xhtml+xml"/>
    <item id="ch2" href="ch2.xhtml" media-type="application/xhtml+xml"/>
    <item id="img1" href="img/cover.png" media-type="image/png"/>
  </manifest>
  <spine>
    <itemref idref="ch1"/>
    <itemref idref="ch2"/>
  </spine>
</package>
"#;

    let ch1 = r#"<?xml version="1.0" encoding="UTF-8"?>
<html xmlns="http://www.w3.org/1999/xhtml">
  <head><title>Chapter One</title></head>
  <body>
    <h1>Intro</h1>
    <p>Hello <b>world</b>.</p>
    <img src="img/cover.png" alt="cover"/>
  </body>
</html>
"#;

    let ch2 = r#"<?xml version="1.0" encoding="UTF-8"?>
<html xmlns="http://www.w3.org/1999/xhtml">
  <head><title>Chapter Two</title></head>
  <body>
    <p>Second.</p>
    <ul><li>a</li><li>b</li></ul>
  </body>
</html>
"#;

    let png: &[u8] = &[
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x04, 0x00, 0x00, 0x00, 0xb5,
        0x1c, 0x0c, 0x02, 0x00, 0x00, 0x00, 0x0b, 0x49, 0x44, 0x41, 0x54, 0x78, 0xda, 0x63, 0xfc,
        0xff, 0x1f, 0x00, 0x03, 0x03, 0x01, 0xff, 0xaa, 0xb3, 0xd2, 0x00, 0x00, 0x00, 0x00, 0x49,
        0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ];

    let epub = zip_bytes(&[
        ("META-INF/container.xml", container_xml.as_bytes()),
        ("OEBPS/content.opf", opf.as_bytes()),
        ("OEBPS/ch1.xhtml", ch1.as_bytes()),
        ("OEBPS/ch2.xhtml", ch2.as_bytes()),
        ("OEBPS/img/cover.png", png),
    ])?;

    let doc = office_parser::epub::parse(&epub)?;
    assert_eq!(doc.metadata.format, office_parser::Format::Epub);
    assert_eq!(doc.metadata.title.as_deref(), Some("Test Book"));

    let md = office_parser::render::to_markdown(&doc);
    assert!(md.contains("Chapter One"));
    assert!(md.contains("Hello world"));
    assert!(md.contains("- a"));
    assert!(md.contains("[image:sha256:"));

    assert_eq!(doc.images.len(), 1);
    assert_eq!(doc.images[0].mime_type, "image/png");
    Ok(())
}
