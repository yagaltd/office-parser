use office_parser::document_ast::Block;

fn fixture_bytes(name: &str) -> Vec<u8> {
    let root = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let p = root.join("tests").join("fixtures").join(name);
    std::fs::read(p).expect("read fixture")
}

fn zip_bytes(entries: &[(&str, &[u8])]) -> anyhow::Result<Vec<u8>> {
    use std::io::Write;
    use zip::write::FileOptions;

    let cur = std::io::Cursor::new(Vec::<u8>::new());
    let mut w = zip::ZipWriter::new(cur);
    let opts = FileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, bytes) in entries {
        w.start_file(*name, opts)?;
        w.write_all(bytes)?;
    }
    let cur = w.finish()?;
    Ok(cur.into_inner())
}

#[test]
fn docx_semantic_blocks_and_images() {
    let bytes = fixture_bytes("docx_headings_lists_tables_images.docx");
    let doc = office_parser::docx::parse(&bytes).expect("parse docx");
    assert_eq!(doc.metadata.format, office_parser::Format::Docx);
    assert_eq!(doc.images.len(), 1);
    assert!(
        doc.blocks
            .iter()
            .any(|b| matches!(b, Block::Heading { .. }))
    );
    assert!(doc.blocks.iter().any(|b| matches!(b, Block::List { .. })));
    assert!(doc.blocks.iter().any(|b| matches!(b, Block::Table { .. })));
    assert!(doc.blocks.iter().any(|b| matches!(b, Block::Image { .. })));

    let md = office_parser::render::to_markdown(&doc);
    assert!(md.contains("[image:"));
}

#[test]
fn odt_semantic_blocks_and_images() {
    let bytes = fixture_bytes("odt_headings_lists_tables.odt");
    let doc = office_parser::odt::parse(&bytes).expect("parse odt");
    assert_eq!(doc.metadata.format, office_parser::Format::Odt);
    assert_eq!(doc.images.len(), 1);
    assert!(
        doc.blocks
            .iter()
            .any(|b| matches!(b, Block::Heading { .. }))
    );
    assert!(doc.blocks.iter().any(|b| matches!(b, Block::List { .. })));
    assert!(doc.blocks.iter().any(|b| matches!(b, Block::Table { .. })));
    assert!(doc.blocks.iter().any(|b| matches!(b, Block::Image { .. })));
}

#[test]
fn rtf_semantic_blocks_and_images() {
    let bytes = fixture_bytes("rtf_basic.rtf");
    let doc = office_parser::rtf::parse(&bytes).expect("parse rtf");
    assert_eq!(doc.metadata.format, office_parser::Format::Rtf);
    assert!(
        doc.blocks
            .iter()
            .any(|b| matches!(b, Block::Paragraph { .. }))
    );
}

#[test]
fn pptx_semantic_blocks_images_and_notes_synthetic() -> anyhow::Result<()> {
    let img: &[u8] = include_bytes!("fixtures/tiny.png");

    let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>
"#;

    let presentation_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>
"#;

    let slide1_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld name="Slide 1">
    <p:spTree>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="1" name="Title 1"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:txBody>
          <a:p><a:r><a:t>Hello Title</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="3" name="Picture 1" descr="alt text"/>
          <p:cNvPicPr/>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="rIdImg1"/>
        </p:blipFill>
      </p:pic>
    </p:spTree>
  </p:cSld>
</p:sld>
"#;

    let slide1_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rIdNotes1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
</Relationships>
"#;

    let notes1_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p><a:r><a:t>Note one</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>
"#;

    let pptx = zip_bytes(&[
        ("ppt/presentation.xml", presentation_xml.as_bytes()),
        (
            "ppt/_rels/presentation.xml.rels",
            presentation_rels.as_bytes(),
        ),
        ("ppt/slides/slide1.xml", slide1_xml.as_bytes()),
        ("ppt/slides/_rels/slide1.xml.rels", slide1_rels.as_bytes()),
        ("ppt/notesSlides/notesSlide1.xml", notes1_xml.as_bytes()),
        ("ppt/media/image1.png", img),
    ])?;

    let doc = office_parser::pptx::parse(&pptx).expect("parse pptx");
    assert_eq!(doc.metadata.format, office_parser::Format::Pptx);
    assert_eq!(doc.images.len(), 1);
    assert_eq!(doc.metadata.slide_count, Some(1));
    assert!(
        doc.blocks
            .iter()
            .any(|b| matches!(b, Block::Heading { .. }))
    );
    assert!(doc.blocks.iter().any(|b| matches!(b, Block::Image { .. })));

    let md = office_parser::render::to_markdown(&doc);
    assert!(md.contains("Hello Title"));
    assert!(md.contains("Note one"));
    Ok(())
}

#[test]
fn odp_semantic_blocks_and_images_synthetic() -> anyhow::Result<()> {
    let img: &[u8] = include_bytes!("fixtures/tiny.png");
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:presentation>
      <draw:page draw:name="Slide 1">
        <draw:frame presentation:class="title">
          <draw:text-box><text:p>My ODP Title</text:p></draw:text-box>
        </draw:frame>
        <draw:frame>
          <draw:image xlink:href="Pictures/image1.png"/>
        </draw:frame>
      </draw:page>
    </office:presentation>
  </office:body>
</office:document-content>
"#;

    let odp = zip_bytes(&[
        ("content.xml", content_xml.as_bytes()),
        ("Pictures/image1.png", img),
    ])?;
    let doc = office_parser::odp::parse(&odp).expect("parse odp");
    assert_eq!(doc.metadata.format, office_parser::Format::Odp);
    assert_eq!(doc.images.len(), 1);
    assert_eq!(doc.metadata.slide_count, Some(1));
    Ok(())
}

#[cfg(feature = "pdf")]
#[test]
fn pdf_page_count_is_set_synthetic() -> anyhow::Result<()> {
    use lopdf::content::{Content, Operation};
    use lopdf::dictionary;
    use lopdf::{Document, Object, Stream};

    let mut doc = Document::with_version("1.5");

    let pages_id = doc.new_object_id();
    let catalog_id = doc.new_object_id();
    let page_id = doc.new_object_id();
    let contents_id = doc.new_object_id();
    let font_id = doc.new_object_id();

    doc.objects.insert(
        font_id,
        Object::Dictionary(dictionary! {
            "Type" => "Font",
            "Subtype" => "Type1",
            "BaseFont" => "Helvetica",
        }),
    );

    let content = Content {
        operations: vec![
            Operation::new("BT", vec![]),
            Operation::new(
                "Tf",
                vec![Object::Name(b"F1".to_vec()), Object::Integer(12)],
            ),
            Operation::new("Td", vec![Object::Integer(50), Object::Integer(700)]),
            Operation::new("Tj", vec![Object::string_literal("Hello")]),
            Operation::new("ET", vec![]),
        ],
    };
    let stream = Stream::new(dictionary! {}, content.encode()?);
    doc.objects.insert(contents_id, Object::Stream(stream));

    let resources = dictionary! {
        "Font" => dictionary! { "F1" => font_id }
    };
    doc.objects.insert(
        page_id,
        Object::Dictionary(dictionary! {
            "Type" => "Page",
            "Parent" => pages_id,
            "MediaBox" => vec![0.into(), 0.into(), 612.into(), 792.into()],
            "Contents" => contents_id,
            "Resources" => resources,
        }),
    );

    doc.objects.insert(
        pages_id,
        Object::Dictionary(dictionary! {
            "Type" => "Pages",
            "Kids" => vec![page_id.into()],
            "Count" => 1,
        }),
    );

    doc.objects.insert(
        catalog_id,
        Object::Dictionary(dictionary! {
            "Type" => "Catalog",
            "Pages" => pages_id,
        }),
    );

    doc.trailer.set("Root", catalog_id);
    let mut bytes = Vec::new();
    doc.save_to(&mut bytes)?;

    let parsed = office_parser::pdf::parse(&bytes)?;
    assert_eq!(parsed.metadata.format, office_parser::Format::Pdf);
    assert_eq!(parsed.metadata.page_count, Some(1));
    Ok(())
}
