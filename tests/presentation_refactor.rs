use anyhow::Result;

fn zip_bytes(entries: &[(&str, &[u8])]) -> Result<Vec<u8>> {
    use std::io::Write;

    let mut out: Vec<u8> = Vec::new();
    {
        let cur = std::io::Cursor::new(&mut out);
        let mut z = zip::ZipWriter::new(cur);
        let opt = zip::write::FileOptions::default();
        for (name, bytes) in entries {
            z.start_file(*name, opt)?;
            z.write_all(bytes)?;
        }
        z.finish()?;
    }
    Ok(out)
}

fn tiny_png_1x1() -> Vec<u8> {
    vec![
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F,
        0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00,
        0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49,
        0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ]
}

#[test]
fn pptx_diagram_graph_contract_includes_mermaid_block() -> Result<()> {
    let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldSz cx="9144000" cy="6858000"/>
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
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="1" name="Box1"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>
          <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody><a:p><a:r><a:t>Start</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Box2"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="3000000" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>
          <a:prstGeom prst="diamond"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody><a:p><a:r><a:t>Process</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:cxnSp>
        <p:nvCxnSpPr><p:cNvPr id="10" name="Conn"/></p:nvCxnSpPr>
        <p:spPr>
          <a:cxnSpPr>
            <a:stCxn id="1"/>
            <a:endCxn id="2"/>
          </a:cxnSpPr>
        </p:spPr>
      </p:cxnSp>
    </p:spTree>
  </p:cSld>
</p:sld>
"#;

    let slide1_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
"#;

    let pptx = zip_bytes(&[
        ("ppt/presentation.xml", presentation_xml.as_bytes()),
        (
            "ppt/_rels/presentation.xml.rels",
            presentation_rels.as_bytes(),
        ),
        ("ppt/slides/slide1.xml", slide1_xml.as_bytes()),
        ("ppt/slides/_rels/slide1.xml.rels", slide1_rels.as_bytes()),
    ])?;

    let doc = office_parser::pptx::parse(&pptx)?;

    let slides = doc
        .metadata
        .extra
        .get("slides")
        .and_then(|v| v.as_array())
        .expect("slides");
    assert_eq!(slides.len(), 1);
    assert!(
        slides[0]
            .get("score")
            .and_then(|v| v.as_f64())
            .unwrap_or(0.0)
            > 0.0
    );

    let graphs = doc
        .metadata
        .extra
        .get("diagram_graphs")
        .and_then(|v| v.as_array())
        .expect("diagram_graphs");
    assert_eq!(graphs.len(), 1);
    let edges = graphs[0]
        .get("edges")
        .and_then(|v| v.as_array())
        .expect("edges");
    assert!(edges.iter().any(|e| {
        e.get("from").and_then(|v| v.as_str()) == Some("n1")
            && e.get("to").and_then(|v| v.as_str()) == Some("n2")
    }));

    assert!(doc.blocks.iter().any(|b| {
        matches!(b, office_parser::document_ast::Block::Heading { level: 3, text, .. } if text == "Diagram")
    }));

    assert!(doc.blocks.iter().any(|b| {
        matches!(b, office_parser::document_ast::Block::Paragraph { text, .. } if text.contains("```mermaid") && text.contains("flowchart"))
    }));

    assert!(doc.blocks.iter().any(|b| {
        matches!(b, office_parser::document_ast::Block::Paragraph { text, .. } if text.contains("shape: circle") && text.contains("shape: diamond"))
    }));
    Ok(())
}

#[test]
fn pptx_chart_data_json_contract() -> Result<()> {
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
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <p:cSld>
    <p:spTree>
      <p:graphicFrame>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart r:id="rIdChart1"/>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
</p:sld>
"#;

    let slide1_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>
"#;

    let chart1_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:p><a:r><a:t>Revenue</a:t></a:r></a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:tx><c:v>2025</c:v></c:tx>
          <c:cat>
            <c:strRef>
              <c:strCache>
                <c:pt idx="0"><c:v>Q1</c:v></c:pt>
                <c:pt idx="1"><c:v>Q2</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:pt idx="0"><c:v>10.2</c:v></c:pt>
                <c:pt idx="1"><c:v>11.3</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:ser>
          <c:tx><c:v>2026</c:v></c:tx>
          <c:cat>
            <c:strRef>
              <c:strCache>
                <c:pt idx="0"><c:v>Q1</c:v></c:pt>
                <c:pt idx="1"><c:v>Q2</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:pt idx="0"><c:v>20.2</c:v></c:pt>
                <c:pt idx="1"><c:v>21.3</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
"#;

    let pptx = zip_bytes(&[
        ("ppt/presentation.xml", presentation_xml.as_bytes()),
        (
            "ppt/_rels/presentation.xml.rels",
            presentation_rels.as_bytes(),
        ),
        ("ppt/slides/slide1.xml", slide1_xml.as_bytes()),
        ("ppt/slides/_rels/slide1.xml.rels", slide1_rels.as_bytes()),
        ("ppt/charts/chart1.xml", chart1_xml.as_bytes()),
    ])?;

    let doc = office_parser::pptx::parse(&pptx)?;

    let charts = doc
        .metadata
        .extra
        .get("charts")
        .and_then(|v| v.as_array())
        .expect("charts");
    assert_eq!(charts.len(), 1);
    assert_eq!(charts[0].get("slide").and_then(|v| v.as_u64()), Some(1));
    assert_eq!(
        charts[0].get("chart_type").and_then(|v| v.as_str()),
        Some("bar")
    );
    assert_eq!(
        charts[0].get("title").and_then(|v| v.as_str()),
        Some("Revenue")
    );
    let cats = charts[0]
        .get("categories")
        .and_then(|v| v.as_array())
        .unwrap();
    assert_eq!(
        cats.iter().filter_map(|v| v.as_str()).collect::<Vec<_>>(),
        vec!["Q1", "Q2"]
    );

    let chart_table = doc
        .blocks
        .iter()
        .find_map(|b| match b {
            office_parser::document_ast::Block::Table { rows, .. } => Some(rows),
            _ => None,
        })
        .expect("chart table block");
    assert_eq!(chart_table.len(), 1 + 2);
    assert_eq!(chart_table[0].len(), 1 + 2);
    assert_eq!(chart_table[1].len(), 1 + 2);
    assert_eq!(chart_table[2].len(), 1 + 2);

    assert!(doc.blocks.iter().any(|b| {
        matches!(b, office_parser::document_ast::Block::Paragraph { text, .. } if text.contains("Note:") && text.contains("bar chart"))
    }));

    Ok(())
}

#[test]
fn odp_diagram_graph_contract_includes_mermaid_block() -> Result<()> {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
  <office:styles>
    <style:page-layout style:name="pm1">
      <style:page-layout-properties fo:page-width="28cm" fo:page-height="21cm"/>
    </style:page-layout>
  </office:styles>
</office:document-styles>
"#;

    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">
  <office:body>
    <office:presentation>
      <draw:page draw:name="Slide 1">
        <draw:frame draw:id="f1" svg:x="1cm" svg:y="1cm" svg:width="6cm" svg:height="2cm">
          <draw:text-box><text:p>Start</text:p></draw:text-box>
        </draw:frame>
        <draw:frame draw:id="f2" svg:x="12cm" svg:y="1cm" svg:width="6cm" svg:height="2cm">
          <draw:text-box><text:p>End</text:p></draw:text-box>
        </draw:frame>
        <draw:connector draw:start-shape="f1" draw:end-shape="f2" svg:x1="7cm" svg:y1="2cm" svg:x2="12cm" svg:y2="2cm" draw:marker-end="Arrow"/>
      </draw:page>
    </office:presentation>
  </office:body>
</office:document-content>
"#;

    let odp = zip_bytes(&[
        ("content.xml", content_xml.as_bytes()),
        ("styles.xml", styles_xml.as_bytes()),
    ])?;
    let doc = office_parser::odp::parse(&odp)?;

    let slides = doc
        .metadata
        .extra
        .get("slides")
        .and_then(|v| v.as_array())
        .expect("slides");

    let graphs = doc
        .metadata
        .extra
        .get("diagram_graphs")
        .and_then(|v| v.as_array())
        .expect("diagram_graphs");
    assert_eq!(graphs.len(), 1);
    let edges = graphs[0].get("edges").and_then(|v| v.as_array()).unwrap();
    assert!(edges.iter().any(|e| {
        e.get("from").and_then(|v| v.as_str()) == Some("n1")
            && e.get("to").and_then(|v| v.as_str()) == Some("n2")
    }));

    assert!(doc.blocks.iter().any(|b| {
        matches!(b, office_parser::document_ast::Block::Heading { level: 3, text, .. } if text == "Diagram")
    }));
    assert!(doc.blocks.iter().any(|b| {
        matches!(b, office_parser::document_ast::Block::Paragraph { text, .. } if text.contains("```mermaid") && text.contains("flowchart"))
    }));

    assert!(
        doc.images
            .iter()
            .all(|img| img.source_ref.as_deref() != Some("slide:1:snapshot"))
    );

    // Connectors are now extracted into Mermaid/graph JSON, not treated as excluded visuals.
    let excluded = slides[0]
        .get("excluded_visuals")
        .and_then(|v| v.as_array())
        .unwrap();
    assert!(!excluded.iter().any(|v| v.as_str() == Some("connectors")));
    Ok(())
}

#[test]
fn pptx_parse_options_do_not_emit_slide_snapshots() -> Result<()> {
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
  <p:cSld>
    <p:spTree>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="9" name="Img"/></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rIdImg1"/></p:blipFill>
      </p:pic>
    </p:spTree>
  </p:cSld>
</p:sld>
"#;
    let slide1_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
"#;
    let png = tiny_png_1x1();
    let pptx = zip_bytes(&[
        ("ppt/presentation.xml", presentation_xml.as_bytes()),
        (
            "ppt/_rels/presentation.xml.rels",
            presentation_rels.as_bytes(),
        ),
        ("ppt/slides/slide1.xml", slide1_xml.as_bytes()),
        ("ppt/slides/_rels/slide1.xml.rels", slide1_rels.as_bytes()),
        ("ppt/media/image1.png", png.as_slice()),
    ])?;

    let with_false = office_parser::pptx::parse_with_options(
        &pptx,
        office_parser::pptx::ParseOptions {
            include_slide_snapshots: false,
        },
    )?;
    let with_true = office_parser::pptx::parse_with_options(
        &pptx,
        office_parser::pptx::ParseOptions {
            include_slide_snapshots: true,
        },
    )?;

    for doc in [&with_false, &with_true] {
        assert!(doc.images.iter().all(|img| {
            !matches!(
                img.source_ref.as_deref(),
                Some(src) if src.ends_with(":snapshot")
            )
        }));
        assert!(!doc.images.iter().any(|img| {
            matches!(
                img.filename.as_deref(),
                Some(name) if name.starts_with("slide_") && name.ends_with(".png")
            )
        }));
    }
    Ok(())
}
