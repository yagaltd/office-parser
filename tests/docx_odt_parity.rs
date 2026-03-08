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

#[test]
fn docx_chart_note_includes_type_and_units_and_diagram_mermaid() -> Result<()> {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart r:id="rIdChart1"/>
            </a:graphicData>
          </a:graphic>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wps:wsp>
            <a:cNvPr id="1" name="Box1"/>
            <a:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>
              <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
            </a:spPr>
            <wps:txbx>
              <w:txbxContent>
                <w:p><w:r><w:t>Start</w:t></w:r></w:p>
              </w:txbxContent>
            </wps:txbx>
          </wps:wsp>
          <wps:wsp>
            <a:cNvPr id="2" name="Box2"/>
            <a:spPr>
              <a:xfrm><a:off x="3000000" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>
              <a:prstGeom prst="diamond"><a:avLst/></a:prstGeom>
            </a:spPr>
            <wps:txbx>
              <w:txbxContent>
                <w:p><w:r><w:t>End</w:t></w:r></w:p>
              </w:txbxContent>
            </wps:txbx>
          </wps:wsp>
          <a:cxnSp>
            <a:cxnSpPr>
              <a:stCxn id="1"/>
              <a:endCxn id="2"/>
            </a:cxnSpPr>
          </a:cxnSp>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>
"#;

    let rels_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
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
      <c:pieChart>
        <c:ser>
          <c:tx><c:v>Share</c:v></c:tx>
          <c:cat>
            <c:strRef>
              <c:strCache>
                <c:pt idx="0"><c:v>A</c:v></c:pt>
                <c:pt idx="1"><c:v>B</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:formatCode>0%</c:formatCode>
                <c:pt idx="0"><c:v>10</c:v></c:pt>
                <c:pt idx="1"><c:v>20</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
"#;

    let docx = zip_bytes(&[
        ("word/document.xml", document_xml.as_bytes()),
        ("word/_rels/document.xml.rels", rels_xml.as_bytes()),
        ("word/charts/chart1.xml", chart1_xml.as_bytes()),
    ])?;

    let doc = office_parser::docx::parse(&docx)?;

    assert!(doc.blocks.iter().any(|b| {
        matches!(b, office_parser::document_ast::Block::Paragraph { text, .. } if text.contains("Note:") && text.contains("pie chart") && text.contains("units"))
    }));

    assert!(doc.blocks.iter().any(|b| {
        matches!(b, office_parser::document_ast::Block::Table { rows, .. } if rows.iter().flatten().any(|c| c.text.contains('%')))
    }));

    assert!(doc.blocks.iter().any(|b| {
        matches!(b, office_parser::document_ast::Block::Paragraph { text, .. } if text.contains("```mermaid") && text.contains("flowchart") && text.contains("shape: circle") && text.contains("shape: diamond"))
    }));

    let charts = doc
        .metadata
        .extra
        .get("charts")
        .and_then(|v| v.as_array())
        .expect("charts");
    assert_eq!(charts.len(), 1);

    let graphs = doc
        .metadata
        .extra
        .get("diagram_graphs")
        .and_then(|v| v.as_array())
        .expect("diagram_graphs");
    assert_eq!(graphs.len(), 1);

    Ok(())
}

#[test]
fn odt_embedded_chart_object_and_diagram_mermaid() -> Result<()> {
    let mimetype = "application/vnd.oasis.opendocument.text";

    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:text>
      <text:p>
        <draw:frame draw:id="f1" svg:x="1cm" svg:y="1cm" svg:width="5cm" svg:height="2cm">
          <draw:text-box><text:p>Start</text:p></draw:text-box>
        </draw:frame>
        <draw:frame draw:id="f2" svg:x="10cm" svg:y="1cm" svg:width="5cm" svg:height="2cm">
          <draw:text-box><text:p>End</text:p></draw:text-box>
        </draw:frame>
        <draw:connector draw:start-shape="f1" draw:end-shape="f2" svg:x1="6cm" svg:y1="2cm" svg:x2="10cm" svg:y2="2cm" draw:marker-end="Arrow"/>
        <draw:object xlink:href="./Object 1"/>
      </text:p>
    </office:text>
  </office:body>
</office:document-content>
"#;

    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>
"#;

    let object1_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0">
  <office:body>
    <office:chart>
      <chart:chart chart:class="chart:pie">
        <chart:title><text:p>Share</text:p></chart:title>
        <table:table>
          <table:table-row>
            <table:table-cell><text:p>series</text:p></table:table-cell>
            <table:table-cell><text:p>A</text:p></table:table-cell>
            <table:table-cell><text:p>B</text:p></table:table-cell>
          </table:table-row>
          <table:table-row>
            <table:table-cell><text:p>Share</text:p></table:table-cell>
            <table:table-cell><text:p>10%</text:p></table:table-cell>
            <table:table-cell><text:p>20%</text:p></table:table-cell>
          </table:table-row>
        </table:table>
      </chart:chart>
    </office:chart>
  </office:body>
</office:document-content>
"#;

    let odt = zip_bytes(&[
        ("mimetype", mimetype.as_bytes()),
        ("content.xml", content_xml.as_bytes()),
        ("styles.xml", styles_xml.as_bytes()),
        ("Object 1/content.xml", object1_xml.as_bytes()),
    ])?;

    let doc = office_parser::odt::parse(&odt)?;

    assert!(doc.blocks.iter().any(|b| {
        matches!(b, office_parser::document_ast::Block::Paragraph { text, .. } if text.contains("```mermaid") && text.contains("flowchart"))
    }));
    assert!(doc.blocks.iter().any(|b| {
        matches!(b, office_parser::document_ast::Block::Paragraph { text, .. } if text.contains("Note:") && text.contains("pie chart") && text.contains("units: %"))
    }));
    assert!(doc.blocks.iter().any(|b| {
        matches!(b, office_parser::document_ast::Block::Table { rows, .. } if rows.iter().flatten().any(|c| c.text.contains('%')))
    }));

    let charts = doc
        .metadata
        .extra
        .get("charts")
        .and_then(|v| v.as_array())
        .expect("charts");
    assert_eq!(charts.len(), 1);

    let graphs = doc
        .metadata
        .extra
        .get("diagram_graphs")
        .and_then(|v| v.as_array())
        .expect("diagram_graphs");
    assert_eq!(graphs.len(), 1);

    Ok(())
}
