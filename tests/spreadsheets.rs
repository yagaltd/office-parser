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

fn xlsx_minimal_bytes() -> Result<Vec<u8>> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
</Types>
"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#;

    let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Second" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>
"#;

    let sheet1 = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>id</t></is></c>
      <c r="B1" t="inlineStr"><is><t>name</t></is></c>
    </row>
    <row r="2">
      <c r="A2"><v>1</v></c>
      <c r="B2" t="inlineStr"><is><t>Alice</t></is></c>
    </row>
    <row r="3">
      <c r="A3"><v>2</v></c>
      <c r="B3" t="inlineStr"><is><t>Bob</t></is></c>
    </row>
    <row r="4"></row>
    <row r="5">
      <c r="A5" t="inlineStr"><is><t>key</t></is></c>
      <c r="B5" t="inlineStr"><is><t>value</t></is></c>
    </row>
    <row r="6">
      <c r="A6" t="inlineStr"><is><t>x</t></is></c>
      <c r="B6"><v>42</v></c>
    </row>
  </sheetData>
  <tableParts count="1">
    <tablePart r:id="rIdTable1"/>
  </tableParts>
</worksheet>
"#;

    let sheet2 = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>a</t></is></c>
      <c r="B1" t="inlineStr"><is><t>b</t></is></c>
      <c r="D1" t="inlineStr"><is><t>x</t></is></c>
      <c r="E1" t="inlineStr"><is><t>y</t></is></c>
    </row>
    <row r="2">
      <c r="A2"><v>3</v></c>
      <c r="B2"><v>4</v></c>
      <c r="D2"><v>30</v></c>
      <c r="E2"><v>40</v></c>
    </row>
  </sheetData>
</worksheet>
"#;

    let sheet1_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>
"#;

    let table1 = r#"<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="People" displayName="People" ref="A1:B3" headerRowCount="1">
  <tableColumns count="2">
    <tableColumn id="1" name="id"/>
    <tableColumn id="2" name="name"/>
  </tableColumns>
</table>
"#;

    zip_bytes(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", rels.as_bytes()),
        ("xl/workbook.xml", workbook.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/worksheets/sheet1.xml", sheet1.as_bytes()),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            sheet1_rels.as_bytes(),
        ),
        ("xl/worksheets/sheet2.xml", sheet2.as_bytes()),
        ("xl/tables/table1.xml", table1.as_bytes()),
    ])
}

fn ods_minimal_bytes() -> Result<Vec<u8>> {
    use std::io::Write;

    let mimetype = b"application/vnd.oasis.opendocument.spreadsheet";
    let content = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2"><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell office:value-type="string"><text:p>id</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>name</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>Alice</text:p></table:table-cell></table:table-row><table:table-row/><table:table-row><table:table-cell office:value-type="string"><text:p>key</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>value</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="42"><text:p>42</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;

    let manifest = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"#;

    let mut out: Vec<u8> = Vec::new();
    {
        let cur = std::io::Cursor::new(&mut out);
        let mut z = zip::ZipWriter::new(cur);

        let opt_store =
            zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Stored);
        z.start_file("mimetype", opt_store)?;
        z.write_all(mimetype)?;

        let opt = zip::write::FileOptions::default();
        z.start_file("content.xml", opt)?;
        z.write_all(content.as_bytes())?;
        z.start_file("META-INF/manifest.xml", opt)?;
        z.write_all(manifest.as_bytes())?;

        z.finish()?;
    }
    Ok(out)
}

#[test]
fn xlsx_parses_sheets_and_regions() -> Result<()> {
    let bytes = xlsx_minimal_bytes()?;
    let doc = office_parser::xlsx::parse(&bytes)?;
    assert_eq!(doc.metadata.format, office_parser::Format::Xlsx);

    let md = office_parser::render::to_markdown(&doc);
    assert!(md.contains("## Sheet: Sheet1"));
    assert!(md.contains("## Sheet: Second"));
    assert!(md.contains("### Table: People"));
    assert!(md.contains("### Table 1"));
    assert!(md.contains("### Table 2"));
    assert!(md.contains("| id | name |"));
    assert!(md.contains("| key | value |"));
    Ok(())
}

#[test]
fn ods_parses_regions() -> Result<()> {
    let bytes = ods_minimal_bytes()?;
    let doc = office_parser::ods::parse(&bytes)?;
    assert_eq!(doc.metadata.format, office_parser::Format::Ods);
    let md = office_parser::render::to_markdown(&doc);
    assert!(md.contains("## Sheet: Sheet1"));
    assert!(md.contains("### Table 1"));
    assert!(md.contains("### Table 2"));
    assert!(md.contains("| id | name |"));
    assert!(md.contains("| key | value |"));
    Ok(())
}

#[test]
fn csv_tsv_parse_and_chunk_tables() -> Result<()> {
    let csv_bytes = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/people.csv"),
    )?;
    let doc = office_parser::csv::parse(&csv_bytes)?;
    let md = office_parser::render::to_markdown(&doc);
    assert!(md.contains("| id | name |"));

    let chunks = office_parser::render::to_chunks(&doc, 900);
    assert!(chunks.len() >= 2);
    for ch in chunks.iter().filter(|c| c.content.contains('|')) {
        assert!(ch.content.contains("| id | name |"));
    }

    let tsv_bytes = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/people.tsv"),
    )?;
    let doc = office_parser::tsv::parse(&tsv_bytes)?;
    let md = office_parser::render::to_markdown(&doc);
    assert!(md.contains("| id | name |"));
    Ok(())
}
