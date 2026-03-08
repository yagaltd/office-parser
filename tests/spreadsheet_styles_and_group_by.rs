use anyhow::Result;

#[cfg(feature = "xlsx")]
fn build_minimal_xlsx_with_datetime() -> Result<Vec<u8>> {
    use std::io::Write;
    use zip::write::FileOptions;

    let mut buf = Vec::new();
    {
        let mut w = zip::ZipWriter::new(std::io::Cursor::new(&mut buf));
        let opt = FileOptions::default();

        w.start_file("[Content_Types].xml", opt)?;
        w.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#,
        )?;

        w.add_directory("_rels/", opt)?;
        w.start_file("_rels/.rels", opt)?;
        w.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
        )?;

        w.add_directory("xl/", opt)?;
        w.add_directory("xl/_rels/", opt)?;
        w.add_directory("xl/worksheets/", opt)?;

        w.start_file("xl/workbook.xml", opt)?;
        w.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#,
        )?;

        w.start_file("xl/_rels/workbook.xml.rels", opt)?;
        w.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
        )?;

        w.start_file("xl/styles.xml", opt)?;
        w.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>"#,
        )?;

        w.start_file("xl/worksheets/sheet1.xml", opt)?;
        w.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>x</t></is></c>
    </row>
    <row r="2">
      <c r="A2" s="1"><v>44128.76593707176</v></c>
    </row>
  </sheetData>
</worksheet>"#,
        )?;

        w.finish()?;
    }
    Ok(buf)
}

#[cfg(feature = "xlsx")]
#[test]
fn xlsx_applies_style_aware_datetime_formatting() -> Result<()> {
    let xlsx = build_minimal_xlsx_with_datetime()?;
    let doc = office_parser::xlsx::parse(&xlsx)?;
    let md = office_parser::render::to_markdown(&doc);
    // Calamine may emit either ISO-with-T or space-separated datetime; accept either.
    assert!(md.contains("2020-10-23"));
    assert!(md.contains("18:22"));
    assert!(!md.contains("44128.765937"));
    Ok(())
}

#[test]
fn csv_group_by_splits_segments_on_key_change() -> Result<()> {
    let csv = b"user,amount\nalice,10\nalice,11\nbob,5\nbob,6\n";
    let opts = office_parser::spreadsheet::ParseOptions {
        group_by: Some("user".to_string()),
        max_table_rows_per_segment: 10,
        ..Default::default()
    };
    let doc = office_parser::csv::parse_with_options(csv, opts)?;
    let md = office_parser::render::to_markdown(&doc);
    assert!(md.contains("user: alice"));
    assert!(md.contains("user: bob"));
    Ok(())
}
