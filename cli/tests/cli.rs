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

fn fresh_temp_dir() -> std::path::PathBuf {
    use std::time::{SystemTime, UNIX_EPOCH};
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    std::env::temp_dir().join(format!(
        "office_parser_cli_test_{}_{}",
        std::process::id(),
        nanos
    ))
}

fn tiny_png_1x1() -> Vec<u8> {
    // 1x1 transparent PNG
    vec![
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F,
        0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00,
        0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49,
        0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ]
}

fn xlsx_minimal_bytes() -> Result<Vec<u8>> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
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
  </sheets>
</workbook>
"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"#;

    let sheet1 = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>id</t></is></c>
      <c r="B1" t="inlineStr"><is><t>name</t></is></c>
    </row>
    <row r="2">
      <c r="A2"><v>1</v></c>
      <c r="B2" t="inlineStr"><is><t>Alice</t></is></c>
    </row>
  </sheetData>
</worksheet>
"#;

    zip_bytes(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", rels.as_bytes()),
        ("xl/workbook.xml", workbook.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/worksheets/sheet1.xml", sheet1.as_bytes()),
    ])
}

fn ods_minimal_bytes() -> Result<Vec<u8>> {
    use std::io::Write;

    let mimetype = b"application/vnd.oasis.opendocument.spreadsheet";
    let content = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2"><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell office:value-type="string"><text:p>id</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>name</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>Alice</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;

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
fn cli_pptx_emits_mermaid_for_connectors_and_no_snapshots() -> Result<()> {
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
      <p:sp>
        <p:nvSpPr><p:cNvPr id="1" name="T"/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm></p:spPr>
        <p:txBody><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="9" name="Img"/></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rIdImg1"/></p:blipFill>
        <p:spPr><a:xfrm><a:off x="2500000" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm></p:spPr>
      </p:pic>
      <p:cxnSp>
        <p:nvCxnSpPr><p:cNvPr id="3" name="Connector"/></p:nvCxnSpPr>
        <p:spPr>
          <a:xfrm><a:off x="2000000" y="500000"/><a:ext cx="500000" cy="0"/></a:xfrm>
          <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
        </p:spPr>
        <a:stCxn id="1" idx="0"/>
        <a:endCxn id="9" idx="0"/>
      </p:cxnSp>
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

    let out_dir = fresh_temp_dir();
    std::fs::create_dir_all(&out_dir)?;
    let input_path = out_dir.join("in.pptx");
    std::fs::write(&input_path, &pptx)?;

    let exe = std::env::var("CARGO_BIN_EXE_office-parser-cli")
        .unwrap_or_else(|_| env!("CARGO_BIN_EXE_office-parser-cli").to_string());
    let status = std::process::Command::new(exe)
        .arg(&input_path)
        .arg("--out")
        .arg(&out_dir)
        .arg("--format")
        .arg("markdown")
        .status()?;
    assert!(status.success());

    let md = std::fs::read_to_string(out_dir.join("in.md"))?;
    assert!(md.contains("```mermaid"));
    assert!(md.contains("flowchart"));
    assert!(md.contains("asset/image1.png"));
    assert!(!md.contains("img:"));

    assert!(!out_dir.join("asset").join("slide_0001.png").exists());

    let _ = std::fs::remove_dir_all(&out_dir);
    Ok(())
}

#[test]
fn cli_marks_excluded_visuals_for_smartart() -> Result<()> {
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
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:graphicFrame>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>
        </a:graphic>
      </p:graphicFrame>
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

    let out_dir = fresh_temp_dir();
    std::fs::create_dir_all(&out_dir)?;
    let input_path = out_dir.join("smartart.pptx");
    std::fs::write(&input_path, &pptx)?;

    let exe = std::env::var("CARGO_BIN_EXE_office-parser-cli")
        .unwrap_or_else(|_| env!("CARGO_BIN_EXE_office-parser-cli").to_string());
    let status = std::process::Command::new(exe)
        .arg(&input_path)
        .arg("--out")
        .arg(&out_dir)
        .status()?;
    assert!(status.success());

    let md = std::fs::read_to_string(out_dir.join("smartart.md"))?;
    assert!(md.contains("SmartArt detected"));

    assert!(!out_dir.join("asset").join("slide_0001.png").exists());

    let _ = std::fs::remove_dir_all(&out_dir);
    Ok(())
}

#[test]
fn cli_config_json_emits_markdown_kv_and_tables() -> Result<()> {
    let json = r#"{
  "server": {"host": "localhost", "port": 8080},
  "users": [{"id": 1, "name": "alice"}, {"id": 2, "name": "bob"}]
}"#;

    let out_dir = fresh_temp_dir();
    std::fs::create_dir_all(&out_dir)?;
    let input_path = out_dir.join("cfg.json");
    std::fs::write(&input_path, json.as_bytes())?;

    let exe = std::env::var("CARGO_BIN_EXE_office-parser-cli")
        .unwrap_or_else(|_| env!("CARGO_BIN_EXE_office-parser-cli").to_string());
    let status = std::process::Command::new(exe)
        .arg(&input_path)
        .arg("--out")
        .arg(&out_dir)
        .status()?;
    assert!(status.success());

    let md = std::fs::read_to_string(out_dir.join("cfg.md"))?;
    assert!(md.contains("# server"));
    assert!(md.contains("host: localhost"));
    assert!(md.contains("| # |"));
    assert!(md.contains("| id |"));
    assert!(md.contains("| name |"));

    let _ = std::fs::remove_dir_all(&out_dir);
    Ok(())
}

#[test]
fn cli_epub_emits_markdown_and_assets() -> Result<()> {
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
    <item id="img1" href="img/cover.png" media-type="image/png"/>
  </manifest>
  <spine>
    <itemref idref="ch1"/>
  </spine>
</package>
"#;

    let ch1 = r#"<?xml version="1.0" encoding="UTF-8"?>
<html xmlns="http://www.w3.org/1999/xhtml">
  <head><title>Chapter One</title></head>
  <body>
    <p>Hello.</p>
    <img src="img/cover.png" alt="cover"/>
  </body>
</html>
"#;

    let png = tiny_png_1x1();
    let epub = zip_bytes(&[
        ("META-INF/container.xml", container_xml.as_bytes()),
        ("OEBPS/content.opf", opf.as_bytes()),
        ("OEBPS/ch1.xhtml", ch1.as_bytes()),
        ("OEBPS/img/cover.png", png.as_slice()),
    ])?;

    let out_dir = fresh_temp_dir();
    std::fs::create_dir_all(&out_dir)?;
    let input_path = out_dir.join("book.epub");
    std::fs::write(&input_path, &epub)?;

    let exe = std::env::var("CARGO_BIN_EXE_office-parser-cli")
        .unwrap_or_else(|_| env!("CARGO_BIN_EXE_office-parser-cli").to_string());
    let status = std::process::Command::new(exe)
        .arg(&input_path)
        .arg("--out")
        .arg(&out_dir)
        .status()?;
    assert!(status.success());

    let md = std::fs::read_to_string(out_dir.join("book.md"))?;
    assert!(md.contains("Chapter One"));
    assert!(md.contains("asset/cover.png"));
    assert!(out_dir.join("asset").join("cover.png").exists());

    let _ = std::fs::remove_dir_all(&out_dir);
    Ok(())
}

#[test]
fn cli_xlsx_emits_markdown_tables() -> Result<()> {
    let xlsx = xlsx_minimal_bytes()?;

    let out_dir = fresh_temp_dir();
    std::fs::create_dir_all(&out_dir)?;
    let input_path = out_dir.join("sheet.xlsx");
    std::fs::write(&input_path, &xlsx)?;

    let exe = std::env::var("CARGO_BIN_EXE_office-parser-cli")
        .unwrap_or_else(|_| env!("CARGO_BIN_EXE_office-parser-cli").to_string());
    let status = std::process::Command::new(exe)
        .arg(&input_path)
        .arg("--out")
        .arg(&out_dir)
        .status()?;
    assert!(status.success());

    let md = std::fs::read_to_string(out_dir.join("sheet.md"))?;
    assert!(md.contains("## Sheet: Sheet1"));
    assert!(md.contains("| id | name |"));

    let _ = std::fs::remove_dir_all(&out_dir);
    Ok(())
}

#[test]
fn cli_ods_emits_markdown_tables() -> Result<()> {
    let ods = ods_minimal_bytes()?;

    let out_dir = fresh_temp_dir();
    std::fs::create_dir_all(&out_dir)?;
    let input_path = out_dir.join("sheet.ods");
    std::fs::write(&input_path, &ods)?;

    let exe = std::env::var("CARGO_BIN_EXE_office-parser-cli")
        .unwrap_or_else(|_| env!("CARGO_BIN_EXE_office-parser-cli").to_string());
    let status = std::process::Command::new(exe)
        .arg(&input_path)
        .arg("--out")
        .arg(&out_dir)
        .status()?;
    assert!(status.success());

    let md = std::fs::read_to_string(out_dir.join("sheet.md"))?;
    assert!(md.contains("## Sheet: Sheet1"));
    assert!(md.contains("| id | name |"));

    let _ = std::fs::remove_dir_all(&out_dir);
    Ok(())
}

#[test]
fn cli_xml_wp_like_emits_markdown_and_parses_cdata_html() -> Result<()> {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<rss version="2.0">
  <channel>
    <title>example</title>
    <item>
      <title>Hello</title>
      <content:encoded><![CDATA[<!-- wp:paragraph --><p>Hi <strong>there</strong>.</p><!-- /wp:paragraph -->]]></content:encoded>
    </item>
  </channel>
</rss>
"#;

    let out_dir = fresh_temp_dir();
    std::fs::create_dir_all(&out_dir)?;
    let input_path = out_dir.join("wp.xml");
    std::fs::write(&input_path, xml.as_bytes())?;

    let exe = std::env::var("CARGO_BIN_EXE_office-parser-cli")
        .unwrap_or_else(|_| env!("CARGO_BIN_EXE_office-parser-cli").to_string());
    let status = std::process::Command::new(exe)
        .arg(&input_path)
        .arg("--out")
        .arg(&out_dir)
        .status()?;
    assert!(status.success());

    let md = std::fs::read_to_string(out_dir.join("wp.md"))?;
    assert!(md.contains("# rss"));
    assert!(md.contains("## channel"));
    assert!(md.contains("title: example"));
    // HTML->Markdown best-effort: paragraph text should show up unescaped.
    assert!(md.contains("Hi there."));

    let _ = std::fs::remove_dir_all(&out_dir);
    Ok(())
}
