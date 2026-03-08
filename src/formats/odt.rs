use std::collections::HashMap;

use anyhow::{Context, Result, anyhow};
use quick_xml::Reader;
use quick_xml::events::Event;

use crate::document_ast::{Block, Cell, LinkKind, ListItem, SourceSpan};

use super::{
    ParsedImage, ParsedOfficeDocument, mime_from_filename, read_zip_file, read_zip_file_utf8,
    sha256_hex,
};

fn attr_val(e: &quick_xml::events::BytesStart<'_>, key: &[u8]) -> Option<String> {
    for a in e.attributes().flatten() {
        let k = a.key.as_ref();
        if k == key {
            return Some(String::from_utf8_lossy(&a.value).to_string());
        }
    }
    None
}

fn odf_len_to_px(v: &str) -> Option<f64> {
    let v = v.trim();
    if v.is_empty() {
        return None;
    }
    let mut num_end = 0usize;
    for (i, ch) in v.char_indices() {
        if ch.is_ascii_digit() || ch == '.' || ch == '-' || ch == '+' {
            num_end = i + ch.len_utf8();
            continue;
        }
        break;
    }
    let (num, unit) = v.split_at(num_end);
    let n: f64 = num.parse().ok()?;
    let unit = unit.trim();
    let px = match unit {
        "" => n,
        "px" => n,
        "in" => n * 96.0,
        "cm" => n * (96.0 / 2.54),
        "mm" => n * (96.0 / 25.4),
        "pt" => n * (96.0 / 72.0),
        _ => return None,
    };
    Some(px)
}

fn write_start_tag(out: &mut Vec<u8>, e: &quick_xml::events::BytesStart<'_>) {
    out.push(b'<');
    out.extend_from_slice(e.name().as_ref());
    for a in e.attributes().flatten() {
        out.push(b' ');
        out.extend_from_slice(a.key.as_ref());
        out.extend_from_slice(b"=\"");
        out.extend_from_slice(&a.value);
        out.push(b'"');
    }
    out.push(b'>');
}

fn write_empty_tag(out: &mut Vec<u8>, e: &quick_xml::events::BytesStart<'_>) {
    out.push(b'<');
    out.extend_from_slice(e.name().as_ref());
    for a in e.attributes().flatten() {
        out.push(b' ');
        out.extend_from_slice(a.key.as_ref());
        out.extend_from_slice(b"=\"");
        out.extend_from_slice(&a.value);
        out.push(b'"');
    }
    out.extend_from_slice(b"/>");
}

fn write_end_tag(out: &mut Vec<u8>, name: &[u8]) {
    out.extend_from_slice(b"</");
    out.extend_from_slice(name);
    out.push(b'>');
}

fn parse_list_style_ordered(styles_xml: &str) -> HashMap<String, bool> {
    // list-style name -> ordered? (best-effort)
    let mut out: HashMap<String, bool> = HashMap::new();
    let mut reader = Reader::from_str(styles_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut cur: Option<String> = None;
    let mut cur_ordered = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"text:list-style" {
                    cur = attr_val(&e, b"style:name");
                    cur_ordered = false;
                }
                if cur.is_some() {
                    if e.name().as_ref() == b"text:list-level-style-number" {
                        cur_ordered = true;
                    }
                    if e.name().as_ref() == b"text:list-level-style-bullet" {
                        // keep false
                    }
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"text:list-style" {
                    if let Some(name) = cur.take() {
                        out.insert(name, cur_ordered);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    out
}

fn extract_odt_title(bytes: &[u8]) -> Option<String> {
    let meta_xml = read_zip_file_utf8(bytes, "meta.xml").ok()?;
    let mut reader = Reader::from_str(&meta_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut in_title = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.local_name().as_ref() == b"title" {
                    in_title = true;
                }
            }
            Ok(Event::Text(e)) => {
                if in_title {
                    if let Ok(t) = String::from_utf8(e.to_vec()) {
                        let t = t.trim();
                        if !t.is_empty() {
                            return Some(t.to_string());
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                if e.local_name().as_ref() == b"title" {
                    in_title = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    None
}

fn extract_office_text_children_xml(content_xml: &str) -> Result<Vec<(String, Vec<u8>)>> {
    // Returns (tag_name, xml_bytes) for top-level elements under office:text.
    let mut out = Vec::new();
    let mut reader = Reader::from_str(content_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut in_text = false;
    let mut depth = 0usize;
    let mut capturing: Option<(String, Vec<u8>, usize, Vec<u8>)> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name().as_ref().to_vec();
                if e.name().as_ref() == b"office:text" {
                    in_text = true;
                    depth = 0;
                } else if in_text {
                    if capturing.is_none() {
                        depth += 1;
                    }
                    if capturing.is_none() && depth == 1 {
                        match e.name().as_ref() {
                            b"text:h" | b"text:p" | b"text:list" | b"table:table" => {
                                let mut v = Vec::with_capacity(4096);
                                write_start_tag(&mut v, &e);
                                capturing = Some((
                                    String::from_utf8_lossy(e.name().as_ref()).to_string(),
                                    v,
                                    1,
                                    name,
                                ));
                                buf.clear();
                                continue;
                            }
                            _ => {}
                        }
                    }

                    if let Some((_t, v, d, _n)) = capturing.as_mut() {
                        *d += 1;
                        write_start_tag(v, &e);
                    }
                } else if let Some((_t, v, d, _n)) = capturing.as_mut() {
                    *d += 1;
                    write_start_tag(v, &e);
                }
            }
            Ok(Event::Empty(e)) => {
                if let Some((_t, v, _d, _n)) = capturing.as_mut() {
                    write_empty_tag(v, &e);
                }
            }
            Ok(Event::Text(e)) => {
                if let Some((_t, v, _d, _n)) = capturing.as_mut() {
                    v.extend_from_slice(e.as_ref());
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"office:text" {
                    in_text = false;
                }

                if let Some((t, mut v, d, n)) = capturing.take() {
                    write_end_tag(&mut v, e.name().as_ref());
                    let d = d.saturating_sub(1);
                    if d == 0 && e.name().as_ref() == n.as_slice() {
                        out.push((t, v));
                        if in_text {
                            depth = depth.saturating_sub(1);
                        }
                    } else {
                        capturing = Some((t, v, d, n));
                    }
                } else if in_text {
                    depth = depth.saturating_sub(1);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("odt content.xml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn extract_text_and_links(xml: &[u8]) -> Result<(String, Vec<(String, Option<String>)>)> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut out = String::new();
    let mut links: Vec<(String, Option<String>)> = Vec::new();
    let mut in_a = false;
    let mut cur_href: Option<String> = None;
    let mut cur_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:a" => {
                    in_a = true;
                    cur_href = attr_val(&e, b"xlink:href");
                    cur_text.clear();
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"text:tab" => {
                    out.push('\t');
                    if in_a {
                        cur_text.push('\t');
                    }
                }
                b"text:line-break" => {
                    out.push('\n');
                    if in_a {
                        cur_text.push('\n');
                    }
                }
                b"text:s" => {
                    let count = attr_val(&e, b"text:c")
                        .and_then(|s| s.parse::<usize>().ok())
                        .unwrap_or(1);
                    for _ in 0..count {
                        out.push(' ');
                        if in_a {
                            cur_text.push(' ');
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                let t = e.decode().map_err(|ee| anyhow!("odt text decode: {ee}"))?;
                out.push_str(&t);
                if in_a {
                    cur_text.push_str(&t);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"text:a" {
                    in_a = false;
                    if let Some(href) = cur_href.take() {
                        let display = cur_text.trim().to_string();
                        links.push((
                            href,
                            if display.is_empty() {
                                None
                            } else {
                                Some(display)
                            },
                        ));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("odt text parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok((out.trim().to_string(), links))
}

fn parse_heading(xml: &[u8]) -> Result<(u8, String, Vec<(String, Option<String>)>)> {
    let mut level: u8 = 1;
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut inner: Vec<u8> = Vec::new();
    let mut depth = 0usize;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if depth == 0 && e.name().as_ref() == b"text:h" {
                    if let Some(v) = attr_val(&e, b"text:outline-level") {
                        level = v.parse::<u8>().unwrap_or(1).max(1);
                    }
                }
                depth += 1;
                write_start_tag(&mut inner, &e);
            }
            Ok(Event::Empty(e)) => write_empty_tag(&mut inner, &e),
            Ok(Event::Text(e)) => inner.extend_from_slice(e.as_ref()),
            Ok(Event::End(e)) => {
                write_end_tag(&mut inner, e.name().as_ref());
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("odt heading parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    let (text, links) = extract_text_and_links(&inner)?;
    Ok((level, text, links))
}

fn parse_paragraph(xml: &[u8]) -> Result<(String, Vec<(String, Option<String>)>)> {
    extract_text_and_links(xml)
}

fn parse_list(xml: &[u8], ordered: bool) -> Result<(Vec<ListItem>, Vec<(String, Option<String>)>)> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut items: Vec<ListItem> = Vec::new();
    let mut links: Vec<(String, Option<String>)> = Vec::new();
    let mut list_depth: u8 = 0;

    // Capture each text:p inside text:list-item at current level.
    let mut capturing_p: Option<(Vec<u8>, usize, u8)> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:list" => {
                    list_depth = list_depth.saturating_add(1);
                }
                b"text:list-item" => {}
                b"text:p" => {
                    if capturing_p.is_none() {
                        let mut v = Vec::with_capacity(1024);
                        write_start_tag(&mut v, &e);
                        let level = list_depth.saturating_sub(1);
                        capturing_p = Some((v, 1, level));
                    } else if let Some((v, d, _lvl)) = capturing_p.as_mut() {
                        *d = d.saturating_add(1);
                        write_start_tag(v, &e);
                    }
                }
                _ => {
                    if let Some((v, d, _lvl)) = capturing_p.as_mut() {
                        *d = d.saturating_add(1);
                        write_start_tag(v, &e);
                    }
                }
            },
            Ok(Event::Empty(e)) => {
                if let Some((v, _d, _lvl)) = capturing_p.as_mut() {
                    write_empty_tag(v, &e);
                }
            }
            Ok(Event::Text(e)) => {
                if let Some((v, _d, _lvl)) = capturing_p.as_mut() {
                    v.extend_from_slice(e.as_ref());
                }
            }
            Ok(Event::End(e)) => {
                if let Some((mut v, d, lvl)) = capturing_p.take() {
                    write_end_tag(&mut v, e.name().as_ref());
                    let d = d.saturating_sub(1);
                    if d == 0 && e.name().as_ref() == b"text:p" {
                        let (t, l) = extract_text_and_links(&v)?;
                        links.extend(l);
                        if !t.trim().is_empty() {
                            items.push(ListItem {
                                level: lvl,
                                text: t,
                                source: SourceSpan::default(),
                            });
                        }
                    } else {
                        capturing_p = Some((v, d, lvl));
                    }
                }

                if e.name().as_ref() == b"text:list" {
                    list_depth = list_depth.saturating_sub(1);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("odt list parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    let _ = ordered; // marker chosen during rendering, list items are semantic.
    Ok((items, links))
}

fn parse_table(xml: &[u8]) -> Result<(Vec<Vec<Cell>>, Vec<(String, Option<String>)>)> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut rows: Vec<Vec<Cell>> = Vec::new();
    let mut cur_row: Vec<Cell> = Vec::new();

    let mut capturing_cell: Option<(Vec<u8>, usize, usize, usize)> = None;
    // (xml_bytes, depth, colspan, rowspan)
    let mut links: Vec<(String, Option<String>)> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"table:table-row" {
                    cur_row = Vec::new();
                }
                if e.name().as_ref() == b"table:table-cell" {
                    if capturing_cell.is_none() {
                        let colspan = attr_val(&e, b"table:number-columns-spanned")
                            .and_then(|s| s.parse::<usize>().ok())
                            .unwrap_or(1)
                            .max(1);
                        let rowspan = attr_val(&e, b"table:number-rows-spanned")
                            .and_then(|s| s.parse::<usize>().ok())
                            .unwrap_or(1)
                            .max(1);
                        let mut v = Vec::with_capacity(1024);
                        write_start_tag(&mut v, &e);
                        capturing_cell = Some((v, 1, colspan, rowspan));
                    }
                } else if let Some((v, d, _cs, _rs)) = capturing_cell.as_mut() {
                    *d += 1;
                    write_start_tag(v, &e);
                }
            }
            Ok(Event::Empty(e)) => {
                if let Some((v, _d, _cs, _rs)) = capturing_cell.as_mut() {
                    write_empty_tag(v, &e);
                }
            }
            Ok(Event::Text(e)) => {
                if let Some((v, _d, _cs, _rs)) = capturing_cell.as_mut() {
                    v.extend_from_slice(e.as_ref());
                }
            }
            Ok(Event::End(e)) => {
                if let Some((mut v, d, cs, rs)) = capturing_cell.take() {
                    write_end_tag(&mut v, e.name().as_ref());
                    let d = d.saturating_sub(1);
                    if d == 0 && e.name().as_ref() == b"table:table-cell" {
                        let (t, l) = extract_text_and_links(&v)?;
                        links.extend(l);
                        cur_row.push(Cell {
                            text: t,
                            colspan: cs,
                            rowspan: rs,
                        });
                    } else {
                        capturing_cell = Some((v, d, cs, rs));
                    }
                }

                if e.name().as_ref() == b"table:table-row" {
                    rows.push(cur_row.clone());
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("odt table parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok((rows, links))
}

fn unit_from_table_rows(rows: &[Vec<Cell>]) -> Option<String> {
    for r in rows {
        for c in r {
            let t = c.text.trim();
            if t.contains('%') {
                return Some("%".to_string());
            }
            for sym in ["$", "€", "£", "¥", "₩", "₹"] {
                if t.contains(sym) {
                    return Some(sym.to_string());
                }
            }
        }
    }
    None
}

fn extract_draw_object_hrefs(xml: &[u8]) -> Vec<String> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"draw:object" || e.name().as_ref() == b"draw:object-ole" {
                    if let Some(href) = attr_val(&e, b"xlink:href") {
                        let mut p = href.trim().to_string();
                        while p.starts_with("./") {
                            p = p.trim_start_matches("./").to_string();
                        }
                        if !p.is_empty() {
                            out.push(p);
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn extract_first_table_xml(xml: &str) -> Result<Option<Vec<u8>>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut capturing: Option<(Vec<u8>, usize, Vec<u8>)> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if let Some((v, d, _n)) = capturing.as_mut() {
                    *d += 1;
                    write_start_tag(v, &e);
                } else if e.name().as_ref() == b"table:table" {
                    let mut v = Vec::with_capacity(4096);
                    write_start_tag(&mut v, &e);
                    capturing = Some((v, 1, e.name().as_ref().to_vec()));
                }
            }
            Ok(Event::Empty(e)) => {
                if let Some((v, _d, _n)) = capturing.as_mut() {
                    write_empty_tag(v, &e);
                }
            }
            Ok(Event::Text(e)) => {
                if let Some((v, _d, _n)) = capturing.as_mut() {
                    v.extend_from_slice(e.as_ref());
                }
            }
            Ok(Event::End(e)) => {
                if let Some((mut v, d, n)) = capturing.take() {
                    write_end_tag(&mut v, e.name().as_ref());
                    let d = d.saturating_sub(1);
                    if d == 0 && e.name().as_ref() == n.as_slice() {
                        return Ok(Some(v));
                    }
                    capturing = Some((v, d, n));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("odt chart object xml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(None)
}

fn extract_chart_type_and_title(xml: &str) -> (Option<String>, Option<String>) {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut chart_type: Option<String> = None;
    let mut in_chart_title = false;
    let mut title = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"chart:chart" {
                    if chart_type.is_none() {
                        chart_type =
                            attr_val(&e, b"chart:class").or_else(|| attr_val(&e, b"class"));
                        chart_type = chart_type
                            .map(|v| v.rsplit(':').next().unwrap_or(v.as_str()).to_string());
                    }
                }
                if e.name().as_ref() == b"chart:title" {
                    in_chart_title = true;
                    title.clear();
                }
            }
            Ok(Event::Text(e)) => {
                if in_chart_title {
                    title.push_str(&String::from_utf8_lossy(e.as_ref()));
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"chart:title" {
                    in_chart_title = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    let t = title.trim();
    let title = if t.is_empty() {
        None
    } else {
        Some(t.to_string())
    };
    (chart_type, title)
}

fn mermaid_shape_from_odf_type(t: Option<&str>) -> &'static str {
    let t = t.unwrap_or("").trim();
    match t {
        "" => "rect",
        "rectangle" | "rect" => "rect",
        "round-rectangle" | "roundrect" => "rounded",
        "ellipse" | "circle" => "circle",
        "diamond" => "diamond",
        "hexagon" => "hex",
        "parallelogram" => "lean-r",
        "trapezoid" => "trap-b",
        "triangle" | "isosceles-triangle" | "right-triangle" => "tri",
        "cylinder" => "cyl",
        _ => "rect",
    }
}

fn diagram_from_xml(xml: &[u8]) -> Result<Option<(String, serde_json::Value)>> {
    #[derive(Clone, Debug)]
    struct Node {
        id: u32,
        label: String,
        mermaid_shape: &'static str,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
    }
    #[derive(Clone, Copy, Debug)]
    enum ArrowDir {
        None,
        Forward,
        Reverse,
        Both,
    }
    #[derive(Clone, Debug)]
    struct Edge {
        from: u32,
        to: u32,
        dir: ArrowDir,
        label: Option<String>,
    }

    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut nodes: Vec<Node> = Vec::new();
    let mut node_by_id: HashMap<String, u32> = Default::default();
    let mut next_id: u32 = 1;

    let mut pending_frame: Vec<(
        Option<String>,
        Option<String>,
        Option<(f64, f64, f64, f64)>,
        String,
    )> = Vec::new();
    // (draw:id, shape_type, bbox, text)

    let mut in_connector = false;
    let mut cur_conn_text = String::new();
    let mut conns: Vec<(
        Option<String>,
        Option<String>,
        ArrowDir,
        Option<(f64, f64, f64, f64)>,
        Option<String>,
    )> = Vec::new();
    // (start_id, end_id, dir, coords(x1,y1,x2,y2), label)

    let marker_present = |v: Option<String>| -> bool {
        let Some(v) = v else { return false };
        let t = v.trim();
        !(t.is_empty() || t.eq_ignore_ascii_case("none"))
    };

    let arrow_dir_from_elem = |e: &quick_xml::events::BytesStart<'_>| -> ArrowDir {
        let has_start = marker_present(attr_val(e, b"draw:marker-start"))
            || marker_present(attr_val(e, b"svg:marker-start"));
        let has_end = marker_present(attr_val(e, b"draw:marker-end"))
            || marker_present(attr_val(e, b"svg:marker-end"));
        match (has_start, has_end) {
            (false, false) => ArrowDir::None,
            (false, true) => ArrowDir::Forward,
            (true, false) => ArrowDir::Reverse,
            (true, true) => ArrowDir::Both,
        }
    };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"draw:frame" | b"draw:custom-shape" => {
                    let id = attr_val(&e, b"draw:id").or_else(|| attr_val(&e, b"xml:id"));
                    let x = attr_val(&e, b"svg:x").and_then(|v| odf_len_to_px(&v));
                    let y = attr_val(&e, b"svg:y").and_then(|v| odf_len_to_px(&v));
                    let w = attr_val(&e, b"svg:width").and_then(|v| odf_len_to_px(&v));
                    let h = attr_val(&e, b"svg:height").and_then(|v| odf_len_to_px(&v));
                    let bbox = match (x, y, w, h) {
                        (Some(x), Some(y), Some(w), Some(h)) => Some((x, y, w, h)),
                        _ => None,
                    };
                    pending_frame.push((id, None, bbox, String::new()));
                }
                b"draw:connector" | b"draw:line" => {
                    in_connector = true;
                    let start_id = attr_val(&e, b"draw:start-shape");
                    let end_id = attr_val(&e, b"draw:end-shape");
                    let x1 = attr_val(&e, b"svg:x1").and_then(|v| odf_len_to_px(&v));
                    let y1 = attr_val(&e, b"svg:y1").and_then(|v| odf_len_to_px(&v));
                    let x2 = attr_val(&e, b"svg:x2").and_then(|v| odf_len_to_px(&v));
                    let y2 = attr_val(&e, b"svg:y2").and_then(|v| odf_len_to_px(&v));
                    let coords = match (x1, y1, x2, y2) {
                        (Some(x1), Some(y1), Some(x2), Some(y2)) => Some((x1, y1, x2, y2)),
                        _ => None,
                    };
                    let dir = arrow_dir_from_elem(&e);
                    cur_conn_text.clear();
                    conns.push((start_id, end_id, dir, coords, None));
                }
                b"draw:enhanced-geometry" => {
                    if let Some((_, shape_type, _, _)) = pending_frame.last_mut() {
                        if shape_type.is_none() {
                            *shape_type = attr_val(&e, b"draw:type");
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"draw:connector" | b"draw:line" => {
                    let start_id = attr_val(&e, b"draw:start-shape");
                    let end_id = attr_val(&e, b"draw:end-shape");
                    let x1 = attr_val(&e, b"svg:x1").and_then(|v| odf_len_to_px(&v));
                    let y1 = attr_val(&e, b"svg:y1").and_then(|v| odf_len_to_px(&v));
                    let x2 = attr_val(&e, b"svg:x2").and_then(|v| odf_len_to_px(&v));
                    let y2 = attr_val(&e, b"svg:y2").and_then(|v| odf_len_to_px(&v));
                    let coords = match (x1, y1, x2, y2) {
                        (Some(x1), Some(y1), Some(x2), Some(y2)) => Some((x1, y1, x2, y2)),
                        _ => None,
                    };
                    let dir = arrow_dir_from_elem(&e);
                    conns.push((start_id, end_id, dir, coords, None));
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                let trimmed = String::from_utf8_lossy(e.as_ref()).trim().to_string();
                if trimmed.is_empty() {
                    // ignore
                } else {
                    if let Some((_, _, _, t)) = pending_frame.last_mut() {
                        if !t.is_empty() {
                            t.push(' ');
                        }
                        t.push_str(&trimmed);
                    }
                    if in_connector {
                        if !cur_conn_text.is_empty() {
                            cur_conn_text.push(' ');
                        }
                        cur_conn_text.push_str(&trimmed);
                    }
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"draw:frame" || e.name().as_ref() == b"draw:custom-shape" {
                    if let Some((id, shape_type, bbox, text)) = pending_frame.pop() {
                        let t = text.trim();
                        if let (Some(id_s), Some((x, y, w, h))) = (id.as_deref(), bbox) {
                            if !t.is_empty() {
                                let node_id =
                                    *node_by_id.entry(id_s.to_string()).or_insert_with(|| {
                                        let id = next_id;
                                        next_id = next_id.saturating_add(1);
                                        id
                                    });
                                nodes.push(Node {
                                    id: node_id,
                                    label: t.to_string(),
                                    mermaid_shape: mermaid_shape_from_odf_type(
                                        shape_type.as_deref(),
                                    ),
                                    x,
                                    y,
                                    w,
                                    h,
                                });
                            }
                        }
                    }
                }
                if e.name().as_ref() == b"draw:connector" || e.name().as_ref() == b"draw:line" {
                    in_connector = false;
                    if let Some(last) = conns.last_mut() {
                        let t = cur_conn_text.trim();
                        if !t.is_empty() {
                            last.4 = Some(t.to_string());
                        }
                    }
                    cur_conn_text.clear();
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("odt diagram xml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    if nodes.is_empty() || conns.is_empty() {
        return Ok(None);
    }

    // Build edges
    let nearest = |x: f64, y: f64| -> Option<u32> {
        let mut best: Option<(u32, f64)> = None;
        for n in &nodes {
            let cx = n.x + n.w / 2.0;
            let cy = n.y + n.h / 2.0;
            let d = (cx - x).hypot(cy - y);
            best = match best {
                None => Some((n.id, d)),
                Some((_bid, bd)) if d < bd => Some((n.id, d)),
                Some(v) => Some(v),
            };
        }
        let (id, d) = best?;
        if d <= 120.0 { Some(id) } else { None }
    };

    let mut edges: Vec<Edge> = Vec::new();
    for (s, t, dir, coords, label) in conns {
        let mut from = s.as_deref().and_then(|id| node_by_id.get(id).copied());
        let mut to = t.as_deref().and_then(|id| node_by_id.get(id).copied());
        if from.is_none() || to.is_none() {
            if let Some((x1, y1, x2, y2)) = coords {
                if from.is_none() {
                    from = nearest(x1, y1);
                }
                if to.is_none() {
                    to = nearest(x2, y2);
                }
            }
        }
        let (Some(from), Some(to)) = (from, to) else {
            continue;
        };
        if from == to {
            continue;
        }
        edges.push(Edge {
            from,
            to,
            dir,
            label,
        });
    }

    if edges.is_empty() {
        return Ok(None);
    }

    let esc = |s: &str| -> String {
        let mut out = s.replace('"', "'");
        out = out.replace('|', "/");
        out = out.replace('\r', "");
        out = out.replace('\n', "<br/>");
        out.trim().to_string()
    };

    let mut mermaid = String::new();
    mermaid.push_str("```mermaid\nflowchart LR\n");
    let mut ids: std::collections::BTreeSet<u32> = Default::default();
    for e in &edges {
        ids.insert(e.from);
        ids.insert(e.to);
    }
    let node_map: HashMap<u32, &Node> = nodes.iter().map(|n| (n.id, n)).collect();
    for id in ids.iter().copied() {
        if let Some(n) = node_map.get(&id).copied() {
            mermaid.push_str(&format!(
                "  n{id}@{{ shape: {}, label: \"{}\" }}\n",
                n.mermaid_shape,
                esc(&n.label)
            ));
        }
    }
    for e in &edges {
        let op = match e.dir {
            ArrowDir::None => "---",
            ArrowDir::Forward => "-->",
            ArrowDir::Reverse => "<--",
            ArrowDir::Both => "<-->",
        };
        if let Some(lbl) = e.label.as_deref().map(str::trim).filter(|s| !s.is_empty()) {
            mermaid.push_str(&format!("  n{} {op}|{}| n{}\n", e.from, esc(lbl), e.to));
        } else {
            mermaid.push_str(&format!("  n{} {op} n{}\n", e.from, e.to));
        }
    }
    mermaid.push_str("```\n");

    let graph_json = serde_json::json!({
        "nodes": nodes.iter().map(|n| serde_json::json!({
            "id": format!("n{}", n.id),
            "text": n.label.clone(),
            "bbox": {"x": n.x, "y": n.y, "w": n.w, "h": n.h},
        })).collect::<Vec<_>>(),
        "edges": edges.iter().map(|e| serde_json::json!({
            "from": format!("n{}", e.from),
            "to": format!("n{}", e.to),
            "kind": "connector",
            "label": e.label,
        })).collect::<Vec<_>>(),
        "warnings": []
    });

    Ok(Some((mermaid, graph_json)))
}

fn extract_odt_images(
    bytes: &[u8],
    content_xml: &str,
) -> Result<Vec<(String, Option<String>, Vec<u8>, String)>> {
    // Returns (href, alt, bytes, filename)
    let mut reader = Reader::from_str(content_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    let mut pending_alt: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"draw:frame" {
                    pending_alt = attr_val(&e, b"draw:name");
                }
                if e.name().as_ref() == b"draw:image" {
                    if let Some(href) = attr_val(&e, b"xlink:href") {
                        let mut p = href.clone();
                        while p.starts_with("./") {
                            p = p.trim_start_matches("./").to_string();
                        }
                        let entry = p;
                        if let Ok(img_bytes) = read_zip_file(bytes, &entry) {
                            let filename = entry.split('/').last().unwrap_or("").to_string();
                            let _mime = mime_from_filename(&filename);
                            out.push((href, pending_alt.take(), img_bytes, filename));
                            pending_alt = None;
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

pub fn parse_odt(bytes: &[u8]) -> Result<Vec<Block>> {
    Ok(parse_odt_full(bytes)?.blocks)
}

pub fn parse_odt_full(bytes: &[u8]) -> Result<ParsedOfficeDocument> {
    let mimetype = read_zip_file_utf8(bytes, "mimetype").unwrap_or_else(|_| "".to_string());
    if !mimetype.contains("opendocument.text") {
        return Err(anyhow!("not an ODT (mimetype={})", mimetype.trim()));
    }

    let content_xml = read_zip_file_utf8(bytes, "content.xml").context("read ODT content.xml")?;
    let styles_xml = read_zip_file_utf8(bytes, "styles.xml").unwrap_or_else(|_| "".to_string());
    let list_styles = parse_list_style_ordered(&styles_xml);

    let children = extract_office_text_children_xml(&content_xml)?;

    // Extract and dedupe images referenced anywhere.
    let images_raw = extract_odt_images(bytes, &content_xml).unwrap_or_default();
    let mut images: Vec<ParsedImage> = Vec::new();
    let mut image_by_hash: HashMap<String, usize> = HashMap::new();

    for (_href, _alt, img_bytes, filename) in &images_raw {
        let hash = sha256_hex(img_bytes);
        image_by_hash.entry(hash.clone()).or_insert_with(|| {
            images.push(ParsedImage {
                id: format!("sha256:{hash}"),
                bytes: img_bytes.clone(),
                mime_type: mime_from_filename(filename),
                filename: if filename.is_empty() {
                    None
                } else {
                    Some(filename.clone())
                },
            });
            images.len() - 1
        });
    }

    let mut blocks: Vec<Block> = Vec::new();
    let mut next_block_index: usize = 0;

    let mut charts_meta: Vec<serde_json::Value> = Vec::new();
    let mut diagram_graphs_meta: Vec<serde_json::Value> = Vec::new();

    for (tag, xml) in children {
        // Embedded objects (charts, etc.) referenced from within this element.
        for href in extract_draw_object_hrefs(&xml) {
            let entry = format!("{href}/content.xml");
            if let Ok(obj_xml) = read_zip_file_utf8(bytes, &entry) {
                let (chart_type, title) = extract_chart_type_and_title(&obj_xml);
                if let Ok(Some(table_xml)) = extract_first_table_xml(&obj_xml) {
                    let (rows, links) = parse_table(&table_xml)?;
                    if !rows.is_empty() {
                        let chart_type = chart_type.unwrap_or_else(|| "chart".to_string());
                        let chart_shape = format!("{chart_type} chart");
                        let mut note = format!("Note: {chart_shape}");
                        let unit = unit_from_table_rows(&rows);
                        if let Some(u) = unit.as_deref() {
                            note.push_str(&format!("; units: {u}"));
                        }

                        charts_meta.push(serde_json::json!({
                            "chart_type": chart_type,
                            "title": title,
                            "unit": unit,
                        }));

                        blocks.push(Block::Heading {
                            block_index: next_block_index,
                            level: 3,
                            text: title.clone().unwrap_or_else(|| "Chart".to_string()),
                            source: SourceSpan::default(),
                        });
                        next_block_index += 1;
                        blocks.push(Block::Paragraph {
                            block_index: next_block_index,
                            text: note,
                            source: SourceSpan::default(),
                        });
                        next_block_index += 1;
                        blocks.push(Block::Table {
                            block_index: next_block_index,
                            rows,
                            source: SourceSpan::default(),
                        });
                        next_block_index += 1;
                        for (href, t) in links {
                            blocks.push(Block::Link {
                                block_index: next_block_index,
                                url: href,
                                text: t,
                                kind: LinkKind::Unknown,
                                source: SourceSpan::default(),
                            });
                            next_block_index += 1;
                        }
                    }
                }
            }
        }

        match tag.as_str() {
            "text:h" => {
                let (level, text, links) = parse_heading(&xml)?;
                if !text.trim().is_empty() {
                    blocks.push(Block::Heading {
                        block_index: next_block_index,
                        level,
                        text,
                        source: SourceSpan::default(),
                    });
                    next_block_index += 1;
                }
                for (href, t) in links {
                    blocks.push(Block::Link {
                        block_index: next_block_index,
                        url: href,
                        text: t,
                        kind: LinkKind::Unknown,
                        source: SourceSpan::default(),
                    });
                    next_block_index += 1;
                }
            }
            "text:p" => {
                let (text, links) = parse_paragraph(&xml)?;
                if !text.trim().is_empty() {
                    blocks.push(Block::Paragraph {
                        block_index: next_block_index,
                        text,
                        source: SourceSpan::default(),
                    });
                    next_block_index += 1;
                }
                for (href, t) in links {
                    blocks.push(Block::Link {
                        block_index: next_block_index,
                        url: href,
                        text: t,
                        kind: LinkKind::Unknown,
                        source: SourceSpan::default(),
                    });
                    next_block_index += 1;
                }
            }
            "text:list" => {
                // Determine ordered via style-name.
                let style_name = {
                    let mut reader = Reader::from_reader(xml.as_slice());
                    reader.config_mut().trim_text(true);
                    let mut b = Vec::new();
                    let mut name: Option<String> = None;
                    if let Ok(Event::Start(e)) = reader.read_event_into(&mut b) {
                        if e.name().as_ref() == b"text:list" {
                            name = attr_val(&e, b"text:style-name");
                        }
                    }
                    name
                };
                let ordered = style_name
                    .as_deref()
                    .and_then(|n| list_styles.get(n).copied())
                    .unwrap_or(false);

                let (items, links) = parse_list(&xml, ordered)?;
                if !items.is_empty() {
                    blocks.push(Block::List {
                        block_index: next_block_index,
                        ordered,
                        items,
                        source: SourceSpan::default(),
                    });
                    next_block_index += 1;
                }
                for (href, t) in links {
                    blocks.push(Block::Link {
                        block_index: next_block_index,
                        url: href,
                        text: t,
                        kind: LinkKind::Unknown,
                        source: SourceSpan::default(),
                    });
                    next_block_index += 1;
                }
            }
            "table:table" => {
                let (rows, links) = parse_table(&xml)?;
                blocks.push(Block::Table {
                    block_index: next_block_index,
                    rows,
                    source: SourceSpan::default(),
                });
                next_block_index += 1;
                for (href, t) in links {
                    blocks.push(Block::Link {
                        block_index: next_block_index,
                        url: href,
                        text: t,
                        kind: LinkKind::Unknown,
                        source: SourceSpan::default(),
                    });
                    next_block_index += 1;
                }
            }
            _ => {}
        }

        if let Some((mermaid, graph_json)) = diagram_from_xml(&xml)? {
            blocks.push(Block::Heading {
                block_index: next_block_index,
                level: 3,
                text: "Diagram".to_string(),
                source: SourceSpan::default(),
            });
            next_block_index += 1;
            blocks.push(Block::Paragraph {
                block_index: next_block_index,
                text: mermaid,
                source: SourceSpan::default(),
            });
            next_block_index += 1;
            diagram_graphs_meta.push(graph_json);
        }
    }

    // Emit images as separate blocks (stable order by first appearance in content.xml scan).
    for (href, alt, img_bytes, filename) in images_raw {
        let hash = sha256_hex(&img_bytes);
        let ct = mime_from_filename(&filename);
        blocks.push(Block::Image {
            block_index: next_block_index,
            id: format!("sha256:{hash}"),
            filename: if filename.is_empty() {
                None
            } else {
                Some(filename)
            },
            content_type: Some(ct),
            alt,
            source: SourceSpan::default(),
        });
        next_block_index += 1;
        let _ = href;
    }

    // Renumber indices.
    for (idx, b) in blocks.iter_mut().enumerate() {
        match b {
            Block::Heading { block_index, .. }
            | Block::Paragraph { block_index, .. }
            | Block::List { block_index, .. }
            | Block::Table { block_index, .. }
            | Block::Image { block_index, .. }
            | Block::Link { block_index, .. } => *block_index = idx,
        }
    }

    Ok(ParsedOfficeDocument {
        blocks,
        images,
        metadata_json: serde_json::json!({
            "kind": "odt",
            "title": extract_odt_title(bytes),
            "charts": charts_meta,
            "diagram_graphs": diagram_graphs_meta,
        }),
    })
}

pub fn parse(bytes: &[u8]) -> crate::Result<crate::Document> {
    let parsed =
        parse_odt_full(bytes).map_err(|e| crate::Error::from_parse(crate::Format::Odt, e))?;
    Ok(super::finalize(crate::Format::Odt, parsed))
}
