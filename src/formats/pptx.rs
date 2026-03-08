use std::collections::HashMap;

use anyhow::{Context, Result, anyhow};
use quick_xml::Reader;
use quick_xml::events::Event;

use crate::document_ast::{Block, Cell, ListItem, SourceSpan};

use super::{
    ParsedImage, ParsedOfficeDocument, mime_from_filename, read_zip_file, read_zip_file_utf8,
    sha256_hex,
};

const RELTYPE_NOTES_SLIDE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";

fn attr_val_local(e: &quick_xml::events::BytesStart<'_>, key_local: &[u8]) -> Option<String> {
    for a in e.attributes().flatten() {
        let k = a.key.as_ref();
        let local = k.split(|b| *b == b':').last().unwrap_or(k);
        if local == key_local {
            return Some(String::from_utf8_lossy(&a.value).to_string());
        }
    }
    None
}

fn parse_relationships(rels_xml: &str) -> Result<HashMap<String, (String, String)>> {
    fn scan_attr(tag: &[u8], key: &[u8]) -> Option<String> {
        let mut i = 0usize;
        while i + key.len() + 3 < tag.len() {
            if &tag[i..i + key.len()] == key {
                let mut j = i + key.len();
                while j < tag.len() && (tag[j] == b' ' || tag[j] == b'\n' || tag[j] == b'\t') {
                    j += 1;
                }
                if j >= tag.len() || tag[j] != b'=' {
                    i += 1;
                    continue;
                }
                j += 1;
                while j < tag.len() && (tag[j] == b' ' || tag[j] == b'\n' || tag[j] == b'\t') {
                    j += 1;
                }
                if j >= tag.len() {
                    return None;
                }
                let quote = tag[j];
                if quote != b'"' && quote != b'\'' {
                    i += 1;
                    continue;
                }
                j += 1;
                let start = j;
                while j < tag.len() && tag[j] != quote {
                    j += 1;
                }
                return Some(String::from_utf8_lossy(&tag[start..j]).to_string());
            }
            i += 1;
        }
        None
    }

    // Id -> (type, target)
    let mut out = HashMap::new();
    let bytes = rels_xml.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() {
        let Some(p) = bytes[i..]
            .windows(b"<Relationship".len())
            .position(|w| w == b"<Relationship")
        else {
            break;
        };
        let start = i + p;
        let Some(end_rel) = bytes[start..].iter().position(|b| *b == b'>') else {
            break;
        };
        let tag = &bytes[start..start + end_rel + 1];

        let id = scan_attr(tag, b"Id")
            .or_else(|| scan_attr(tag, b"id"))
            .unwrap_or_default();
        if !id.is_empty() {
            let ty = scan_attr(tag, b"Type")
                .or_else(|| scan_attr(tag, b"type"))
                .unwrap_or_default();
            let target = scan_attr(tag, b"Target")
                .or_else(|| scan_attr(tag, b"target"))
                .unwrap_or_default();
            out.insert(id, (ty, target));
        }

        i = start + end_rel + 1;
    }

    Ok(out)
}

fn dirname(path: &str) -> &str {
    match path.rfind('/') {
        Some(i) => &path[..=i],
        None => "",
    }
}

fn basename(path: &str) -> &str {
    match path.rfind('/') {
        Some(i) => &path[i + 1..],
        None => path,
    }
}

fn resolve_target(base_part: &str, target: &str) -> String {
    if target.starts_with('/') {
        return target.trim_start_matches('/').to_string();
    }
    let base_dir = dirname(base_part);
    let joined = format!("{}{}", base_dir, target);

    // Normalize ./ and ../
    let mut stack: Vec<&str> = Vec::new();
    for comp in joined.split('/') {
        if comp.is_empty() || comp == "." {
            continue;
        }
        if comp == ".." {
            let _ = stack.pop();
            continue;
        }
        stack.push(comp);
    }
    stack.join("/")
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

fn capture_sp_tree_children(slide_xml: &str) -> Result<Vec<(String, Vec<u8>)>> {
    // Return (tag_name, xml_bytes) for direct children of p:spTree that we care about.
    let mut out: Vec<(String, Vec<u8>)> = Vec::new();
    let mut reader = Reader::from_str(slide_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut in_tree = false;
    let mut depth = 0usize;
    let mut capturing: Option<(String, Vec<u8>, usize, Vec<u8>)> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.local_name().as_ref() == b"spTree" {
                    in_tree = true;
                    depth = 0;
                } else if in_tree {
                    if capturing.is_none() {
                        depth += 1;
                    }
                    if capturing.is_none() && depth == 1 {
                        match e.local_name().as_ref() {
                            b"sp" | b"pic" | b"graphicFrame" | b"cxnSp" | b"grpSp" => {
                                let mut v = Vec::with_capacity(4096);
                                write_start_tag(&mut v, &e);
                                capturing = Some((
                                    String::from_utf8_lossy(e.name().as_ref()).to_string(),
                                    v,
                                    1,
                                    e.name().as_ref().to_vec(),
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
                if e.local_name().as_ref() == b"spTree" {
                    // unlikely
                }
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
                if e.local_name().as_ref() == b"spTree" {
                    in_tree = false;
                }

                if let Some((t, mut v, d, n)) = capturing.take() {
                    write_end_tag(&mut v, e.name().as_ref());
                    let d = d.saturating_sub(1);
                    if d == 0 && e.name().as_ref() == n.as_slice() {
                        out.push((t, v));
                        if in_tree {
                            depth = depth.saturating_sub(1);
                        }
                    } else {
                        capturing = Some((t, v, d, n));
                    }
                } else if in_tree {
                    depth = depth.saturating_sub(1);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("slide xml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn extract_placeholder_type(xml: &[u8]) -> Option<String> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                if e.local_name().as_ref() == b"ph" {
                    if let Some(t) = attr_val_local(&e, b"type") {
                        return Some(t);
                    }
                    return Some("body".to_string());
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

fn extract_cnvpr_alt(xml: &[u8]) -> Option<String> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                if e.local_name().as_ref() == b"cNvPr" {
                    let descr = attr_val_local(&e, b"descr").filter(|s| !s.trim().is_empty());
                    let title = attr_val_local(&e, b"title").filter(|s| !s.trim().is_empty());
                    return descr.or(title);
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

fn extract_first_blip_embed(xml: &[u8]) -> Option<String> {
    fn scan_attr(xml: &[u8], key: &[u8]) -> Option<String> {
        let mut i = 0usize;
        while i + key.len() + 3 < xml.len() {
            if &xml[i..i + key.len()] == key {
                let mut j = i + key.len();
                while j < xml.len() && (xml[j] == b' ' || xml[j] == b'\n' || xml[j] == b'\t') {
                    j += 1;
                }
                if j >= xml.len() || xml[j] != b'=' {
                    i += 1;
                    continue;
                }
                j += 1;
                while j < xml.len() && (xml[j] == b' ' || xml[j] == b'\n' || xml[j] == b'\t') {
                    j += 1;
                }
                if j >= xml.len() {
                    return None;
                }
                let quote = xml[j];
                if quote != b'"' && quote != b'\'' {
                    i += 1;
                    continue;
                }
                j += 1;
                let start = j;
                while j < xml.len() && xml[j] != quote {
                    j += 1;
                }
                if j <= xml.len() {
                    return Some(String::from_utf8_lossy(&xml[start..j]).to_string());
                }
                return None;
            }
            i += 1;
        }
        None
    }

    // Fast-path: avoid XML parsing (snippets may not carry namespace declarations).
    scan_attr(xml, b"r:embed").or_else(|| scan_attr(xml, b"embed"))
}

fn parse_table_from_graphic_frame(xml: &[u8]) -> Result<Option<Vec<Vec<Cell>>>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut in_tbl = false;
    let mut in_tc = false;
    let mut in_t = false;
    let mut cur_cell = String::new();
    let mut cur_row: Vec<Cell> = Vec::new();
    let mut rows: Vec<Vec<Cell>> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"tbl" => in_tbl = true,
                b"tr" if in_tbl => {
                    cur_row = Vec::new();
                }
                b"tc" if in_tbl => {
                    in_tc = true;
                    cur_cell.clear();
                }
                b"t" if in_tbl && in_tc => in_t = true,
                _ => {}
            },
            Ok(Event::Empty(e)) => {
                if in_tbl && in_tc {
                    if e.local_name().as_ref() == b"br" {
                        if !cur_cell.ends_with('\n') {
                            cur_cell.push('\n');
                        }
                    }
                }
            }
            Ok(Event::Text(e)) => {
                if in_tbl && in_tc && in_t {
                    if let Ok(t) = String::from_utf8(e.to_vec()) {
                        cur_cell.push_str(&t);
                    }
                }
            }
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"t" => in_t = false,
                b"tc" if in_tbl => {
                    in_tc = false;
                    let txt = cur_cell.trim().to_string();
                    cur_row.push(Cell {
                        text: txt,
                        colspan: 1,
                        rowspan: 1,
                    });
                }
                b"tr" if in_tbl => {
                    if !cur_row.is_empty() {
                        rows.push(std::mem::take(&mut cur_row));
                    }
                }
                b"tbl" => {
                    break;
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("graphicFrame parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    if rows.is_empty() {
        Ok(None)
    } else {
        Ok(Some(rows))
    }
}

fn graphic_frame_has_diagram_or_chart(xml: &[u8]) -> bool {
    // OOXML: charts and SmartArt are encoded via DrawingML graphicData URIs.
    // We keep this intentionally heuristic and namespace-agnostic.
    let s = String::from_utf8_lossy(xml);
    s.contains("drawingml/2006/diagram") || s.contains("drawingml/2006/chart")
}

fn slide_xml_has_shape_diagram(slide_xml: &str) -> bool {
    // Shape-based diagrams often use connector shapes (cxnSp) and/or grouped shapes (grpSp).
    // We treat these as a signal that layout carries meaning worth rasterizing.
    slide_xml.contains("cxnSp") || slide_xml.contains("grpSp")
}

fn extract_slide_rids(presentation_xml: &str) -> Result<Vec<String>> {
    let mut reader = Reader::from_str(presentation_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut rids = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                if e.name().as_ref().ends_with(b"sldId") {
                    for a in e.attributes().flatten() {
                        let v = String::from_utf8_lossy(&a.value).to_string();
                        // Relationship ids are typically rId{N}; the slide numeric id is a number.
                        if v.starts_with("rId") || v.starts_with("rid") {
                            rids.push(v);
                            break;
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("presentation.xml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(rids)
}

#[derive(Clone, Debug)]
struct SlideMeta {
    block_first: usize,
    block_last: usize,
    has_diagram: bool,
    analysis: serde_json::Value,
}

fn flush_list_block(
    blocks: &mut Vec<Block>,
    block_index: &mut usize,
    cur_list: &mut Option<(bool, Vec<ListItem>)>,
) {
    if let Some((ordered, items)) = cur_list.take() {
        if !items.is_empty() {
            blocks.push(Block::List {
                block_index: *block_index,
                ordered,
                items,
                source: SourceSpan::default(),
            });
            *block_index += 1;
        }
    }
}

fn push_paragraph_or_list_item(
    blocks: &mut Vec<Block>,
    block_index: &mut usize,
    cur_list: &mut Option<(bool, Vec<ListItem>)>,
    is_list: bool,
    ordered: bool,
    level: u8,
    text: String,
) {
    if is_list {
        match cur_list {
            Some((cur_ord, items)) if *cur_ord == ordered => {
                items.push(ListItem {
                    level,
                    text,
                    source: SourceSpan::default(),
                });
            }
            _ => {
                flush_list_block(blocks, block_index, cur_list);
                *cur_list = Some((
                    ordered,
                    vec![ListItem {
                        level,
                        text,
                        source: SourceSpan::default(),
                    }],
                ));
            }
        }
        return;
    }

    flush_list_block(blocks, block_index, cur_list);
    blocks.push(Block::Paragraph {
        block_index: *block_index,
        text,
        source: SourceSpan::default(),
    });
    *block_index += 1;
}

fn extract_text_with_list_markers(xml: &[u8]) -> Result<Vec<(bool, bool, u8, String)>> {
    // Returns (is_list, ordered, level, text)
    fn scan_attr(tag: &[u8], key: &[u8]) -> Option<String> {
        let mut i = 0usize;
        while i + key.len() + 3 < tag.len() {
            if &tag[i..i + key.len()] == key {
                let mut j = i + key.len();
                while j < tag.len() && (tag[j] == b' ' || tag[j] == b'\n' || tag[j] == b'\t') {
                    j += 1;
                }
                if j >= tag.len() || tag[j] != b'=' {
                    i += 1;
                    continue;
                }
                j += 1;
                while j < tag.len() && (tag[j] == b' ' || tag[j] == b'\n' || tag[j] == b'\t') {
                    j += 1;
                }
                if j >= tag.len() {
                    return None;
                }
                let quote = tag[j];
                if quote != b'"' && quote != b'\'' {
                    i += 1;
                    continue;
                }
                j += 1;
                let start = j;
                while j < tag.len() && tag[j] != quote {
                    j += 1;
                }
                return Some(String::from_utf8_lossy(&tag[start..j]).to_string());
            }
            i += 1;
        }
        None
    }

    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut out: Vec<(bool, bool, u8, String)> = Vec::new();
    let mut in_p = false;
    let mut cur_text = String::new();
    let mut cur_ordered = false;
    let mut cur_bulleted = false;
    let mut cur_lvl: u8 = 0;
    let mut in_t = false;
    let mut p_depth = 0usize;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.local_name().as_ref() == b"p" {
                    in_p = true;
                    p_depth = 1;
                    cur_text.clear();
                    cur_ordered = false;
                    cur_bulleted = false;
                    cur_lvl = 0;
                } else if in_p {
                    p_depth += 1;
                    if e.local_name().as_ref() == b"pPr" {
                        if let Some(v) = scan_attr(&e.to_vec(), b"lvl") {
                            cur_lvl = v.parse::<u8>().unwrap_or(0);
                        }
                    }
                    if e.local_name().as_ref() == b"buAutoNum" {
                        cur_ordered = true;
                    }
                    if e.local_name().as_ref() == b"buChar" {
                        cur_bulleted = true;
                    }
                    if e.local_name().as_ref() == b"t" {
                        in_t = true;
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                if in_p {
                    if e.local_name().as_ref() == b"pPr" {
                        if let Some(v) =
                            scan_attr(&e.to_vec(), b"lvl").or_else(|| attr_val_local(&e, b"lvl"))
                        {
                            cur_lvl = v.parse::<u8>().unwrap_or(0);
                        }
                    }
                    if e.local_name().as_ref() == b"buAutoNum" {
                        cur_ordered = true;
                    }
                    if e.local_name().as_ref() == b"buChar" {
                        cur_bulleted = true;
                    }
                    if e.local_name().as_ref() == b"br" {
                        if !cur_text.ends_with('\n') {
                            cur_text.push('\n');
                        }
                    }
                    if e.local_name().as_ref() == b"tab" {
                        cur_text.push('\t');
                    }
                }
            }
            Ok(Event::Text(e)) => {
                if in_p && in_t {
                    if let Ok(t) = String::from_utf8(e.to_vec()) {
                        cur_text.push_str(&t);
                    }
                }
            }
            Ok(Event::End(e)) => {
                if in_p {
                    if e.local_name().as_ref() == b"t" {
                        in_t = false;
                    }
                    if e.local_name().as_ref() == b"p" {
                        let trimmed = cur_text.trim();
                        if !trimmed.is_empty() {
                            let is_list = cur_ordered || cur_bulleted;
                            out.push((is_list, cur_ordered, cur_lvl, trimmed.to_string()));
                        }
                        in_p = false;
                        p_depth = 0;
                    } else {
                        p_depth = p_depth.saturating_sub(1);
                        if p_depth == 0 {
                            in_p = false;
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("drawingml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn extract_slide_name(slide_xml: &str) -> Option<String> {
    let mut reader = Reader::from_str(slide_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"cSld" {
                    if let Some(name) = attr_val_local(&e, b"name") {
                        if !name.trim().is_empty() {
                            return Some(name);
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
    None
}

fn extract_slide_size_emu(presentation_xml: &str) -> Option<(i64, i64)> {
    let mut reader = Reader::from_str(presentation_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"sldSz" {
                    let cx = attr_val_local(&e, b"cx").and_then(|v| v.parse::<i64>().ok());
                    let cy = attr_val_local(&e, b"cy").and_then(|v| v.parse::<i64>().ok());
                    if let (Some(cx), Some(cy)) = (cx, cy) {
                        return Some((cx, cy));
                    }
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

fn extract_shape_id(xml: &[u8]) -> Option<u32> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"cNvPr" {
                    if let Some(id) = attr_val_local(&e, b"id").and_then(|v| v.parse().ok()) {
                        return Some(id);
                    }
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

fn extract_bbox_emu(xml: &[u8]) -> Option<(i64, i64, i64, i64)> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut x: Option<i64> = None;
    let mut y: Option<i64> = None;
    let mut w: Option<i64> = None;
    let mut h: Option<i64> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"off" {
                    x = attr_val_local(&e, b"x").and_then(|v| v.parse().ok());
                    y = attr_val_local(&e, b"y").and_then(|v| v.parse().ok());
                }
                if e.local_name().as_ref() == b"ext" {
                    w = attr_val_local(&e, b"cx").and_then(|v| v.parse().ok());
                    h = attr_val_local(&e, b"cy").and_then(|v| v.parse().ok());
                }
                if let (Some(x), Some(y), Some(w), Some(h)) = (x, y, w, h) {
                    return Some((x, y, w, h));
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

fn extract_connector_endpoints(xml: &[u8]) -> (Option<u32>, Option<u32>) {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut from: Option<u32> = None;
    let mut to: Option<u32> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"stCxn" {
                    from = attr_val_local(&e, b"id").and_then(|v| v.parse().ok());
                }
                if e.local_name().as_ref() == b"endCxn" {
                    to = attr_val_local(&e, b"id").and_then(|v| v.parse().ok());
                }
                if from.is_some() && to.is_some() {
                    return (from, to);
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    (from, to)
}

fn extract_text_compact(xml: &[u8]) -> anyhow::Result<Option<String>> {
    let paras = extract_text_with_list_markers(xml)?;
    let t = paras
        .iter()
        .map(|(_, _, _, t)| t.trim())
        .filter(|t| !t.is_empty())
        .collect::<Vec<_>>()
        .join(" ")
        .trim()
        .to_string();
    if t.is_empty() { Ok(None) } else { Ok(Some(t)) }
}

fn extract_prst_geom(xml: &[u8]) -> Option<String> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"prstGeom" {
                    if let Some(prst) = attr_val_local(&e, b"prst") {
                        if !prst.trim().is_empty() {
                            return Some(prst);
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
    None
}

fn mermaid_shape_from_prst(prst: Option<&str>) -> &'static str {
    let p = prst.unwrap_or("").trim();
    match p {
        "" => "rect",
        // Common shapes
        "rect" | "flowChartProcess" => "rect",
        "roundRect" | "flowChartTerminator" => "rounded",
        "ellipse" => "circle",
        "diamond" | "flowChartDecision" => "diamond",
        "hexagon" => "hex",
        // Office uses `can` for cylinder-like database shapes.
        "can" => "cyl",
        // Slanted quadrilaterals
        "parallelogram" => "lean-r",
        "trapezoid" => "trap-b",
        // Triangles (there are many variants; collapse them).
        "triangle" | "rtTriangle" | "ltTriangle" | "upTriangle" | "downTriangle" => "tri",
        // Subprocess/subroutine-ish
        "flowChartPredefinedProcess" => "subproc",
        _ => "rect",
    }
}

fn line_arrow_dir(xml: &[u8]) -> ArrowDir {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut head = false;
    let mut tail = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let ln = e.local_name();
                let ln = ln.as_ref();
                if ln == b"headEnd" {
                    if let Some(ty) = attr_val_local(&e, b"type") {
                        let t = ty.trim().to_ascii_lowercase();
                        if !t.is_empty() && t != "none" {
                            head = true;
                        }
                    }
                }
                if ln == b"tailEnd" {
                    if let Some(ty) = attr_val_local(&e, b"type") {
                        let t = ty.trim().to_ascii_lowercase();
                        if !t.is_empty() && t != "none" {
                            tail = true;
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

    match (head, tail) {
        (false, false) => ArrowDir::None,
        (true, false) => ArrowDir::Reverse,
        (false, true) => ArrowDir::Forward,
        (true, true) => ArrowDir::Both,
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ArrowDir {
    None,
    Forward,
    Reverse,
    Both,
}

fn parse_notes_text(notes_xml: &str) -> Result<Vec<(bool, bool, u8, String)>> {
    extract_text_with_list_markers(notes_xml.as_bytes())
}

fn extract_chart_rid_from_graphic_frame(xml: &[u8]) -> Option<String> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"chart" {
                    if let Some(rid) = attr_val_local(&e, b"id") {
                        if !rid.trim().is_empty() {
                            return Some(rid);
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
    None
}

fn graphic_frame_has_smartart(xml: &[u8]) -> bool {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"graphicData" {
                    if let Some(uri) = attr_val_local(&e, b"uri") {
                        if uri.contains("/diagram") {
                            return true;
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
    false
}

#[derive(Clone, Debug, Default)]
struct ChartSeriesData {
    name: Option<String>,
    values: Vec<f64>,
    format_code: Option<String>,
}

#[derive(Clone, Debug, Default)]
struct ChartDataParsed {
    title: Option<String>,
    categories: Vec<String>,
    series: Vec<ChartSeriesData>,
}

fn chart_type_from_xml(chart_xml: &str) -> Option<&'static str> {
    let mut reader = Reader::from_str(chart_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"barChart" => return Some("bar"),
                b"lineChart" => return Some("line"),
                b"pieChart" => return Some("pie"),
                b"areaChart" => return Some("area"),
                b"scatterChart" => return Some("scatter"),
                b"radarChart" => return Some("radar"),
                b"doughnutChart" => return Some("doughnut"),
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    None
}

fn score_from_reasons(reasons: &[String]) -> f32 {
    let mut s = 0.0f32;
    for r in reasons {
        s = s.max(match r.as_str() {
            "connectors" => 0.95,
            "chart" => 0.9,
            "groups" => 0.8,
            "smartart" => 0.9,
            _ => 0.7,
        });
    }
    s
}

fn parse_chart_xml_cached_values(chart_xml: &str) -> Result<ChartDataParsed> {
    let mut out: ChartDataParsed = Default::default();

    let mut reader = Reader::from_str(chart_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut in_title = false;
    let mut in_v = false;
    let mut in_title_t = false;

    let mut in_cat = false;
    let mut in_val = false;
    let mut in_cat_strcache = false;
    let mut in_val_numcache = false;
    let mut in_format_code = false;

    // Many charts repeat identical categories per-series; capture them once.
    let mut captured_categories = false;

    let mut cur_series: Option<ChartSeriesData> = None;
    let mut cur_series_name: Option<String> = None;
    let mut cur_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"ser" => {
                    if let Some(s) = cur_series.take() {
                        out.series.push(s);
                    }
                    cur_series = Some(Default::default());
                    cur_series_name = None;
                }
                b"title" => in_title = true,
                b"t" => {
                    if in_title {
                        in_title_t = true;
                        cur_text.clear();
                    }
                }
                b"cat" => {
                    in_cat = !captured_categories;
                    in_cat_strcache = false;
                }
                b"val" => {
                    in_val = true;
                    in_val_numcache = false;
                }
                b"strCache" => {
                    if in_cat {
                        in_cat_strcache = true;
                    }
                }
                b"numCache" => {
                    if in_val {
                        in_val_numcache = true;
                    }
                }
                b"formatCode" => {
                    if in_val_numcache {
                        in_format_code = true;
                        cur_text.clear();
                    }
                }
                b"numFmt" => {
                    if let Some(s) = cur_series.as_mut() {
                        if s.format_code.is_none() {
                            for a in e.attributes().flatten() {
                                if a.key.as_ref() == b"formatCode" {
                                    let code = String::from_utf8_lossy(&a.value).to_string();
                                    let code = code.trim().to_string();
                                    if !code.is_empty() {
                                        s.format_code = Some(code);
                                    }
                                }
                            }
                        }
                    }
                }
                b"v" => {
                    in_v = true;
                    cur_text.clear();
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                match e.local_name().as_ref() {
                    b"ser" => {
                        if let Some(s) = cur_series.take() {
                            out.series.push(s);
                        }
                        cur_series_name = None;
                    }
                    b"title" => in_title = false,
                    b"t" => {
                        if in_title_t {
                            in_title_t = false;
                            let v = cur_text.trim();
                            if !v.is_empty() {
                                out.title = out.title.take().or(Some(v.to_string()));
                            }
                        }
                    }
                    b"formatCode" => {
                        if in_format_code {
                            in_format_code = false;
                            let v = cur_text.trim();
                            if !v.is_empty() {
                                if cur_series.is_none() {
                                    cur_series = Some(Default::default());
                                }
                                if let Some(s) = cur_series.as_mut() {
                                    if s.format_code.is_none() {
                                        s.format_code = Some(v.to_string());
                                    }
                                }
                            }
                        }
                    }
                    b"v" => {
                        in_v = false;
                        let v = cur_text.trim();
                        if v.is_empty() {
                            // skip
                        } else if in_title {
                            // title text can appear multiple times; keep first non-empty.
                            out.title = out.title.take().or(Some(v.to_string()));
                        } else if in_cat_strcache {
                            out.categories.push(v.to_string());
                        } else if in_val_numcache {
                            if let Ok(n) = v.parse::<f64>() {
                                if cur_series.is_none() {
                                    cur_series = Some(Default::default());
                                }
                                if let Some(s) = cur_series.as_mut() {
                                    s.values.push(n);
                                }
                            }
                        } else {
                            // Likely series name.
                            if cur_series_name.is_none() {
                                cur_series_name = Some(v.to_string());
                                if cur_series.is_none() {
                                    cur_series = Some(Default::default());
                                }
                                if let Some(s) = cur_series.as_mut() {
                                    s.name = cur_series_name.clone();
                                }
                            }
                        }
                    }
                    b"strCache" => in_cat_strcache = false,
                    b"numCache" => in_val_numcache = false,
                    b"cat" => {
                        if in_cat {
                            captured_categories = true;
                        }
                        in_cat = false;
                        in_cat_strcache = false;
                    }
                    b"val" => {
                        in_val = false;
                        in_val_numcache = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(e)) => {
                if in_v || in_title_t || in_format_code {
                    cur_text.push_str(&String::from_utf8_lossy(e.as_ref()));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("chart xml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn unit_from_numfmt(code: Option<&str>) -> Option<String> {
    let code = code?.trim();
    if code.is_empty() {
        return None;
    }
    if code.contains('%') {
        return Some("%".to_string());
    }
    for sym in ["$", "€", "£", "¥", "₩", "₹"] {
        if code.contains(sym) {
            return Some(sym.to_string());
        }
    }
    if let Some(start) = code.find("[$") {
        let tail = &code[start + 2..];
        if let Some(end) = tail.find(']') {
            let inner = &tail[..end];
            if let Some(sym) = inner.chars().next() {
                if !sym.is_ascii_alphanumeric() {
                    return Some(sym.to_string());
                }
            }
        }
    }
    // e.g. 0.0"kg" -> kg
    if let Some(q1) = code.find('"') {
        let tail = &code[q1 + 1..];
        if let Some(q2) = tail.find('"') {
            let u = tail[..q2].trim();
            if !u.is_empty() {
                return Some(u.to_string());
            }
        }
    }
    None
}

fn format_value_with_unit(v: f64, unit: Option<&str>) -> String {
    let Some(u) = unit.map(str::trim).filter(|s| !s.is_empty()) else {
        return format!("{v}");
    };
    if u == "%" {
        format!("{v}%")
    } else if ["$", "€", "£", "¥", "₩", "₹"].contains(&u) {
        format!("{u}{v}")
    } else {
        format!("{v} {u}")
    }
}

fn parse_pptx_full_inner(
    bytes: &[u8],
    include_slide_snapshots: bool,
) -> Result<ParsedOfficeDocument> {
    let presentation_xml =
        read_zip_file_utf8(bytes, "ppt/presentation.xml").context("read ppt/presentation.xml")?;
    let pres_rels_xml = read_zip_file_utf8(bytes, "ppt/_rels/presentation.xml.rels")
        .context("read ppt/_rels/presentation.xml.rels")?;

    let mut slide_rids = extract_slide_rids(&presentation_xml).context("parse slide rIds")?;
    let pres_rels = parse_relationships(&pres_rels_xml).context("parse presentation rels")?;

    if slide_rids.is_empty() {
        // Fallback for cases where presentation.xml parsing fails (we still want a deterministic order).
        slide_rids = pres_rels
            .iter()
            .filter_map(|(id, (ty, _target))| {
                if ty.ends_with("/slide") {
                    Some(id.clone())
                } else {
                    None
                }
            })
            .collect();

        slide_rids.sort_by_key(|id| {
            id.trim_start_matches("rId")
                .parse::<u32>()
                .unwrap_or(u32::MAX)
        });
    }

    if slide_rids.is_empty() {
        return Err(anyhow!("no slides found in ppt/presentation.xml"));
    }

    let (_slide_cx_emu, _slide_cy_emu) =
        extract_slide_size_emu(&presentation_xml).unwrap_or((9_144_000_i64, 6_858_000_i64));

    let mut blocks: Vec<Block> = Vec::new();
    let mut images: Vec<ParsedImage> = Vec::new();
    let mut slides_meta: Vec<SlideMeta> = Vec::new();
    let mut charts_meta: Vec<serde_json::Value> = Vec::new();
    let mut diagram_graphs_meta: Vec<serde_json::Value> = Vec::new();

    let mut block_index: usize = 0;

    for (i, rid) in slide_rids.iter().enumerate() {
        let slide_idx_1 = i + 1;

        let (_ty, target) = pres_rels
            .get(rid)
            .ok_or_else(|| anyhow!("missing slide relationship {rid}"))
            .cloned()?;
        let slide_part = resolve_target("ppt/presentation.xml", &target);

        let slide_xml =
            read_zip_file_utf8(bytes, &slide_part).with_context(|| format!("read {slide_part}"))?;

        let slide_rels_part = {
            let d = dirname(&slide_part);
            let rel = format!("{}{}_rels/{}.rels", d, "", basename(&slide_part));
            // If slide_part is ppt/slides/slide1.xml, rel becomes ppt/slides/_rels/slide1.xml.rels
            // because basename includes .xml.
            // The above currently yields ppt/slides/_rels/slide1.xml.rels (since d ends with /).
            rel
        };
        let slide_rels_xml = read_zip_file_utf8(bytes, &slide_rels_part)
            .unwrap_or_else(|_| "<Relationships/>".to_string());
        let slide_rels = parse_relationships(&slide_rels_xml).unwrap_or_default();

        // Extract ordered shapes under spTree.
        let shapes = capture_sp_tree_children(&slide_xml).context("capture slide shapes")?;

        // Title detection.
        let mut title: Option<String> = None;
        let mut inferred_title = false;

        // Collect first non-empty text shape for fallback.
        let mut first_text: Option<String> = None;
        for (tag, xml) in &shapes {
            // Only consider shapes that are text boxes.
            let is_sp = tag.ends_with(":sp") || tag == "sp";
            if !is_sp {
                continue;
            }
            let ph = extract_placeholder_type(xml);
            let paras = extract_text_with_list_markers(xml)?;
            let text = paras
                .iter()
                .map(|(_, _, _, t)| t.as_str())
                .collect::<Vec<_>>()
                .join("\n")
                .trim()
                .to_string();
            if text.is_empty() {
                continue;
            }

            if first_text.is_none() {
                first_text = Some(text.clone());
            }

            if let Some(ph) = ph {
                let p = ph.to_ascii_lowercase();
                if p == "title" || p == "ctrtitle" {
                    title = Some(text);
                    break;
                }
                if p == "subtitle" {
                    title = title.or(Some(text));
                }
            }
        }

        if title.is_none() {
            title = extract_slide_name(&slide_xml);
            if title.is_none() {
                title = first_text;
                if title.is_some() {
                    inferred_title = true;
                }
            }
        }

        let slide_block_first = blocks.len();
        let cur_block_first_index = block_index;

        // Emit a leading slide title heading (always, for determinism).
        let heading_text = title
            .clone()
            .unwrap_or_else(|| format!("Slide {slide_idx_1}"));
        blocks.push(Block::Heading {
            block_index,
            level: 1,
            text: heading_text,
            source: SourceSpan::default(),
        });
        block_index += 1;

        let mut cur_list: Option<(bool, Vec<ListItem>)> = None;
        let mut slide_has_diagram = slide_xml_has_shape_diagram(&slide_xml);

        let mut text_chars: usize = 0;
        let mut embedded_images: usize = 0;
        let mut tables: usize = 0;
        let mut connectors: usize = 0;
        let mut groups: usize = 0;
        let mut charts: usize = 0;
        let mut smartart: usize = 0;
        let mut excluded_visuals: Vec<String> = Vec::new();

        #[derive(Clone, Debug)]
        enum NodeKind {
            Text,
            Image { image_id: String },
        }

        #[derive(Clone, Debug)]
        struct Node {
            id: u32,
            label: String,
            kind: NodeKind,
            mermaid_shape: &'static str,
            x: f64,
            y: f64,
            w: f64,
            h: f64,
        }

        #[derive(Clone, Debug)]
        struct Edge {
            from: Option<u32>,
            to: Option<u32>,
            label: Option<String>,
        }

        #[derive(Clone, Debug)]
        struct Line {
            x1: f64,
            y1: f64,
            x2: f64,
            y2: f64,
            dir: ArrowDir,
            label: Option<String>,
        }

        let mut nodes: Vec<Node> = Vec::new();
        let mut edges: Vec<Edge> = Vec::new();
        let mut lines: Vec<Line> = Vec::new();

        // Emit body in order.
        for (tag, xml) in &shapes {
            // Filter out common footer placeholders if they only add noise.
            if let Some(ph) = extract_placeholder_type(xml) {
                let p = ph.to_ascii_lowercase();
                if p == "dt" || p == "ftr" || p == "sldnum" || p == "hdr" {
                    continue;
                }
            }

            let is_cxn = tag.ends_with(":cxnSp") || tag == "cxnSp";
            if is_cxn {
                connectors += 1;
                slide_has_diagram = true;
                let (from, to) = extract_connector_endpoints(xml);
                let label = extract_text_compact(xml)?;
                edges.push(Edge { from, to, label });
                continue;
            }

            let is_grp = tag.ends_with(":grpSp") || tag == "grpSp";
            if is_grp {
                groups += 1;
                slide_has_diagram = true;
                continue;
            }

            let is_pic = tag.ends_with(":pic") || tag == "pic";
            if is_pic {
                if let Some(rid) = extract_first_blip_embed(xml) {
                    if let Some((_ty, target)) = slide_rels.get(&rid).cloned() {
                        let media_part = resolve_target(&slide_part, &target);
                        if let Ok(img_bytes) = read_zip_file(bytes, &media_part) {
                            embedded_images += 1;
                            let filename = Some(basename(&media_part).to_string());
                            let mime_type = filename
                                .as_deref()
                                .map(mime_from_filename)
                                .unwrap_or_else(|| "application/octet-stream".to_string());
                            let hash = sha256_hex(&img_bytes);
                            let id = format!("sha256:{hash}");
                            let alt = extract_cnvpr_alt(xml);

                            if let (Some(shape_id), Some((x, y, w, h))) =
                                (extract_shape_id(xml), extract_bbox_emu(xml))
                            {
                                const EMUS_PER_INCH: f64 = 914_400.0;
                                const PX_PER_INCH: f64 = 96.0;
                                let f = PX_PER_INCH / EMUS_PER_INCH;
                                let label = alt
                                    .as_deref()
                                    .map(str::trim)
                                    .filter(|s| !s.is_empty())
                                    .map(|s| s.to_string())
                                    .or_else(|| filename.clone())
                                    .unwrap_or_else(|| "image".to_string());
                                nodes.push(Node {
                                    id: shape_id,
                                    label,
                                    kind: NodeKind::Image {
                                        image_id: id.clone(),
                                    },
                                    mermaid_shape: "rect",
                                    x: x as f64 * f,
                                    y: y as f64 * f,
                                    w: w as f64 * f,
                                    h: h as f64 * f,
                                });
                            }
                            images.push(ParsedImage {
                                id: id.clone(),
                                bytes: img_bytes,
                                mime_type: mime_type.clone(),
                                filename: filename.clone(),
                            });
                            flush_list_block(&mut blocks, &mut block_index, &mut cur_list);
                            blocks.push(Block::Image {
                                block_index,
                                id,
                                filename,
                                content_type: Some(mime_type),
                                alt,
                                source: SourceSpan::default(),
                            });
                            block_index += 1;
                        }
                    }
                }
                continue;
            }

            let is_graphic_frame = tag.ends_with(":graphicFrame") || tag == "graphicFrame";
            if is_graphic_frame {
                if let Some(rows) = parse_table_from_graphic_frame(xml)? {
                    tables += 1;
                    flush_list_block(&mut blocks, &mut block_index, &mut cur_list);
                    blocks.push(Block::Table {
                        block_index,
                        rows,
                        source: SourceSpan::default(),
                    });
                    block_index += 1;
                } else if let Some(rid) = extract_chart_rid_from_graphic_frame(xml) {
                    charts += 1;
                    slide_has_diagram = true;

                    if let Some((_ty, target)) = slide_rels.get(&rid).cloned() {
                        let chart_part = resolve_target(&slide_part, &target);
                        if let Ok(chart_xml) = read_zip_file_utf8(bytes, &chart_part) {
                            if let Ok(chart) = parse_chart_xml_cached_values(&chart_xml) {
                                let chart_type =
                                    chart_type_from_xml(&chart_xml).unwrap_or("unknown");
                                charts_meta.push(serde_json::json!({
                                    "slide": slide_idx_1,
                                    "chart_type": chart_type,
                                    "title": chart.title,
                                    "categories": chart.categories,
                                    "series": chart.series.iter().map(|s| serde_json::json!({
                                        "name": s.name,
                                        "values": s.values,
                                        "unit": unit_from_numfmt(s.format_code.as_deref()),
                                    })).collect::<Vec<_>>(),
                                }));

                                flush_list_block(&mut blocks, &mut block_index, &mut cur_list);
                                blocks.push(Block::Heading {
                                    block_index,
                                    level: 3,
                                    text: chart
                                        .title
                                        .clone()
                                        .unwrap_or_else(|| "Chart".to_string()),
                                    source: SourceSpan::default(),
                                });
                                block_index += 1;

                                let chart_shape = if chart_type == "unknown" {
                                    "chart".to_string()
                                } else {
                                    format!("{chart_type} chart")
                                };
                                let mut units: Vec<(String, String)> = Vec::new();
                                for (i, s) in chart.series.iter().enumerate() {
                                    if let Some(u) = unit_from_numfmt(s.format_code.as_deref()) {
                                        let name = s
                                            .name
                                            .clone()
                                            .filter(|t| !t.trim().is_empty())
                                            .unwrap_or_else(|| format!("series{}", i + 1));
                                        units.push((name, u));
                                    }
                                }
                                let mut note = format!("Note: {chart_shape}");
                                if !units.is_empty() {
                                    let uniq = units
                                        .iter()
                                        .map(|(_n, u)| u.clone())
                                        .collect::<std::collections::BTreeSet<_>>();
                                    if uniq.len() == 1 {
                                        note.push_str(&format!(
                                            "; units: {}",
                                            uniq.iter().next().unwrap()
                                        ));
                                    } else {
                                        note.push_str("; units: ");
                                        note.push_str(
                                            &units
                                                .iter()
                                                .map(|(n, u)| format!("{n}={u}"))
                                                .collect::<Vec<_>>()
                                                .join(", "),
                                        );
                                    }
                                }
                                blocks.push(Block::Paragraph {
                                    block_index,
                                    text: note,
                                    source: SourceSpan::default(),
                                });
                                block_index += 1;

                                if !chart.categories.is_empty() && !chart.series.is_empty() {
                                    let mut rows: Vec<Vec<Cell>> = Vec::new();
                                    let mut hdr = Vec::new();
                                    hdr.push(Cell {
                                        text: "series".to_string(),
                                        colspan: 1,
                                        rowspan: 1,
                                    });
                                    for c in &chart.categories {
                                        hdr.push(Cell {
                                            text: c.clone(),
                                            colspan: 1,
                                            rowspan: 1,
                                        });
                                    }
                                    rows.push(hdr);

                                    for s in &chart.series {
                                        let unit = unit_from_numfmt(s.format_code.as_deref());
                                        let mut r = Vec::new();
                                        r.push(Cell {
                                            text: s
                                                .name
                                                .clone()
                                                .unwrap_or_else(|| "series".to_string()),
                                            colspan: 1,
                                            rowspan: 1,
                                        });
                                        for v in &s.values {
                                            r.push(Cell {
                                                text: format_value_with_unit(*v, unit.as_deref()),
                                                colspan: 1,
                                                rowspan: 1,
                                            });
                                        }
                                        rows.push(r);
                                    }

                                    blocks.push(Block::Table {
                                        block_index,
                                        rows,
                                        source: SourceSpan::default(),
                                    });
                                    block_index += 1;
                                }
                            }
                        }
                    }
                } else if graphic_frame_has_smartart(xml) {
                    smartart += 1;
                    slide_has_diagram = true;
                    excluded_visuals.push("smartart".to_string());
                    flush_list_block(&mut blocks, &mut block_index, &mut cur_list);
                    blocks.push(Block::Paragraph {
                        block_index,
                        text: "Note: SmartArt detected on this slide; semantic extraction not implemented; excluded.".to_string(),
                        source: SourceSpan::default(),
                    });
                    block_index += 1;
                } else if graphic_frame_has_diagram_or_chart(xml) {
                    slide_has_diagram = true;
                }
                continue;
            }

            // Text shapes
            let is_sp = tag.ends_with(":sp") || tag == "sp";
            if is_sp {
                let prst = extract_prst_geom(xml);

                // Line/arrow shapes (common for arrows between boxes).
                if prst.as_deref() == Some("line") {
                    if let Some((x, y, w, h)) = extract_bbox_emu(xml) {
                        const EMUS_PER_INCH: f64 = 914_400.0;
                        const PX_PER_INCH: f64 = 96.0;
                        let f = PX_PER_INCH / EMUS_PER_INCH;
                        let dir = line_arrow_dir(xml);
                        if dir != ArrowDir::None {
                            connectors += 1;
                            slide_has_diagram = true;
                        }
                        lines.push(Line {
                            x1: x as f64 * f,
                            y1: y as f64 * f,
                            x2: (x + w) as f64 * f,
                            y2: (y + h) as f64 * f,
                            dir,
                            label: extract_text_compact(xml)?,
                        });
                    }
                }

                if let (Some(id), Some((x, y, w, h))) =
                    (extract_shape_id(xml), extract_bbox_emu(xml))
                {
                    let paras_for_node = extract_text_with_list_markers(xml)?;
                    let full_text = paras_for_node
                        .iter()
                        .map(|(_, _, _, t)| t.trim())
                        .filter(|t| !t.is_empty())
                        .collect::<Vec<_>>()
                        .join("\n")
                        .trim()
                        .to_string();

                    if !full_text.is_empty() {
                        const EMUS_PER_INCH: f64 = 914_400.0;
                        const PX_PER_INCH: f64 = 96.0;
                        let f = PX_PER_INCH / EMUS_PER_INCH;

                        let mermaid_shape = mermaid_shape_from_prst(prst.as_deref());

                        nodes.push(Node {
                            id,
                            label: full_text,
                            kind: NodeKind::Text,
                            mermaid_shape,
                            x: x as f64 * f,
                            y: y as f64 * f,
                            w: w as f64 * f,
                            h: h as f64 * f,
                        });
                    }
                }

                let paras = extract_text_with_list_markers(xml)?;
                for (is_list, ordered, lvl, text) in paras {
                    text_chars = text_chars.saturating_add(text.chars().count());
                    push_paragraph_or_list_item(
                        &mut blocks,
                        &mut block_index,
                        &mut cur_list,
                        is_list,
                        ordered,
                        lvl,
                        text,
                    );
                }
            }
        }

        flush_list_block(&mut blocks, &mut block_index, &mut cur_list);

        // Resolve arrowed `line` shapes into edges by snapping their endpoints to the nearest nodes.
        if !nodes.is_empty() && !lines.is_empty() {
            let threshold_px: f64 = 120.0;
            let nearest = |x: f64, y: f64| -> Option<u32> {
                let mut best: Option<(u32, f64)> = None;
                for n in &nodes {
                    let cx = n.x + n.w / 2.0;
                    let cy = n.y + n.h / 2.0;
                    let dx = cx - x;
                    let dy = cy - y;
                    let d2 = dx * dx + dy * dy;
                    match best {
                        None => best = Some((n.id, d2)),
                        Some((_, cur)) if d2 < cur => best = Some((n.id, d2)),
                        _ => {}
                    }
                }
                best.and_then(|(id, d2)| {
                    if d2.sqrt() <= threshold_px {
                        Some(id)
                    } else {
                        None
                    }
                })
            };

            for l in &lines {
                if l.dir == ArrowDir::None {
                    continue;
                }
                let (sx, sy, ex, ey) = match l.dir {
                    ArrowDir::Forward | ArrowDir::Both => (l.x1, l.y1, l.x2, l.y2),
                    ArrowDir::Reverse => (l.x2, l.y2, l.x1, l.y1),
                    ArrowDir::None => continue,
                };

                let from = nearest(sx, sy);
                let to = nearest(ex, ey);
                if let (Some(from), Some(to)) = (from, to) {
                    if from != to {
                        edges.push(Edge {
                            from: Some(from),
                            to: Some(to),
                            label: l.label.clone(),
                        });
                        if l.dir == ArrowDir::Both {
                            edges.push(Edge {
                                from: Some(to),
                                to: Some(from),
                                label: l.label.clone(),
                            });
                        }
                    }
                }
            }
        }

        // Emit a Mermaid graph when we have explicit connectors/lines between nodes.
        if connectors > 0 && !nodes.is_empty() {
            let node_by_id: std::collections::HashMap<u32, &Node> =
                nodes.iter().map(|n| (n.id, n)).collect();

            let mut edge_items: Vec<(u32, u32, String)> = edges
                .iter()
                .filter_map(|e| {
                    let from = e.from?;
                    let to = e.to?;
                    let label = e.label.clone().unwrap_or_default();
                    Some((from, to, label.trim().to_string()))
                })
                .collect();
            edge_items.sort_by(|a, b| (a.0, a.1, &a.2).cmp(&(b.0, b.1, &b.2)));

            if !edge_items.is_empty() {
                let esc = |s: &str| -> String {
                    let mut out = s.replace('"', "'");
                    out = out.replace('|', "/");
                    out = out.replace('\r', "");
                    out = out.replace('\n', "<br/>");
                    let out = out.trim();
                    if out.chars().count() > 240 {
                        let mut s = out.chars().take(240).collect::<String>();
                        s.push_str("...");
                        s
                    } else {
                        out.to_string()
                    }
                };

                let mut node_ids: std::collections::BTreeSet<u32> = Default::default();
                for (from, to, _) in &edge_items {
                    node_ids.insert(*from);
                    node_ids.insert(*to);
                }

                let mut mermaid = String::new();
                mermaid.push_str("```mermaid\n");
                mermaid.push_str("flowchart LR\n");
                for id in node_ids.iter().copied() {
                    if let Some(n) = node_by_id.get(&id).copied() {
                        let label = esc(&n.label);
                        match &n.kind {
                            NodeKind::Text => {
                                mermaid.push_str(&format!(
                                    "  n{id}@{{ shape: {}, label: \"{}\" }}\n",
                                    n.mermaid_shape, label
                                ));
                            }
                            NodeKind::Image { image_id } => {
                                // Mermaid image nodes are not universally supported across renderers.
                                // Emit a plain rectangle node with the (rewritten) asset path as the label.
                                mermaid.push_str(&format!(
                                    "  n{id}@{{ shape: rect, label: \"office-image:{}\" }}\n",
                                    image_id
                                ));
                            }
                        }
                    } else {
                        mermaid.push_str(&format!("  n{id}@{{ shape: rect, label: \"n{id}\" }}\n"));
                    }
                }
                for (from, to, label) in edge_items.into_iter() {
                    if label.trim().is_empty() {
                        mermaid.push_str(&format!("  n{from} --> n{to}\n"));
                    } else {
                        let l = esc(&label);
                        mermaid.push_str(&format!("  n{from} -->|{l}| n{to}\n"));
                    }
                }
                mermaid.push_str("```\n");

                blocks.push(Block::Heading {
                    block_index,
                    level: 3,
                    text: "Diagram".to_string(),
                    source: SourceSpan::default(),
                });
                block_index += 1;
                blocks.push(Block::Paragraph {
                    block_index,
                    text: mermaid,
                    source: SourceSpan::default(),
                });
                block_index += 1;
            }
        }

        let mut reasons: Vec<String> = Vec::new();
        if charts > 0 {
            reasons.push("chart".to_string());
        }
        if connectors > 0 {
            reasons.push("connectors".to_string());
        }
        if groups > 0 {
            reasons.push("groups".to_string());
        }
        if smartart > 0 {
            reasons.push("smartart".to_string());
        }
        reasons.sort();
        reasons.dedup();
        let needs_semantics = !reasons.is_empty();
        let score = score_from_reasons(&reasons);

        let snapshot_image_id: Option<String> = None;
        let _ = include_slide_snapshots;

        // Speaker notes.
        let notes_target = slide_rels.iter().find_map(|(_id, (ty, target))| {
            if ty == RELTYPE_NOTES_SLIDE {
                Some(target.clone())
            } else {
                None
            }
        });

        if let Some(target) = notes_target {
            let notes_part = resolve_target(&slide_part, &target);
            if let Ok(notes_xml) = read_zip_file_utf8(bytes, &notes_part) {
                let notes_paras = parse_notes_text(&notes_xml)?;
                if !notes_paras.is_empty() {
                    blocks.push(Block::Heading {
                        block_index,
                        level: 3,
                        text: "Speaker notes".to_string(),
                        source: SourceSpan::default(),
                    });
                    block_index += 1;

                    let mut cur_notes_list: Option<(bool, Vec<ListItem>)> = None;
                    for (is_list, ordered, lvl, text) in notes_paras {
                        push_paragraph_or_list_item(
                            &mut blocks,
                            &mut block_index,
                            &mut cur_notes_list,
                            is_list,
                            ordered,
                            lvl,
                            text,
                        );
                    }
                    flush_list_block(&mut blocks, &mut block_index, &mut cur_notes_list);
                }
            }
        }

        let slide_block_last = blocks.len().saturating_sub(1);
        let slide_last_block_index = block_index.saturating_sub(1);

        let graph_json = if connectors > 0 && !nodes.is_empty() {
            let mut warnings: Vec<String> = Vec::new();
            let mut complete_edges: Vec<(u32, u32, Option<String>)> = Vec::new();
            for e in &edges {
                match (e.from, e.to) {
                    (Some(f), Some(t)) => complete_edges.push((f, t, e.label.clone())),
                    _ => warnings.push("connector missing endpoint".to_string()),
                }
            }

            let mut used: std::collections::HashSet<u32> = Default::default();
            for (f, t, _label) in &complete_edges {
                used.insert(*f);
                used.insert(*t);
            }

            let nodes_json = nodes
                .iter()
                .map(|n| {
                    let kind = match &n.kind {
                        NodeKind::Text => "text",
                        NodeKind::Image { .. } => "image",
                    };
                    serde_json::json!({
                        "id": format!("n{}", n.id),
                        "text": n.label.clone(),
                        "kind": kind,
                        "bbox": {"x": n.x, "y": n.y, "w": n.w, "h": n.h},
                    })
                })
                .collect::<Vec<_>>();
            complete_edges.sort_by(|a, b| (a.0, a.1).cmp(&(b.0, b.1)));
            let edges_json = complete_edges
                .into_iter()
                .map(|(from, to, label)| {
                    serde_json::json!({
                        "from": format!("n{from}"),
                        "to": format!("n{to}"),
                        "kind": "connector",
                        "label": label
                    })
                })
                .collect::<Vec<_>>();

            let unattached_text = nodes
                .iter()
                .filter(|n| !used.contains(&n.id))
                .map(|n| {
                    let kind = match &n.kind {
                        NodeKind::Text => "text",
                        NodeKind::Image { .. } => "image",
                    };
                    serde_json::json!({
                        "id": format!("n{}", n.id),
                        "text": n.label.clone(),
                        "kind": kind,
                        "bbox": {"x": n.x, "y": n.y, "w": n.w, "h": n.h},
                    })
                })
                .collect::<Vec<_>>();

            let g = serde_json::json!({
                "slide": slide_idx_1,
                "nodes": nodes_json,
                "edges": edges_json,
                "unattached_text": unattached_text,
                "warnings": warnings
            });
            diagram_graphs_meta.push(g.clone());
            Some(g)
        } else {
            None
        };

        slides_meta.push(SlideMeta {
            block_first: cur_block_first_index,
            block_last: slide_last_block_index,
            has_diagram: slide_has_diagram,
            analysis: serde_json::json!({
                "index": slide_idx_1,
                "title": title,
                "text_chars": text_chars,
                "text_boxes": nodes.len(),
                "images": embedded_images,
                "tables": tables,
                "connectors": connectors,
                "groups": groups,
                "charts": charts,
                "smartart": smartart,
                "needs_semantics": needs_semantics,
                "reasons": reasons,
                "excluded_visuals": excluded_visuals,
                "score": score,
                "snapshot_image_id": snapshot_image_id,
                "diagram_graph": graph_json,
            }),
        });

        // Also keep an internal sanity check: slide should have at least the heading.
        if slide_block_last < slide_block_first {
            return Err(anyhow!("slide {slide_idx_1} produced no blocks"));
        }

        // Mark whether the title was inferred.
        if inferred_title {
            // We'll add this at the top-level metadata.
        }
    }

    let metadata_json = serde_json::json!({
        "kind": "pptx",
        "slides": slides_meta.iter().map(|s| {
            let mut v = s.analysis.clone();
            if let Some(o) = v.as_object_mut() {
                o.insert("block_first".to_string(), serde_json::json!(s.block_first));
                o.insert("block_last".to_string(), serde_json::json!(s.block_last));
                o.insert("has_diagram".to_string(), serde_json::json!(s.has_diagram));
            }
            v
        }).collect::<Vec<_>>(),
        "charts": charts_meta,
        "diagram_graphs": diagram_graphs_meta,
    });

    Ok(ParsedOfficeDocument {
        blocks,
        images,
        metadata_json,
    })
}

pub fn parse_pptx_full(bytes: &[u8]) -> Result<ParsedOfficeDocument> {
    parse_pptx_full_inner(bytes, true)
}

pub fn parse(bytes: &[u8]) -> crate::Result<crate::Document> {
    let parsed = parse_pptx_full_inner(bytes, true)
        .map_err(|e| crate::Error::from_parse(crate::Format::Pptx, e))?;
    Ok(super::finalize(crate::Format::Pptx, parsed))
}

pub fn parse_with_options(
    bytes: &[u8],
    opts: crate::pptx::ParseOptions,
) -> crate::Result<crate::Document> {
    let parsed = parse_pptx_full_inner(bytes, opts.include_slide_snapshots)
        .map_err(|e| crate::Error::from_parse(crate::Format::Pptx, e))?;
    Ok(super::finalize(crate::Format::Pptx, parsed))
}
