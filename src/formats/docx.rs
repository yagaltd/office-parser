use std::collections::HashMap;

use anyhow::{Result, anyhow};
use quick_xml::Reader;
use quick_xml::events::Event;

use crate::document_ast::{Block, Cell, LinkKind, ListItem, SourceSpan};

use super::{
    ParsedImage, ParsedOfficeDocument, mime_from_filename, read_zip_file, read_zip_file_utf8,
    sha256_hex,
};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct ListCtx {
    num_id: u32,
    ilvl: u8,
    ordered: bool,
}

#[derive(Clone, Debug, Default)]
struct NumberingIndex {
    num_to_abs: HashMap<u32, u32>,
    abs_lvl_ordered: HashMap<(u32, u8), bool>,
}

impl NumberingIndex {
    fn ordered_for(&self, num_id: u32, ilvl: u8) -> Option<bool> {
        let abs = *self.num_to_abs.get(&num_id)?;
        self.abs_lvl_ordered.get(&(abs, ilvl)).copied()
    }
}

fn attr_val(e: &quick_xml::events::BytesStart<'_>, key_local: &[u8]) -> Option<String> {
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
    // Id -> (type, target)
    let mut out = HashMap::new();
    let mut reader = Reader::from_str(rels_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    let id = attr_val(&e, b"Id").unwrap_or_default();
                    let ty = attr_val(&e, b"Type").unwrap_or_default();
                    let target = attr_val(&e, b"Target").unwrap_or_default();
                    if !id.is_empty() {
                        out.insert(id, (ty, target));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("rels xml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn extract_docx_title(bytes: &[u8]) -> Option<String> {
    let core_xml = read_zip_file_utf8(bytes, "docProps/core.xml").ok()?;
    let mut reader = Reader::from_str(&core_xml);
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

fn parse_styles_heading_levels(styles_xml: &str) -> Result<HashMap<String, u8>> {
    // styleId -> heading level (1..)
    let mut out = HashMap::new();
    let mut reader = Reader::from_str(styles_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut cur_style_id: Option<String> = None;
    let mut cur_name: Option<String> = None;
    let mut cur_outline: Option<u8> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"style" {
                    let ty = attr_val(&e, b"type").unwrap_or_default();
                    if ty == "paragraph" {
                        cur_style_id = attr_val(&e, b"styleId");
                        cur_name = None;
                        cur_outline = None;
                    }
                }
                if cur_style_id.is_some() {
                    if e.local_name().as_ref() == b"name" {
                        cur_name = attr_val(&e, b"val");
                    }
                    if e.local_name().as_ref() == b"outlineLvl" {
                        if let Some(v) = attr_val(&e, b"val") {
                            if let Ok(n) = v.parse::<u8>() {
                                cur_outline = Some(n.saturating_add(1));
                            }
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                if e.local_name().as_ref() == b"style" {
                    if let Some(id) = cur_style_id.take() {
                        let mut lvl = cur_outline;

                        if lvl.is_none() {
                            let id_l = id.to_ascii_lowercase();
                            if let Some(n) = id_l.strip_prefix("heading") {
                                if let Ok(d) = n.parse::<u8>() {
                                    lvl = Some(d.max(1));
                                }
                            }
                        }

                        if lvl.is_none() {
                            if let Some(name) = cur_name.as_deref() {
                                let nl = name.to_ascii_lowercase();
                                if let Some(rest) = nl.strip_prefix("heading") {
                                    let rest = rest.trim();
                                    if let Ok(d) = rest.parse::<u8>() {
                                        lvl = Some(d.max(1));
                                    }
                                }
                            }
                        }

                        if let Some(l) = lvl {
                            out.insert(id, l);
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("styles xml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_numbering(numbering_xml: &str) -> Result<NumberingIndex> {
    let mut idx = NumberingIndex::default();

    let mut reader = Reader::from_str(numbering_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut cur_abs: Option<u32> = None;
    let mut cur_ilvl: Option<u8> = None;
    let mut in_abs = false;
    let mut in_num = false;
    let mut cur_num_id: Option<u32> = None;
    let mut cur_num_abs: Option<u32> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let ln = e.local_name();
                if ln.as_ref() == b"abstractNum" {
                    in_abs = true;
                    cur_abs = attr_val(&e, b"abstractNumId").and_then(|s| s.parse().ok());
                }
                if ln.as_ref() == b"num" {
                    in_num = true;
                    cur_num_id = attr_val(&e, b"numId").and_then(|s| s.parse().ok());
                    cur_num_abs = None;
                }
                if in_abs && ln.as_ref() == b"lvl" {
                    cur_ilvl = attr_val(&e, b"ilvl").and_then(|s| s.parse().ok());
                }
                if in_abs && ln.as_ref() == b"numFmt" {
                    if let (Some(abs), Some(ilvl)) = (cur_abs, cur_ilvl) {
                        let fmt = attr_val(&e, b"val")
                            .unwrap_or_default()
                            .to_ascii_lowercase();
                        let ordered = fmt != "bullet";
                        idx.abs_lvl_ordered.insert((abs, ilvl), ordered);
                    }
                }
                if in_num && ln.as_ref() == b"abstractNumId" {
                    cur_num_abs = attr_val(&e, b"val").and_then(|s| s.parse().ok());
                }
            }
            Ok(Event::End(e)) => {
                let ln = e.local_name();
                if ln.as_ref() == b"lvl" {
                    cur_ilvl = None;
                }
                if ln.as_ref() == b"abstractNum" {
                    in_abs = false;
                    cur_abs = None;
                    cur_ilvl = None;
                }
                if ln.as_ref() == b"num" {
                    in_num = false;
                    if let (Some(num_id), Some(abs_id)) = (cur_num_id.take(), cur_num_abs.take()) {
                        idx.num_to_abs.insert(num_id, abs_id);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("numbering xml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(idx)
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

fn extract_body_children_xml(document_xml: &[u8]) -> Result<Vec<(String, Vec<u8>)>> {
    // Returns (kind, xml_bytes) where kind is "p" or "tbl".
    let mut out = Vec::new();

    let mut reader = Reader::from_reader(document_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut in_body = false;
    let mut body_depth = 0usize;
    let mut capturing: Option<(String, Vec<u8>, usize, Vec<u8>)> = None;
    // (kind, bytes, depth, start_tag_name)

    loop {
        let ev = reader.read_event_into(&mut buf);
        match ev {
            Ok(Event::Start(e)) => {
                let ln = e.local_name();
                if ln.as_ref() == b"body" {
                    in_body = true;
                    body_depth = 0;
                } else if in_body {
                    if capturing.is_none() {
                        body_depth += 1;
                    }

                    if capturing.is_none() && body_depth == 1 {
                        if ln.as_ref() == b"p" {
                            let mut v = Vec::with_capacity(4096);
                            write_start_tag(&mut v, &e);
                            capturing = Some(("p".to_string(), v, 1, e.name().as_ref().to_vec()));
                            buf.clear();
                            continue;
                        } else if ln.as_ref() == b"tbl" {
                            let mut v = Vec::with_capacity(8192);
                            write_start_tag(&mut v, &e);
                            capturing = Some(("tbl".to_string(), v, 1, e.name().as_ref().to_vec()));
                            buf.clear();
                            continue;
                        }
                    }

                    if let Some((_k, v, depth, _name)) = capturing.as_mut() {
                        *depth += 1;
                        write_start_tag(v, &e);
                    }
                } else if let Some((_k, v, depth, _name)) = capturing.as_mut() {
                    *depth += 1;
                    write_start_tag(v, &e);
                }
            }
            Ok(Event::Empty(e)) => {
                if let Some((_k, v, _depth, _name)) = capturing.as_mut() {
                    write_empty_tag(v, &e);
                }
            }
            Ok(Event::Text(e)) => {
                if let Some((_k, v, _depth, _name)) = capturing.as_mut() {
                    v.extend_from_slice(e.as_ref());
                }
            }
            Ok(Event::End(e)) => {
                let ln = e.local_name();
                if ln.as_ref() == b"body" {
                    in_body = false;
                }

                if let Some((k, mut v, depth, name)) = capturing.take() {
                    write_end_tag(&mut v, e.name().as_ref());
                    let depth = depth.saturating_sub(1);
                    if depth == 0 && e.name().as_ref() == name.as_slice() {
                        out.push((k, v));
                        if in_body {
                            body_depth = body_depth.saturating_sub(1);
                        }
                    } else {
                        capturing = Some((k, v, depth, name));
                    }
                } else if in_body {
                    body_depth = body_depth.saturating_sub(1);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("docx document.xml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

#[derive(Clone, Debug, Default)]
struct ParagraphParsed {
    text: String,
    heading_level: Option<u8>,
    list_ctx: Option<ListCtx>,
    links: Vec<(String, Option<String>)>,
    image_rel_ids: Vec<(String, Option<String>)>,
    chart_rel_ids: Vec<String>,
    drawings: Vec<Vec<u8>>,
}

fn parse_paragraph_xml(
    p_xml: &[u8],
    heading_styles: &HashMap<String, u8>,
    numbering: &NumberingIndex,
) -> Result<ParagraphParsed> {
    let mut out = ParagraphParsed::default();
    let mut reader = Reader::from_reader(p_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut in_t = false;
    let mut cur_p_style: Option<String> = None;
    let mut cur_outline: Option<u8> = None;
    let mut cur_num_id: Option<u32> = None;
    let mut cur_ilvl: Option<u8> = None;

    let mut in_hyperlink = false;
    let mut cur_hyperlink_rid: Option<String> = None;
    let mut cur_hyperlink_text = String::new();

    let mut pending_alt: Option<String> = None;

    // Capture full <w:drawing> subtrees so we can do diagram extraction later.
    let mut capturing_drawing: Option<(Vec<u8>, usize, Vec<u8>)> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let ln = e.local_name();

                if let Some((v, d, _n)) = capturing_drawing.as_mut() {
                    *d += 1;
                    write_start_tag(v, &e);
                }

                if capturing_drawing.is_none() && ln.as_ref() == b"drawing" {
                    let mut v = Vec::with_capacity(2048);
                    write_start_tag(&mut v, &e);
                    capturing_drawing = Some((v, 1, e.name().as_ref().to_vec()));
                }

                if ln.as_ref() == b"t" {
                    in_t = true;
                }
                if ln.as_ref() == b"pStyle" {
                    cur_p_style = attr_val(&e, b"val");
                }
                if ln.as_ref() == b"outlineLvl" {
                    if let Some(v) = attr_val(&e, b"val") {
                        if let Ok(n) = v.parse::<u8>() {
                            cur_outline = Some(n.saturating_add(1));
                        }
                    }
                }
                if ln.as_ref() == b"hyperlink" {
                    in_hyperlink = true;
                    cur_hyperlink_rid = attr_val(&e, b"id");
                    cur_hyperlink_text.clear();
                }
            }
            Ok(Event::Empty(e)) => {
                let ln = e.local_name();

                if let Some((v, _d, _n)) = capturing_drawing.as_mut() {
                    write_empty_tag(v, &e);
                }
                if ln.as_ref() == b"pStyle" {
                    cur_p_style = attr_val(&e, b"val");
                }
                if ln.as_ref() == b"outlineLvl" {
                    if let Some(v) = attr_val(&e, b"val") {
                        if let Ok(n) = v.parse::<u8>() {
                            cur_outline = Some(n.saturating_add(1));
                        }
                    }
                }
                if ln.as_ref() == b"tab" {
                    out.text.push('\t');
                    if in_hyperlink {
                        cur_hyperlink_text.push('\t');
                    }
                }
                if ln.as_ref() == b"br" || ln.as_ref() == b"cr" {
                    out.text.push('\n');
                    if in_hyperlink {
                        cur_hyperlink_text.push('\n');
                    }
                }
                if ln.as_ref() == b"numId" {
                    if let Some(v) = attr_val(&e, b"val") {
                        cur_num_id = v.parse::<u32>().ok();
                    }
                }
                if ln.as_ref() == b"ilvl" {
                    if let Some(v) = attr_val(&e, b"val") {
                        cur_ilvl = v.parse::<u8>().ok();
                    }
                }
                if ln.as_ref() == b"blip" {
                    if let Some(rid) = attr_val(&e, b"embed") {
                        out.image_rel_ids.push((rid, pending_alt.take()));
                    }
                }
                if ln.as_ref() == b"chart" {
                    if let Some(rid) = attr_val(&e, b"id") {
                        if !rid.trim().is_empty() {
                            out.chart_rel_ids.push(rid);
                        }
                    }
                }
                if ln.as_ref() == b"docPr" {
                    pending_alt = attr_val(&e, b"descr").or_else(|| attr_val(&e, b"title"));
                }
            }
            Ok(Event::Text(e)) => {
                if let Some((v, _d, _n)) = capturing_drawing.as_mut() {
                    v.extend_from_slice(e.as_ref());
                }
                if in_t {
                    let t = e.decode().map_err(|ee| anyhow!("docx text decode: {ee}"))?;
                    out.text.push_str(&t);
                    if in_hyperlink {
                        cur_hyperlink_text.push_str(&t);
                    }
                }
            }
            Ok(Event::End(e)) => {
                let ln = e.local_name();

                if let Some((mut v, d, n)) = capturing_drawing.take() {
                    write_end_tag(&mut v, e.name().as_ref());
                    let d = d.saturating_sub(1);
                    if d == 0 && e.name().as_ref() == n.as_slice() {
                        out.drawings.push(v);
                    } else {
                        capturing_drawing = Some((v, d, n));
                    }
                }

                if ln.as_ref() == b"t" {
                    in_t = false;
                }
                if ln.as_ref() == b"hyperlink" {
                    in_hyperlink = false;
                    if let Some(rid) = cur_hyperlink_rid.take() {
                        let display = cur_hyperlink_text.trim().to_string();
                        out.links.push((
                            rid,
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
            Err(e) => return Err(anyhow!("docx paragraph xml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    let heading_level = cur_p_style
        .as_ref()
        .and_then(|s| heading_styles.get(s).copied())
        .or(cur_outline);
    out.heading_level = heading_level;

    if let (Some(num_id), Some(ilvl)) = (cur_num_id, cur_ilvl) {
        let ordered = numbering.ordered_for(num_id, ilvl).unwrap_or(false);
        out.list_ctx = Some(ListCtx {
            num_id,
            ilvl,
            ordered,
        });
    }

    out.text = out.text.trim().to_string();
    Ok(out)
}

#[derive(Clone, Debug, Default)]
struct TableParsed {
    rows: Vec<Vec<Cell>>,
}

fn parse_table_xml(tbl_xml: &[u8]) -> Result<TableParsed> {
    let mut out = TableParsed::default();
    let mut reader = Reader::from_reader(tbl_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut in_tc = false;
    let mut cell_text = String::new();
    let mut cell_colspan: usize = 1;
    let mut vmerge_mode: Option<bool> = None;
    // None = no vMerge, Some(true)=restart, Some(false)=continue

    let mut cur_row: Vec<Cell> = Vec::new();
    let mut in_t = false;

    // Track vertical merges by grid column position.
    let mut active_vmerge: Vec<Option<(usize, usize)>> = Vec::new();
    let mut row_idx: usize = 0;
    let mut cur_col: usize = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let ln = e.local_name();
                if ln.as_ref() == b"tr" {
                    cur_row = Vec::new();
                    cur_col = 0;
                }
                if ln.as_ref() == b"tc" {
                    in_tc = true;
                    cell_text.clear();
                    cell_colspan = 1;
                    vmerge_mode = None;
                }
                if in_tc && ln.as_ref() == b"t" {
                    in_t = true;
                }
            }
            Ok(Event::Empty(e)) => {
                let ln = e.local_name();
                if in_tc {
                    if ln.as_ref() == b"gridSpan" {
                        if let Some(v) = attr_val(&e, b"val") {
                            cell_colspan = v.parse::<usize>().unwrap_or(1).max(1);
                        }
                    }
                    if ln.as_ref() == b"vMerge" {
                        let val = attr_val(&e, b"val")
                            .unwrap_or_default()
                            .to_ascii_lowercase();
                        if val == "restart" {
                            vmerge_mode = Some(true);
                        } else {
                            // Missing val means "continue".
                            vmerge_mode = Some(false);
                        }
                    }
                    if ln.as_ref() == b"tab" {
                        cell_text.push('\t');
                    }
                    if ln.as_ref() == b"br" || ln.as_ref() == b"cr" {
                        cell_text.push('\n');
                    }
                }
            }
            Ok(Event::Text(e)) => {
                if in_tc && in_t {
                    let t = e
                        .decode()
                        .map_err(|ee| anyhow!("docx table text decode: {ee}"))?;
                    cell_text.push_str(&t);
                }
            }
            Ok(Event::End(e)) => {
                let ln = e.local_name();
                if ln.as_ref() == b"t" {
                    in_t = false;
                }
                if ln.as_ref() == b"tc" {
                    in_tc = false;
                    let txt = cell_text.trim().to_string();
                    let colspan = cell_colspan;
                    if active_vmerge.len() < cur_col + colspan {
                        active_vmerge.resize(cur_col + colspan, None);
                    }

                    match vmerge_mode {
                        Some(false) => {
                            // Continue: increment starting cell's rowspan.
                            for c in cur_col..(cur_col + colspan) {
                                if let Some((sr, sc)) = active_vmerge.get(c).and_then(|x| *x) {
                                    if let Some(row) = out.rows.get_mut(sr) {
                                        if let Some(cell) = row.get_mut(sc) {
                                            cell.rowspan = cell.rowspan.saturating_add(1);
                                        }
                                    }
                                }
                            }
                            cur_row.push(Cell {
                                text: txt,
                                colspan,
                                rowspan: 0,
                            });
                        }
                        Some(true) => {
                            // Restart: set new active merge owner.
                            let cell_idx = cur_row.len();
                            cur_row.push(Cell {
                                text: txt,
                                colspan,
                                rowspan: 1,
                            });
                            for c in cur_col..(cur_col + colspan) {
                                active_vmerge[c] = Some((row_idx, cell_idx));
                            }
                        }
                        None => {
                            // No merge: clear any active merges for these columns.
                            for c in cur_col..(cur_col + colspan) {
                                active_vmerge[c] = None;
                            }
                            cur_row.push(Cell {
                                text: txt,
                                colspan,
                                rowspan: 1,
                            });
                        }
                    }
                    cur_col += colspan;
                }
                if ln.as_ref() == b"tr" {
                    out.rows.push(cur_row.clone());
                    row_idx = out.rows.len();
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("docx table xml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
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
            Ok(Event::End(e)) => match e.local_name().as_ref() {
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
            },
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

fn extract_prst_geom(xml: &[u8]) -> Option<String> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"prstGeom" {
                    for a in e.attributes().flatten() {
                        if a.key
                            .as_ref()
                            .split(|b| *b == b':')
                            .last()
                            .unwrap_or(a.key.as_ref())
                            == b"prst"
                        {
                            let v = String::from_utf8_lossy(&a.value).to_string();
                            let v = v.trim();
                            if !v.is_empty() {
                                return Some(v.to_string());
                            }
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

fn extract_shape_id(xml: &[u8]) -> Option<u32> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"cNvPr" {
                    for a in e.attributes().flatten() {
                        if a.key
                            .as_ref()
                            .split(|b| *b == b':')
                            .last()
                            .unwrap_or(a.key.as_ref())
                            == b"id"
                        {
                            let v = String::from_utf8_lossy(&a.value).to_string();
                            if let Ok(n) = v.trim().parse::<u32>() {
                                return Some(n);
                            }
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
                let ln = e.local_name();
                if ln.as_ref() == b"off" {
                    for a in e.attributes().flatten() {
                        let k = a
                            .key
                            .as_ref()
                            .split(|b| *b == b':')
                            .last()
                            .unwrap_or(a.key.as_ref());
                        let v = String::from_utf8_lossy(&a.value).to_string();
                        if k == b"x" {
                            x = v.trim().parse::<i64>().ok();
                        }
                        if k == b"y" {
                            y = v.trim().parse::<i64>().ok();
                        }
                    }
                }
                if ln.as_ref() == b"ext" {
                    for a in e.attributes().flatten() {
                        let k = a
                            .key
                            .as_ref()
                            .split(|b| *b == b':')
                            .last()
                            .unwrap_or(a.key.as_ref());
                        let v = String::from_utf8_lossy(&a.value).to_string();
                        if k == b"cx" {
                            w = v.trim().parse::<i64>().ok();
                        }
                        if k == b"cy" {
                            h = v.trim().parse::<i64>().ok();
                        }
                    }
                }
                if x.is_some() && y.is_some() && w.is_some() && h.is_some() {
                    return Some((x?, y?, w?, h?));
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

fn extract_text_compact(xml: &[u8]) -> Result<Option<String>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut in_t = false;
    let mut out = String::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let ln = e.local_name();
                if ln.as_ref() == b"t" {
                    in_t = true;
                }
            }
            Ok(Event::Text(e)) => {
                if in_t {
                    let t = e.decode().map_err(|ee| anyhow!("docx text decode: {ee}"))?;
                    let t = t.trim();
                    if !t.is_empty() {
                        if !out.is_empty() {
                            out.push(' ');
                        }
                        out.push_str(t);
                    }
                }
            }
            Ok(Event::End(e)) => {
                if e.local_name().as_ref() == b"t" {
                    in_t = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    let out = out.trim().to_string();
    Ok(if out.is_empty() { None } else { Some(out) })
}

fn extract_first_blip_embed(xml: &[u8]) -> Option<String> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"blip" {
                    for a in e.attributes().flatten() {
                        let k = a
                            .key
                            .as_ref()
                            .split(|b| *b == b':')
                            .last()
                            .unwrap_or(a.key.as_ref());
                        if k == b"embed" {
                            let v = String::from_utf8_lossy(&a.value).to_string();
                            let v = v.trim();
                            if !v.is_empty() {
                                return Some(v.to_string());
                            }
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

fn extract_connector_endpoints(xml: &[u8]) -> (Option<u32>, Option<u32>) {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut from: Option<u32> = None;
    let mut to: Option<u32> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"stCxn" => {
                    if from.is_none() {
                        from = attr_val(&e, b"id").and_then(|s| s.trim().parse::<u32>().ok());
                    }
                }
                b"endCxn" => {
                    if to.is_none() {
                        to = attr_val(&e, b"id").and_then(|s| s.trim().parse::<u32>().ok());
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        if from.is_some() && to.is_some() {
            break;
        }
        buf.clear();
    }
    (from, to)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ArrowDir {
    None,
    Forward,
    Reverse,
    Both,
}

fn line_arrow_dir(xml: &[u8]) -> ArrowDir {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut head = false;
    let mut tail = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"headEnd" => head = true,
                b"tailEnd" => tail = true,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    match (tail, head) {
        (false, false) => ArrowDir::None,
        (false, true) => ArrowDir::Forward,
        (true, false) => ArrowDir::Reverse,
        (true, true) => ArrowDir::Both,
    }
}

#[derive(Clone, Debug)]
enum DiagramNodeKind {
    Text,
    Image { image_id: String },
}

#[derive(Clone, Debug)]
struct DiagramNode {
    id: u32,
    label: String,
    kind: DiagramNodeKind,
    mermaid_shape: &'static str,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
}

#[derive(Clone, Debug)]
struct DiagramEdge {
    from: Option<u32>,
    to: Option<u32>,
    dir: ArrowDir,
    label: Option<String>,
}

#[derive(Clone, Debug)]
struct DiagramLine {
    x1: f64,
    y1: f64,
    x2: f64,
    y2: f64,
    dir: ArrowDir,
    label: Option<String>,
}

fn diagram_from_drawing_xml(
    drawing_xml: &[u8],
    bytes: &[u8],
    rels: &HashMap<String, (String, String)>,
    images: &mut Vec<ParsedImage>,
    image_by_hash: &mut HashMap<String, usize>,
) -> Result<Option<(String, serde_json::Value)>> {
    let mut reader = Reader::from_reader(drawing_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    // Capture individual shapes/connectors so we can parse them independently.
    let mut capturing: Option<(String, Vec<u8>, usize, Vec<u8>)> = None;
    let mut parts: Vec<(String, Vec<u8>)> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if let Some((t, v, d, n)) = capturing.as_mut() {
                    *d += 1;
                    write_start_tag(v, &e);
                    let _ = (t, n);
                } else {
                    match e.local_name().as_ref() {
                        b"wsp" | b"cxnSp" => {
                            let mut v = Vec::with_capacity(4096);
                            write_start_tag(&mut v, &e);
                            capturing = Some((
                                String::from_utf8_lossy(e.name().as_ref()).to_string(),
                                v,
                                1,
                                e.name().as_ref().to_vec(),
                            ));
                        }
                        _ => {}
                    }
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
                if let Some((t, mut v, d, n)) = capturing.take() {
                    write_end_tag(&mut v, e.name().as_ref());
                    let d = d.saturating_sub(1);
                    if d == 0 && e.name().as_ref() == n.as_slice() {
                        parts.push((t, v));
                    } else {
                        capturing = Some((t, v, d, n));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("docx drawing xml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    if parts.is_empty() {
        return Ok(None);
    }

    let mut nodes: Vec<DiagramNode> = Vec::new();
    let mut edges: Vec<DiagramEdge> = Vec::new();
    let mut lines: Vec<DiagramLine> = Vec::new();
    let mut connectors: usize = 0;

    for (tag, xml) in &parts {
        let local = tag.split(':').last().unwrap_or(tag);
        if local == "cxnSp" {
            connectors += 1;
            let (from, to) = extract_connector_endpoints(xml);
            let label = extract_text_compact(xml)?;
            edges.push(DiagramEdge {
                from,
                to,
                dir: ArrowDir::Forward,
                label,
            });
            continue;
        }

        // Shape
        let prst = extract_prst_geom(xml);
        if prst.as_deref() == Some("line") {
            if let Some((x, y, w, h)) = extract_bbox_emu(xml) {
                const EMUS_PER_INCH: f64 = 914_400.0;
                const PX_PER_INCH: f64 = 96.0;
                let f = PX_PER_INCH / EMUS_PER_INCH;
                let x1 = x as f64 * f;
                let y1 = y as f64 * f;
                let x2 = (x + w) as f64 * f;
                let y2 = (y + h) as f64 * f;
                lines.push(DiagramLine {
                    x1,
                    y1,
                    x2,
                    y2,
                    dir: line_arrow_dir(xml),
                    label: extract_text_compact(xml)?,
                });
            }
            continue;
        }

        let id = extract_shape_id(xml);
        let bbox = extract_bbox_emu(xml);
        if let (Some(id), Some((x, y, w, h))) = (id, bbox) {
            const EMUS_PER_INCH: f64 = 914_400.0;
            const PX_PER_INCH: f64 = 96.0;
            let f = PX_PER_INCH / EMUS_PER_INCH;

            // Image node?
            if let Some(rid) = extract_first_blip_embed(xml) {
                if let Some((_ty, target)) = rels.get(&rid) {
                    let entry = resolve_target("word/", target);
                    if let Ok(img_bytes) = read_zip_file(bytes, &entry) {
                        let hash = sha256_hex(&img_bytes);
                        let _idx = *image_by_hash.entry(hash.clone()).or_insert_with(|| {
                            let filename = entry.split('/').last().map(|s| s.to_string());
                            let mime = filename
                                .as_deref()
                                .map(mime_from_filename)
                                .unwrap_or_else(|| "application/octet-stream".to_string());
                            images.push(ParsedImage {
                                id: format!("sha256:{hash}"),
                                bytes: img_bytes,
                                mime_type: mime.clone(),
                                filename: filename.clone(),
                            });
                            images.len() - 1
                        });

                        let image_id = format!("sha256:{hash}");
                        nodes.push(DiagramNode {
                            id,
                            label: format!("office-image:{image_id}"),
                            kind: DiagramNodeKind::Image { image_id },
                            mermaid_shape: "rect",
                            x: x as f64 * f,
                            y: y as f64 * f,
                            w: w as f64 * f,
                            h: h as f64 * f,
                        });
                        continue;
                    }
                }
            }

            if let Some(text) = extract_text_compact(xml)? {
                let text = text.trim();
                if !text.is_empty() {
                    nodes.push(DiagramNode {
                        id,
                        label: text.to_string(),
                        kind: DiagramNodeKind::Text,
                        mermaid_shape: mermaid_shape_from_prst(prst.as_deref()),
                        x: x as f64 * f,
                        y: y as f64 * f,
                        w: w as f64 * f,
                        h: h as f64 * f,
                    });
                }
            }
        }
    }

    if !lines.is_empty() && !nodes.is_empty() {
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

        for l in &lines {
            let from = nearest(l.x1, l.y1);
            let to = nearest(l.x2, l.y2);
            if let (Some(f), Some(t)) = (from, to) {
                if f == t {
                    continue;
                }
                match l.dir {
                    ArrowDir::None => {
                        // keep as a plain connector
                        edges.push(DiagramEdge {
                            from: Some(f),
                            to: Some(t),
                            dir: ArrowDir::None,
                            label: l.label.clone(),
                        });
                    }
                    ArrowDir::Forward => {
                        edges.push(DiagramEdge {
                            from: Some(f),
                            to: Some(t),
                            dir: ArrowDir::Forward,
                            label: l.label.clone(),
                        });
                    }
                    ArrowDir::Reverse => {
                        edges.push(DiagramEdge {
                            from: Some(t),
                            to: Some(f),
                            dir: ArrowDir::Forward,
                            label: l.label.clone(),
                        });
                    }
                    ArrowDir::Both => {
                        edges.push(DiagramEdge {
                            from: Some(f),
                            to: Some(t),
                            dir: ArrowDir::Both,
                            label: l.label.clone(),
                        });
                    }
                }
            }
        }
    }

    if connectors == 0 || nodes.is_empty() {
        return Ok(None);
    }

    let mut edge_items: Vec<(u32, u32, ArrowDir, String)> = edges
        .iter()
        .filter_map(|e| {
            let from = e.from?;
            let to = e.to?;
            Some((
                from,
                to,
                e.dir,
                e.label.clone().unwrap_or_default().trim().to_string(),
            ))
        })
        .collect();
    edge_items.sort_by(|a, b| (a.0, a.1, &a.3).cmp(&(b.0, b.1, &b.3)));

    if edge_items.is_empty() {
        return Ok(None);
    }

    let node_by_id: HashMap<u32, &DiagramNode> = nodes.iter().map(|n| (n.id, n)).collect();
    let mut node_ids: std::collections::BTreeSet<u32> = Default::default();
    for (from, to, _d, _l) in &edge_items {
        node_ids.insert(*from);
        node_ids.insert(*to);
    }

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

    let mut mermaid = String::new();
    mermaid.push_str("```mermaid\n");
    mermaid.push_str("flowchart LR\n");
    for id in node_ids.iter().copied() {
        if let Some(n) = node_by_id.get(&id).copied() {
            match &n.kind {
                DiagramNodeKind::Text => {
                    mermaid.push_str(&format!(
                        "  n{id}@{{ shape: {}, label: \"{}\" }}\n",
                        n.mermaid_shape,
                        esc(&n.label)
                    ));
                }
                DiagramNodeKind::Image { image_id } => {
                    mermaid.push_str(&format!(
                        "  n{id}@{{ shape: rect, label: \"office-image:{}\" }}\n",
                        image_id
                    ));
                }
            }
        }
    }
    for (from, to, dir, label) in edge_items.into_iter() {
        let op = match dir {
            ArrowDir::None => "---",
            ArrowDir::Forward => "-->",
            ArrowDir::Reverse => "<--",
            ArrowDir::Both => "<-->",
        };
        if label.is_empty() {
            mermaid.push_str(&format!("  n{from} {op} n{to}\n"));
        } else {
            mermaid.push_str(&format!("  n{from} {op}|{}| n{to}\n", esc(&label)));
        }
    }
    mermaid.push_str("```\n");

    let nodes_json = nodes
        .iter()
        .map(|n| {
            let kind = match &n.kind {
                DiagramNodeKind::Text => "text",
                DiagramNodeKind::Image { .. } => "image",
            };
            serde_json::json!({
                "id": format!("n{}", n.id),
                "text": n.label.clone(),
                "kind": kind,
                "bbox": {"x": n.x, "y": n.y, "w": n.w, "h": n.h},
            })
        })
        .collect::<Vec<_>>();
    let edges_json = edges
        .iter()
        .filter_map(|e| {
            Some(serde_json::json!({
                "from": format!("n{}", e.from?),
                "to": format!("n{}", e.to?),
                "kind": "connector",
                "label": e.label,
            }))
        })
        .collect::<Vec<_>>();
    let graph_json = serde_json::json!({
        "nodes": nodes_json,
        "edges": edges_json,
        "warnings": []
    });

    Ok(Some((mermaid, graph_json)))
}

fn mermaid_shape_from_prst(prst: Option<&str>) -> &'static str {
    let p = prst.unwrap_or("").trim();
    match p {
        "" => "rect",
        "rect" | "flowChartProcess" => "rect",
        "roundRect" | "flowChartTerminator" => "rounded",
        "ellipse" => "circle",
        "diamond" | "flowChartDecision" => "diamond",
        "hexagon" => "hex",
        "can" => "cyl",
        "parallelogram" => "lean-r",
        "trapezoid" => "trap-b",
        "triangle" | "rtTriangle" | "ltTriangle" | "upTriangle" | "downTriangle" => "tri",
        "flowChartPredefinedProcess" => "subproc",
        _ => "rect",
    }
}

fn resolve_target(base: &str, target: &str) -> String {
    if target.starts_with('/') {
        return target.trim_start_matches('/').to_string();
    }
    if target.starts_with("http://") || target.starts_with("https://") {
        return target.to_string();
    }

    // For docx, targets in document.xml.rels are usually relative to "word/".
    let mut t = target.to_string();
    while t.starts_with("../") {
        t = t.trim_start_matches("../").to_string();
    }
    format!("{base}{t}")
}

pub fn parse_docx(bytes: &[u8]) -> Result<Vec<Block>> {
    Ok(parse_docx_full(bytes)?.blocks)
}

pub fn parse_docx_full(bytes: &[u8]) -> Result<ParsedOfficeDocument> {
    let doc_xml = read_zip_file(bytes, "word/document.xml")?;
    let rels_xml = read_zip_file_utf8(bytes, "word/_rels/document.xml.rels")
        .unwrap_or_else(|_| "".to_string());
    let styles_xml =
        read_zip_file_utf8(bytes, "word/styles.xml").unwrap_or_else(|_| "".to_string());
    let numbering_xml =
        read_zip_file_utf8(bytes, "word/numbering.xml").unwrap_or_else(|_| "".to_string());

    let rels = parse_relationships(&rels_xml).unwrap_or_default();
    let heading_styles = parse_styles_heading_levels(&styles_xml).unwrap_or_default();
    let numbering = if numbering_xml.trim().is_empty() {
        NumberingIndex::default()
    } else {
        parse_numbering(&numbering_xml).unwrap_or_default()
    };

    let body_elems = extract_body_children_xml(&doc_xml)?;

    #[derive(Clone, Debug)]
    enum Elem {
        Para(ParagraphParsed),
        Tbl(TableParsed),
    }

    let mut elems = Vec::new();
    for (kind, xml) in body_elems {
        if kind == "p" {
            let p = parse_paragraph_xml(&xml, &heading_styles, &numbering)?;
            elems.push(Elem::Para(p));
        } else if kind == "tbl" {
            let t = parse_table_xml(&xml)?;
            elems.push(Elem::Tbl(t));
        }
    }

    let mut images: Vec<ParsedImage> = Vec::new();
    let mut image_by_hash: HashMap<String, usize> = HashMap::new();

    let mut blocks: Vec<Block> = Vec::new();
    let mut next_block_index: usize = 0;

    let mut charts_meta: Vec<serde_json::Value> = Vec::new();
    let mut diagram_graphs_meta: Vec<serde_json::Value> = Vec::new();

    let mut i = 0usize;
    while i < elems.len() {
        match &elems[i] {
            Elem::Para(p) => {
                // List grouping.
                if let Some(ctx) = p.list_ctx {
                    let ordered = ctx.ordered;
                    let num_id = ctx.num_id;

                    let mut items: Vec<ListItem> = Vec::new();
                    let mut j = i;
                    while j < elems.len() {
                        let Elem::Para(pp) = &elems[j] else { break };
                        let Some(c) = pp.list_ctx else { break };
                        if c.num_id != num_id {
                            break;
                        }
                        if !pp.text.trim().is_empty() {
                            items.push(ListItem {
                                level: c.ilvl,
                                text: pp.text.trim().to_string(),
                                source: SourceSpan::default(),
                            });
                        }
                        j += 1;
                    }

                    if !items.is_empty() {
                        blocks.push(Block::List {
                            block_index: next_block_index,
                            ordered,
                            items,
                            source: SourceSpan::default(),
                        });
                        next_block_index += 1;
                    }

                    // Emit links/images from paragraphs in the group deterministically after the list.
                    for k in i..j {
                        let Elem::Para(pp) = &elems[k] else { continue };
                        for (rid, txt) in &pp.links {
                            if let Some((_ty, target)) = rels.get(rid) {
                                blocks.push(Block::Link {
                                    block_index: next_block_index,
                                    url: target.clone(),
                                    text: txt.clone(),
                                    kind: LinkKind::Unknown,
                                    source: SourceSpan::default(),
                                });
                                next_block_index += 1;
                            }
                        }
                        for (rid, alt) in &pp.image_rel_ids {
                            if let Some((_ty, target)) = rels.get(rid) {
                                let entry = resolve_target("word/", target);
                                if let Ok(img_bytes) = read_zip_file(bytes, &entry) {
                                    let hash = sha256_hex(&img_bytes);
                                    let idx =
                                        *image_by_hash.entry(hash.clone()).or_insert_with(|| {
                                            let filename =
                                                entry.split('/').last().map(|s| s.to_string());
                                            let mime = filename
                                                .as_deref()
                                                .map(mime_from_filename)
                                                .unwrap_or_else(|| {
                                                    "application/octet-stream".to_string()
                                                });
                                            images.push(ParsedImage {
                                                id: format!("sha256:{hash}"),
                                                bytes: img_bytes,
                                                mime_type: mime.clone(),
                                                filename: filename.clone(),
                                            });
                                            images.len() - 1
                                        });
                                    let filename = images.get(idx).and_then(|p| p.filename.clone());
                                    let content_type = images.get(idx).map(|p| p.mime_type.clone());
                                    blocks.push(Block::Image {
                                        block_index: next_block_index,
                                        id: format!("sha256:{hash}"),
                                        filename,
                                        content_type,
                                        alt: alt.clone(),
                                        source: SourceSpan::default(),
                                    });
                                    next_block_index += 1;
                                }
                            }
                        }

                        for rid in pp.chart_rel_ids.iter() {
                            let Some((_ty, target)) = rels.get(rid) else {
                                continue;
                            };
                            let entry = resolve_target("word/", target);
                            let Ok(chart_xml) = read_zip_file_utf8(bytes, &entry) else {
                                continue;
                            };
                            let Ok(chart) = parse_chart_xml_cached_values(&chart_xml) else {
                                continue;
                            };
                            let chart_type = chart_type_from_xml(&chart_xml).unwrap_or("unknown");

                            charts_meta.push(serde_json::json!({
                                "chart_type": chart_type,
                                "title": chart.title,
                                "categories": chart.categories,
                                "series": chart.series.iter().map(|s| serde_json::json!({
                                    "name": s.name,
                                    "values": s.values,
                                    "unit": unit_from_numfmt(s.format_code.as_deref()),
                                })).collect::<Vec<_>>()
                            }));

                            blocks.push(Block::Heading {
                                block_index: next_block_index,
                                level: 3,
                                text: chart.title.clone().unwrap_or_else(|| "Chart".to_string()),
                                source: SourceSpan::default(),
                            });
                            next_block_index += 1;

                            let chart_shape = if chart_type == "unknown" {
                                "chart".to_string()
                            } else {
                                format!("{chart_type} chart")
                            };
                            let mut units: Vec<(String, String)> = Vec::new();
                            for (idx, s) in chart.series.iter().enumerate() {
                                if let Some(u) = unit_from_numfmt(s.format_code.as_deref()) {
                                    let name = s
                                        .name
                                        .clone()
                                        .filter(|t| !t.trim().is_empty())
                                        .unwrap_or_else(|| format!("series{}", idx + 1));
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
                                block_index: next_block_index,
                                text: note,
                                source: SourceSpan::default(),
                            });
                            next_block_index += 1;

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
                                    block_index: next_block_index,
                                    rows,
                                    source: SourceSpan::default(),
                                });
                                next_block_index += 1;
                            }
                        }

                        for drawing in pp.drawings.iter() {
                            if let Some((mermaid, graph_json)) = diagram_from_drawing_xml(
                                drawing,
                                bytes,
                                &rels,
                                &mut images,
                                &mut image_by_hash,
                            )? {
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
                    }

                    i = j;
                    continue;
                }

                // Heading vs paragraph.
                if let Some(level) = p.heading_level {
                    if !p.text.trim().is_empty() {
                        blocks.push(Block::Heading {
                            block_index: next_block_index,
                            level,
                            text: p.text.clone(),
                            source: SourceSpan::default(),
                        });
                        next_block_index += 1;
                    }
                } else if !p.text.trim().is_empty() {
                    blocks.push(Block::Paragraph {
                        block_index: next_block_index,
                        text: p.text.clone(),
                        source: SourceSpan::default(),
                    });
                    next_block_index += 1;
                }

                // Links/images in the paragraph.
                for (rid, txt) in &p.links {
                    if let Some((_ty, target)) = rels.get(rid) {
                        blocks.push(Block::Link {
                            block_index: next_block_index,
                            url: target.clone(),
                            text: txt.clone(),
                            kind: LinkKind::Unknown,
                            source: SourceSpan::default(),
                        });
                        next_block_index += 1;
                    }
                }

                for (rid, alt) in &p.image_rel_ids {
                    let Some((_ty, target)) = rels.get(rid) else {
                        continue;
                    };
                    let entry = resolve_target("word/", target);
                    let img_bytes = read_zip_file(bytes, &entry).unwrap_or_default();
                    if img_bytes.is_empty() {
                        continue;
                    }
                    let hash = sha256_hex(&img_bytes);
                    let idx = *image_by_hash.entry(hash.clone()).or_insert_with(|| {
                        let filename = entry.split('/').last().map(|s| s.to_string());
                        let mime = filename
                            .as_deref()
                            .map(mime_from_filename)
                            .unwrap_or_else(|| "application/octet-stream".to_string());
                        images.push(ParsedImage {
                            id: format!("sha256:{hash}"),
                            bytes: img_bytes,
                            mime_type: mime.clone(),
                            filename: filename.clone(),
                        });
                        images.len() - 1
                    });
                    let filename = images.get(idx).and_then(|p| p.filename.clone());
                    let content_type = images.get(idx).map(|p| p.mime_type.clone());
                    blocks.push(Block::Image {
                        block_index: next_block_index,
                        id: format!("sha256:{hash}"),
                        filename,
                        content_type,
                        alt: alt.clone(),
                        source: SourceSpan::default(),
                    });
                    next_block_index += 1;
                }

                for rid in p.chart_rel_ids.iter() {
                    let Some((_ty, target)) = rels.get(rid) else {
                        continue;
                    };
                    let entry = resolve_target("word/", target);
                    let Ok(chart_xml) = read_zip_file_utf8(bytes, &entry) else {
                        continue;
                    };
                    let Ok(chart) = parse_chart_xml_cached_values(&chart_xml) else {
                        continue;
                    };
                    let chart_type = chart_type_from_xml(&chart_xml).unwrap_or("unknown");

                    charts_meta.push(serde_json::json!({
                        "chart_type": chart_type,
                        "title": chart.title,
                        "categories": chart.categories,
                        "series": chart.series.iter().map(|s| serde_json::json!({
                            "name": s.name,
                            "values": s.values,
                            "unit": unit_from_numfmt(s.format_code.as_deref()),
                        })).collect::<Vec<_>>()
                    }));

                    blocks.push(Block::Heading {
                        block_index: next_block_index,
                        level: 3,
                        text: chart.title.clone().unwrap_or_else(|| "Chart".to_string()),
                        source: SourceSpan::default(),
                    });
                    next_block_index += 1;

                    let chart_shape = if chart_type == "unknown" {
                        "chart".to_string()
                    } else {
                        format!("{chart_type} chart")
                    };
                    let mut units: Vec<(String, String)> = Vec::new();
                    for (idx, s) in chart.series.iter().enumerate() {
                        if let Some(u) = unit_from_numfmt(s.format_code.as_deref()) {
                            let name = s
                                .name
                                .clone()
                                .filter(|t| !t.trim().is_empty())
                                .unwrap_or_else(|| format!("series{}", idx + 1));
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
                            note.push_str(&format!("; units: {}", uniq.iter().next().unwrap()));
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
                        block_index: next_block_index,
                        text: note,
                        source: SourceSpan::default(),
                    });
                    next_block_index += 1;

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
                                text: s.name.clone().unwrap_or_else(|| "series".to_string()),
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
                            block_index: next_block_index,
                            rows,
                            source: SourceSpan::default(),
                        });
                        next_block_index += 1;
                    }
                }

                for drawing in p.drawings.iter() {
                    if let Some((mermaid, graph_json)) = diagram_from_drawing_xml(
                        drawing,
                        bytes,
                        &rels,
                        &mut images,
                        &mut image_by_hash,
                    )? {
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

                i += 1;
            }
            Elem::Tbl(t) => {
                blocks.push(Block::Table {
                    block_index: next_block_index,
                    rows: t.rows.clone(),
                    source: SourceSpan::default(),
                });
                next_block_index += 1;
                i += 1;
            }
        }
    }

    // Ensure monotonic block indices.
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
            "kind": "docx",
            "title": extract_docx_title(bytes),
            "charts": charts_meta,
            "diagram_graphs": diagram_graphs_meta,
        }),
    })
}

pub fn parse(bytes: &[u8]) -> crate::Result<crate::Document> {
    let parsed =
        parse_docx_full(bytes).map_err(|e| crate::Error::from_parse(crate::Format::Docx, e))?;
    Ok(super::finalize(crate::Format::Docx, parsed))
}
