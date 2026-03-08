use anyhow::{Context, Result, anyhow};
use quick_xml::Reader;
use quick_xml::events::Event;

use crate::document_ast::{Block, Cell, ListItem, SourceSpan};

use super::{
    ParsedImage, ParsedOfficeDocument, mime_from_filename, read_zip_file, read_zip_file_utf8,
    sha256_hex,
};

fn attr_val_exact(e: &quick_xml::events::BytesStart<'_>, key: &[u8]) -> Option<String> {
    for a in e.attributes().flatten() {
        if a.key.as_ref() == key {
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

fn mermaid_shape_from_odf_type(t: Option<&str>) -> &'static str {
    let t = t.unwrap_or("").trim();
    match t {
        "" => "rect",
        // ODF enhanced-geometry types (best effort)
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

fn parse_odp_full_inner(
    bytes: &[u8],
    include_slide_snapshots: bool,
) -> Result<ParsedOfficeDocument> {
    let content_xml = read_zip_file_utf8(bytes, "content.xml").context("read content.xml")?;
    let _ = read_zip_file_utf8(bytes, "styles.xml").ok();

    let mut blocks: Vec<Block> = Vec::new();
    let mut images: Vec<ParsedImage> = Vec::new();
    let mut slides_meta: Vec<serde_json::Value> = Vec::new();
    let mut block_index: usize = 0;

    let mut reader = Reader::from_str(&content_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut in_slide = false;
    let mut slide_depth = 0usize;
    let mut slide_idx_1: usize = 0;
    let mut slide_name: Option<String> = None;

    let mut slide_has_diagram = false;

    let mut in_chart_depth: usize = 0;
    let mut chart_class_stack: Vec<Option<String>> = Vec::new();

    #[derive(Clone, Debug)]
    enum NodeKind {
        Text,
        Image { _image_id: String },
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

    #[derive(Clone, Debug)]
    struct PendingConn {
        start_id: Option<String>,
        end_id: Option<String>,
        x1: Option<f64>,
        y1: Option<f64>,
        x2: Option<f64>,
        y2: Option<f64>,
        dir: ArrowDir,
        label: Option<String>,
    }

    let mut text_boxes: usize = 0;
    let mut frame_bbox_stack: Vec<Option<(f64, f64, f64, f64)>> = Vec::new();
    let mut frame_title_stack: Vec<bool> = Vec::new();
    let mut frame_text_stack: Vec<String> = Vec::new();
    let mut frame_id_stack: Vec<Option<String>> = Vec::new();
    let mut frame_shape_type_stack: Vec<Option<String>> = Vec::new();

    let mut next_node_id: u32 = 1;
    let mut node_by_odf_id: std::collections::HashMap<String, u32> = Default::default();
    let mut node_index_by_id: std::collections::HashMap<u32, usize> = Default::default();
    let mut nodes: Vec<Node> = Vec::new();

    let mut connectors: usize = 0;
    let mut pending_connectors: Vec<PendingConn> = Vec::new();
    let mut cur_connector: Option<PendingConn> = None;
    let mut in_connector = false;
    let mut connector_text = String::new();
    let mut edges: Vec<Edge> = Vec::new();
    let mut diagram_graphs_meta: Vec<serde_json::Value> = Vec::new();

    let mut text_chars: usize = 0;
    let mut embedded_images: usize = 0;
    let mut tables: usize = 0;
    let mut charts: usize = 0;
    let mut smartart: usize = 0;
    let mut excluded_visuals: Vec<String> = Vec::new();
    let mut snapshot_image_id: Option<String> = None;

    let mut title_frame = false;
    let mut title_frame_depth: usize = 0;
    let mut title_text = String::new();

    let mut heading_level: Option<u8> = None;

    let mut cur_para_text = String::new();
    let mut in_text_p = false;
    let mut in_text_span = false;

    let mut list_depth: u8 = 0;
    let mut in_list_item = false;
    let mut cur_list: Option<(bool, Vec<ListItem>)> = None;

    let mut in_table = false;
    let mut cur_row: Vec<Cell> = Vec::new();
    let mut in_table_cell = false;
    let mut cur_cell_text = String::new();
    let mut cur_cell_colspan: usize = 1;
    let mut cur_cell_rowspan: usize = 1;
    let mut cur_cell_repeat: usize = 1;
    let mut table_rows: Vec<Vec<Cell>> = Vec::new();

    let mut pending_slide_block_first: Option<usize> = None;

    let elem_id = |e: &quick_xml::events::BytesStart<'_>| -> Option<String> {
        attr_val_exact(e, b"draw:id")
            .or_else(|| attr_val_exact(e, b"xml:id"))
            .or_else(|| attr_val_exact(e, b"id"))
    };

    let node_id_for_odf = |odf_id: &str,
                           next_node_id: &mut u32,
                           node_by_odf_id: &mut std::collections::HashMap<String, u32>|
     -> u32 {
        if let Some(&id) = node_by_odf_id.get(odf_id) {
            return id;
        }
        let id = *next_node_id;
        *next_node_id = (*next_node_id).saturating_add(1).max(id + 1);
        node_by_odf_id.insert(odf_id.to_string(), id);
        id
    };

    let marker_present = |v: Option<String>| -> bool {
        let Some(v) = v else { return false };
        let t = v.trim();
        !(t.is_empty() || t.eq_ignore_ascii_case("none"))
    };

    let arrow_dir_from_elem = |e: &quick_xml::events::BytesStart<'_>| -> ArrowDir {
        let has_start = marker_present(attr_val_exact(e, b"draw:marker-start"))
            || marker_present(attr_val_exact(e, b"svg:marker-start"));
        let has_end = marker_present(attr_val_exact(e, b"draw:marker-end"))
            || marker_present(attr_val_exact(e, b"svg:marker-end"));
        match (has_start, has_end) {
            (false, false) => ArrowDir::None,
            (false, true) => ArrowDir::Forward,
            (true, false) => ArrowDir::Reverse,
            (true, true) => ArrowDir::Both,
        }
    };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                match e.name().as_ref() {
                    b"draw:page" => {
                        in_slide = true;
                        slide_depth = 1;
                        slide_idx_1 += 1;
                        slide_name = attr_val_exact(&e, b"draw:name");
                        title_text.clear();
                        title_frame = false;
                        title_frame_depth = 0;
                        slide_has_diagram = false;
                        pending_slide_block_first = Some(block_index);

                        frame_bbox_stack.clear();
                        frame_title_stack.clear();
                        frame_text_stack.clear();
                        frame_id_stack.clear();
                        frame_shape_type_stack.clear();
                        text_chars = 0;
                        embedded_images = 0;
                        tables = 0;
                        charts = 0;
                        smartart = 0;
                        excluded_visuals.clear();
                        snapshot_image_id = None;
                        text_boxes = 0;

                        next_node_id = 1;
                        node_by_odf_id.clear();
                        node_index_by_id.clear();
                        nodes.clear();
                        connectors = 0;
                        pending_connectors.clear();
                        cur_connector = None;
                        in_connector = false;
                        connector_text.clear();
                        edges.clear();
                    }
                    _ if in_slide => {
                        slide_depth += 1;

                        if e.name().as_ref() == b"chart:chart" {
                            charts += 1;
                            slide_has_diagram = true;
                            if !excluded_visuals.iter().any(|s| s == "chart") {
                                excluded_visuals.push("chart".to_string());
                            }
                            in_chart_depth = in_chart_depth.saturating_add(1);
                            let cls = attr_val_exact(&e, b"chart:class")
                                .or_else(|| attr_val_exact(&e, b"class"))
                                .map(|v| {
                                    let v = v.trim().to_string();
                                    v.rsplit(':').next().unwrap_or(v.as_str()).to_string()
                                });
                            chart_class_stack.push(cls);
                        }

                        // Best-effort diagram detection (charts, OLE objects, custom shapes).
                        // This is used by higher layers to decide whether slide rasterization is
                        // worth doing for vision pipelines.
                        match e.name().as_ref() {
                            b"chart:chart" => {}
                            b"draw:object" | b"draw:object-ole" => {
                                slide_has_diagram = true;
                            }
                            b"draw:connector" | b"draw:line" => {
                                slide_has_diagram = true;
                            }
                            _ => {}
                        }

                        if e.name().as_ref() == b"draw:connector"
                            || e.name().as_ref() == b"draw:line"
                        {
                            connectors = connectors.saturating_add(1);
                            in_connector = true;
                            connector_text.clear();
                            let start_id = attr_val_exact(&e, b"draw:start-shape");
                            let end_id = attr_val_exact(&e, b"draw:end-shape");
                            let x1 = attr_val_exact(&e, b"svg:x1").and_then(|v| odf_len_to_px(&v));
                            let y1 = attr_val_exact(&e, b"svg:y1").and_then(|v| odf_len_to_px(&v));
                            let x2 = attr_val_exact(&e, b"svg:x2").and_then(|v| odf_len_to_px(&v));
                            let y2 = attr_val_exact(&e, b"svg:y2").and_then(|v| odf_len_to_px(&v));
                            cur_connector = Some(PendingConn {
                                start_id,
                                end_id,
                                x1,
                                y1,
                                x2,
                                y2,
                                dir: arrow_dir_from_elem(&e),
                                label: None,
                            });
                        }

                        if e.name().as_ref() == b"draw:enhanced-geometry" {
                            if let Some(t) = attr_val_exact(&e, b"draw:type") {
                                if let Some(last) = frame_shape_type_stack.last_mut() {
                                    if last.is_none() {
                                        *last = Some(t);
                                    }
                                }
                            }
                        }

                        if e.name().as_ref() == b"draw:frame"
                            || e.name().as_ref() == b"draw:custom-shape"
                        {
                            let is_title = attr_val_exact(&e, b"presentation:class")
                                .map(|cls| {
                                    let c = cls.to_ascii_lowercase();
                                    c == "title" || c == "subtitle"
                                })
                                .unwrap_or(false);
                            if e.name().as_ref() == b"draw:frame" {
                                if is_title {
                                    title_frame_depth = title_frame_depth.saturating_add(1);
                                }
                                title_frame = title_frame_depth > 0;
                            }

                            if e.name().as_ref() == b"draw:custom-shape" {
                                slide_has_diagram = true;
                            }

                            let x = attr_val_exact(&e, b"svg:x").and_then(|v| odf_len_to_px(&v));
                            let y = attr_val_exact(&e, b"svg:y").and_then(|v| odf_len_to_px(&v));
                            let w =
                                attr_val_exact(&e, b"svg:width").and_then(|v| odf_len_to_px(&v));
                            let h =
                                attr_val_exact(&e, b"svg:height").and_then(|v| odf_len_to_px(&v));
                            let bbox = match (x, y, w, h) {
                                (Some(x), Some(y), Some(w), Some(h)) => Some((x, y, w, h)),
                                _ => None,
                            };

                            frame_bbox_stack.push(bbox);
                            frame_title_stack.push(is_title);
                            frame_text_stack.push(String::new());
                            frame_id_stack.push(elem_id(&e));
                            frame_shape_type_stack.push(None);
                        }

                        if e.name().as_ref() == b"table:table" {
                            in_table = true;
                            table_rows.clear();
                        }
                        if e.name().as_ref() == b"table:table-row" && in_table {
                            cur_row = Vec::new();
                        }
                        if e.name().as_ref() == b"table:table-cell" && in_table {
                            in_table_cell = true;
                            cur_cell_text.clear();
                            cur_cell_colspan = attr_val_exact(&e, b"table:number-columns-spanned")
                                .and_then(|v| v.parse::<usize>().ok())
                                .unwrap_or(1);
                            cur_cell_rowspan = attr_val_exact(&e, b"table:number-rows-spanned")
                                .and_then(|v| v.parse::<usize>().ok())
                                .unwrap_or(1);
                            cur_cell_repeat = attr_val_exact(&e, b"table:number-columns-repeated")
                                .and_then(|v| v.parse::<usize>().ok())
                                .unwrap_or(1)
                                .max(1);
                        }

                        if e.name().as_ref() == b"text:list" {
                            if !in_table {
                                list_depth = list_depth.saturating_add(1);
                            }
                        }
                        if e.name().as_ref() == b"text:list-item" {
                            if !in_table {
                                in_list_item = true;
                            }
                        }

                        if e.name().as_ref() == b"text:h" {
                            // ODF headings use text:outline-level (1..). We shift by +1 to keep
                            // slide title at level=1.
                            let lvl = attr_val_exact(&e, b"text:outline-level")
                                .or_else(|| attr_val_exact(&e, b"outline-level"))
                                .and_then(|v| v.parse::<u8>().ok())
                                .unwrap_or(1);
                            heading_level = Some(lvl.saturating_add(1).max(2).min(6));
                            in_text_p = true;
                            cur_para_text.clear();
                        }

                        if e.name().as_ref() == b"text:p" {
                            in_text_p = true;
                            cur_para_text.clear();
                        }
                        if e.name().as_ref() == b"text:span" {
                            in_text_span = true;
                        }
                        if e.name().as_ref() == b"draw:image" {
                            if let Some(href) = attr_val_exact(&e, b"xlink:href") {
                                let href = href.trim().trim_start_matches("./");
                                if !href.is_empty() {
                                    if let Ok(img_bytes) = read_zip_file(bytes, href) {
                                        embedded_images += 1;
                                        let filename = Some(
                                            href.rsplit('/').next().unwrap_or(href).to_string(),
                                        );
                                        let mime_type = filename
                                            .as_deref()
                                            .map(mime_from_filename)
                                            .unwrap_or_else(|| {
                                                "application/octet-stream".to_string()
                                            });
                                        let hash = sha256_hex(&img_bytes);
                                        let id = format!("sha256:{hash}");

                                        images.push(ParsedImage {
                                            id: id.clone(),
                                            bytes: img_bytes,
                                            mime_type: mime_type.clone(),
                                            filename: filename.clone(),
                                        });
                                        flush_list_block(
                                            &mut blocks,
                                            &mut block_index,
                                            &mut cur_list,
                                        );
                                        blocks.push(Block::Image {
                                            block_index,
                                            id: id.clone(),
                                            filename,
                                            content_type: Some(mime_type),
                                            alt: None,
                                            source: SourceSpan::default(),
                                        });
                                        block_index += 1;

                                        if let Some((x, y, w, h)) =
                                            frame_bbox_stack.last().copied().unwrap_or(None)
                                        {
                                            let node_id = if let Some(odf_id) =
                                                frame_id_stack.last().cloned().unwrap_or(None)
                                            {
                                                node_id_for_odf(
                                                    &odf_id,
                                                    &mut next_node_id,
                                                    &mut node_by_odf_id,
                                                )
                                            } else {
                                                let nid = next_node_id;
                                                next_node_id = next_node_id.saturating_add(1);
                                                nid
                                            };

                                            let n = Node {
                                                id: node_id,
                                                label: format!("office-image:{id}"),
                                                kind: NodeKind::Image {
                                                    _image_id: id.clone(),
                                                },
                                                mermaid_shape: "rect",
                                                x,
                                                y,
                                                w,
                                                h,
                                            };
                                            if let Some(&idx) = node_index_by_id.get(&node_id) {
                                                nodes[idx] = n;
                                            } else {
                                                node_index_by_id.insert(node_id, nodes.len());
                                                nodes.push(n);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                if in_slide {
                    match e.name().as_ref() {
                        b"chart:chart" => {
                            charts += 1;
                            slide_has_diagram = true;
                            if !excluded_visuals.iter().any(|s| s == "chart") {
                                excluded_visuals.push("chart".to_string());
                            }

                            let cls = attr_val_exact(&e, b"chart:class")
                                .or_else(|| attr_val_exact(&e, b"class"))
                                .map(|v| {
                                    let v = v.trim().to_string();
                                    v.rsplit(':').next().unwrap_or(v.as_str()).to_string()
                                });
                            chart_class_stack.push(cls);
                            chart_class_stack.pop();
                        }
                        b"draw:object" | b"draw:object-ole" | b"draw:custom-shape" => {
                            slide_has_diagram = true;
                        }
                        b"draw:connector" | b"draw:line" => {
                            slide_has_diagram = true;
                            connectors = connectors.saturating_add(1);
                            let start_id = attr_val_exact(&e, b"draw:start-shape");
                            let end_id = attr_val_exact(&e, b"draw:end-shape");
                            let x1 = attr_val_exact(&e, b"svg:x1").and_then(|v| odf_len_to_px(&v));
                            let y1 = attr_val_exact(&e, b"svg:y1").and_then(|v| odf_len_to_px(&v));
                            let x2 = attr_val_exact(&e, b"svg:x2").and_then(|v| odf_len_to_px(&v));
                            let y2 = attr_val_exact(&e, b"svg:y2").and_then(|v| odf_len_to_px(&v));
                            pending_connectors.push(PendingConn {
                                start_id,
                                end_id,
                                x1,
                                y1,
                                x2,
                                y2,
                                dir: arrow_dir_from_elem(&e),
                                label: None,
                            });
                        }
                        _ => {}
                    }

                    if e.name().as_ref() == b"draw:image" {
                        if let Some(href) = attr_val_exact(&e, b"xlink:href") {
                            let href = href.trim().trim_start_matches("./");
                            if !href.is_empty() {
                                if let Ok(img_bytes) = read_zip_file(bytes, href) {
                                    embedded_images += 1;
                                    let filename =
                                        Some(href.rsplit('/').next().unwrap_or(href).to_string());
                                    let mime_type = filename
                                        .as_deref()
                                        .map(mime_from_filename)
                                        .unwrap_or_else(|| "application/octet-stream".to_string());
                                    let hash = sha256_hex(&img_bytes);
                                    let id = format!("sha256:{hash}");

                                    images.push(ParsedImage {
                                        id: id.clone(),
                                        bytes: img_bytes,
                                        mime_type: mime_type.clone(),
                                        filename: filename.clone(),
                                    });
                                    flush_list_block(&mut blocks, &mut block_index, &mut cur_list);
                                    blocks.push(Block::Image {
                                        block_index,
                                        id: id.clone(),
                                        filename,
                                        content_type: Some(mime_type),
                                        alt: None,
                                        source: SourceSpan::default(),
                                    });
                                    block_index += 1;

                                    if let Some((x, y, w, h)) =
                                        frame_bbox_stack.last().copied().unwrap_or(None)
                                    {
                                        let node_id = if let Some(odf_id) =
                                            frame_id_stack.last().cloned().unwrap_or(None)
                                        {
                                            node_id_for_odf(
                                                &odf_id,
                                                &mut next_node_id,
                                                &mut node_by_odf_id,
                                            )
                                        } else {
                                            let nid = next_node_id;
                                            next_node_id = next_node_id.saturating_add(1);
                                            nid
                                        };

                                        let n = Node {
                                            id: node_id,
                                            label: format!("office-image:{id}"),
                                            kind: NodeKind::Image {
                                                _image_id: id.clone(),
                                            },
                                            mermaid_shape: "rect",
                                            x,
                                            y,
                                            w,
                                            h,
                                        };
                                        if let Some(&idx) = node_index_by_id.get(&node_id) {
                                            nodes[idx] = n;
                                        } else {
                                            node_index_by_id.insert(node_id, nodes.len());
                                            nodes.push(n);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            Ok(Event::Text(e)) => {
                if in_slide {
                    if let Ok(t) = String::from_utf8(e.to_vec()) {
                        let trimmed = t.trim();
                        if in_text_p || in_text_span {
                            if !trimmed.is_empty() {
                                if !cur_para_text.is_empty() {
                                    cur_para_text.push(' ');
                                }
                                cur_para_text.push_str(trimmed);

                                if in_connector {
                                    if !connector_text.is_empty() {
                                        connector_text.push(' ');
                                    }
                                    connector_text.push_str(trimmed);
                                }

                                if let Some(t) = frame_text_stack.last_mut() {
                                    if !t.is_empty() {
                                        t.push(' ');
                                    }
                                    t.push_str(trimmed);
                                }
                            }
                        }

                        if in_table_cell {
                            if !trimmed.is_empty() {
                                if !cur_cell_text.is_empty() {
                                    cur_cell_text.push(' ');
                                }
                                cur_cell_text.push_str(trimmed);
                            }
                        }

                        if title_frame {
                            if !trimmed.is_empty() {
                                if !title_text.is_empty() {
                                    title_text.push(' ');
                                }
                                title_text.push_str(trimmed);
                            }
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                if in_slide {
                    if e.name().as_ref() == b"text:span" {
                        in_text_span = false;
                    }
                    if e.name().as_ref() == b"text:p" {
                        in_text_p = false;
                        let txt = cur_para_text.trim().to_string();
                        if !txt.is_empty() {
                            if in_table {
                                // Table cell content is handled via cur_cell_text aggregation.
                            } else if in_connector {
                                // Connector label text is captured separately.
                            } else if title_frame {
                                // Do not emit as body paragraph.
                            } else {
                                text_chars = text_chars.saturating_add(txt.chars().count());
                                let is_list = list_depth > 0 && in_list_item;
                                push_paragraph_or_list_item(
                                    &mut blocks,
                                    &mut block_index,
                                    &mut cur_list,
                                    is_list,
                                    false,
                                    list_depth.saturating_sub(1),
                                    txt,
                                );
                            }
                        }
                        cur_para_text.clear();
                    }

                    if e.name().as_ref() == b"text:h" {
                        in_text_p = false;
                        let txt = cur_para_text.trim().to_string();
                        if !txt.is_empty() && !in_table && !in_connector {
                            text_chars = text_chars.saturating_add(txt.chars().count());
                            flush_list_block(&mut blocks, &mut block_index, &mut cur_list);
                            blocks.push(Block::Heading {
                                block_index,
                                level: heading_level.unwrap_or(2),
                                text: txt,
                                source: SourceSpan::default(),
                            });
                            block_index += 1;
                        }
                        heading_level = None;
                        cur_para_text.clear();
                    }
                    if e.name().as_ref() == b"text:list-item" {
                        if !in_table {
                            in_list_item = false;
                        }
                    }
                    if e.name().as_ref() == b"text:list" {
                        if !in_table {
                            list_depth = list_depth.saturating_sub(1);
                        }
                    }
                    if e.name().as_ref() == b"draw:frame"
                        || e.name().as_ref() == b"draw:custom-shape"
                    {
                        let text = frame_text_stack.pop().unwrap_or_default();
                        let bbox = frame_bbox_stack.pop().unwrap_or(None);
                        let is_title = frame_title_stack.pop().unwrap_or(false);
                        let odf_id = frame_id_stack.pop().unwrap_or(None);
                        let shape_type = frame_shape_type_stack.pop().unwrap_or(None);

                        if e.name().as_ref() == b"draw:frame" {
                            if is_title {
                                title_frame_depth = title_frame_depth.saturating_sub(1);
                            }
                            title_frame = title_frame_depth > 0;
                        }

                        let t = text.trim();
                        if !is_title {
                            if let Some((x, y, w, h)) = bbox {
                                if !t.is_empty() {
                                    text_boxes = text_boxes.saturating_add(1);
                                    let node_id = if let Some(odf_id) = odf_id.as_deref() {
                                        node_id_for_odf(
                                            odf_id,
                                            &mut next_node_id,
                                            &mut node_by_odf_id,
                                        )
                                    } else {
                                        let nid = next_node_id;
                                        next_node_id = next_node_id.saturating_add(1);
                                        nid
                                    };

                                    let n = Node {
                                        id: node_id,
                                        label: t.to_string(),
                                        kind: NodeKind::Text,
                                        mermaid_shape: mermaid_shape_from_odf_type(
                                            shape_type.as_deref(),
                                        ),
                                        x,
                                        y,
                                        w,
                                        h,
                                    };
                                    if let Some(&idx) = node_index_by_id.get(&node_id) {
                                        if matches!(nodes[idx].kind, NodeKind::Text) {
                                            nodes[idx] = n;
                                        }
                                    } else {
                                        node_index_by_id.insert(node_id, nodes.len());
                                        nodes.push(n);
                                    }
                                }
                            }
                        }
                    }

                    if e.name().as_ref() == b"draw:connector" || e.name().as_ref() == b"draw:line" {
                        in_connector = false;
                        if let Some(mut c) = cur_connector.take() {
                            let label = connector_text.trim();
                            if !label.is_empty() {
                                c.label = Some(label.to_string());
                            }
                            pending_connectors.push(c);
                        }
                        connector_text.clear();
                    }

                    if e.name().as_ref() == b"table:table-cell" && in_table {
                        in_table_cell = false;
                        let txt = cur_cell_text.trim().to_string();
                        let cell = Cell {
                            text: txt,
                            colspan: cur_cell_colspan,
                            rowspan: cur_cell_rowspan,
                        };
                        for _ in 0..cur_cell_repeat {
                            cur_row.push(cell.clone());
                        }
                        cur_cell_text.clear();
                        cur_cell_repeat = 1;
                        cur_cell_colspan = 1;
                        cur_cell_rowspan = 1;
                    }
                    if e.name().as_ref() == b"table:table-row" && in_table {
                        if !cur_row.is_empty() {
                            table_rows.push(std::mem::take(&mut cur_row));
                        }
                    }
                    if e.name().as_ref() == b"table:table" && in_table {
                        // End of table: emit a Block::Table.
                        in_table = false;
                        tables += 1;
                        flush_list_block(&mut blocks, &mut block_index, &mut cur_list);
                        if !table_rows.is_empty() {
                            if in_chart_depth > 0 {
                                let chart_type = chart_class_stack
                                    .last()
                                    .cloned()
                                    .flatten()
                                    .unwrap_or_else(|| "chart".to_string());
                                let chart_shape = format!("{chart_type} chart");
                                let mut note = format!("Note: {chart_shape}");
                                if let Some(u) = unit_from_table_rows(&table_rows) {
                                    note.push_str(&format!("; units: {u}"));
                                }
                                blocks.push(Block::Paragraph {
                                    block_index,
                                    text: note,
                                    source: SourceSpan::default(),
                                });
                                block_index += 1;
                            }
                            blocks.push(Block::Table {
                                block_index,
                                rows: std::mem::take(&mut table_rows),
                                source: SourceSpan::default(),
                            });
                            block_index += 1;
                        }
                    }

                    if e.name().as_ref() == b"chart:chart" {
                        in_chart_depth = in_chart_depth.saturating_sub(1);
                        let _ = chart_class_stack.pop();
                    }

                    slide_depth = slide_depth.saturating_sub(1);
                    if e.name().as_ref() == b"draw:page" || slide_depth == 0 {
                        in_slide = false;
                        flush_list_block(&mut blocks, &mut block_index, &mut cur_list);

                        // Close any unterminated connector.
                        if in_connector {
                            in_connector = false;
                            if let Some(mut c) = cur_connector.take() {
                                let label = connector_text.trim();
                                if !label.is_empty() {
                                    c.label = Some(label.to_string());
                                }
                                pending_connectors.push(c);
                            }
                            connector_text.clear();
                        }

                        // Convert connectors/lines into a diagram graph and Mermaid block.
                        let mut diagram_warnings: Vec<String> = Vec::new();
                        if !pending_connectors.is_empty() && !nodes.is_empty() {
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

                            for c in pending_connectors.drain(..) {
                                let mut from = c
                                    .start_id
                                    .as_deref()
                                    .and_then(|s| node_by_odf_id.get(s).copied());
                                let mut to = c
                                    .end_id
                                    .as_deref()
                                    .and_then(|s| node_by_odf_id.get(s).copied());

                                if (from.is_none() || to.is_none())
                                    && c.x1.is_some()
                                    && c.y1.is_some()
                                    && c.x2.is_some()
                                    && c.y2.is_some()
                                {
                                    let f = nearest(c.x1.unwrap(), c.y1.unwrap());
                                    let t = nearest(c.x2.unwrap(), c.y2.unwrap());
                                    if from.is_none() {
                                        from = f;
                                    }
                                    if to.is_none() {
                                        to = t;
                                    }
                                }

                                let (Some(from), Some(to)) = (from, to) else {
                                    diagram_warnings.push("connector missing endpoint".to_string());
                                    continue;
                                };
                                if from == to {
                                    continue;
                                }

                                edges.push(Edge {
                                    from,
                                    to,
                                    dir: c.dir,
                                    label: c.label.clone(),
                                });
                            }

                            if !edges.is_empty() {
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

                                let node_by_id: std::collections::HashMap<u32, &Node> =
                                    nodes.iter().map(|n| (n.id, n)).collect();
                                let mut node_ids: std::collections::BTreeSet<u32> =
                                    Default::default();
                                for e in &edges {
                                    node_ids.insert(e.from);
                                    node_ids.insert(e.to);
                                }

                                let mut mermaid = String::new();
                                mermaid.push_str("```mermaid\n");
                                mermaid.push_str("flowchart LR\n");
                                for id in node_ids.iter().copied() {
                                    if let Some(n) = node_by_id.get(&id).copied() {
                                        let label = esc(&n.label);
                                        mermaid.push_str(&format!(
                                            "  n{id}@{{ shape: {}, label: \"{}\" }}\n",
                                            n.mermaid_shape, label
                                        ));
                                    } else {
                                        mermaid.push_str(&format!(
                                            "  n{id}@{{ shape: rect, label: \"n{id}\" }}\n"
                                        ));
                                    }
                                }

                                let mut edge_items = edges
                                    .iter()
                                    .map(|e| {
                                        (
                                            e.from,
                                            e.to,
                                            e.dir.clone(),
                                            e.label.clone().unwrap_or_default(),
                                        )
                                    })
                                    .collect::<Vec<_>>();
                                edge_items.sort_by(|a, b| (a.0, a.1, &a.3).cmp(&(b.0, b.1, &b.3)));

                                for (from, to, dir, label) in edge_items.into_iter() {
                                    let op = match dir {
                                        ArrowDir::None => "---",
                                        ArrowDir::Forward => "-->",
                                        ArrowDir::Reverse => "<--",
                                        ArrowDir::Both => "<-->",
                                    };
                                    if label.trim().is_empty() {
                                        mermaid.push_str(&format!("  n{from} {op} n{to}\n"));
                                    } else {
                                        let l = esc(&label);
                                        mermaid.push_str(&format!("  n{from} {op}|{l}| n{to}\n"));
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

                        let block_first = pending_slide_block_first.take().unwrap_or(block_index);
                        let raw_title = title_text.trim();
                        let slide_title = if !raw_title.is_empty() {
                            raw_title.to_string()
                        } else {
                            slide_name
                                .clone()
                                .unwrap_or_else(|| format!("Slide {slide_idx_1}"))
                        };
                        let heading_text = slide_title.clone();

                        let graph_json = if connectors > 0 && !nodes.is_empty() && !edges.is_empty()
                        {
                            let mut used: std::collections::HashSet<u32> = Default::default();
                            for e in &edges {
                                used.insert(e.from);
                                used.insert(e.to);
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

                            let edges_json = edges
                                .iter()
                                .map(|e| {
                                    serde_json::json!({
                                        "from": format!("n{}", e.from),
                                        "to": format!("n{}", e.to),
                                        "kind": "connector",
                                        "label": e.label
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
                                "warnings": diagram_warnings
                            });
                            diagram_graphs_meta.push(g.clone());
                            Some(g)
                        } else {
                            None
                        };

                        let mut reasons: Vec<String> = Vec::new();
                        if charts > 0 {
                            reasons.push("chart".to_string());
                        }
                        if connectors > 0 {
                            reasons.push("connectors".to_string());
                        }
                        if smartart > 0 {
                            reasons.push("smartart".to_string());
                        }
                        if slide_has_diagram {
                            reasons.push("shape_diagram".to_string());
                        }
                        reasons.sort();
                        reasons.dedup();
                        let needs_semantics = !reasons.is_empty();
                        let score = if needs_semantics { 0.8 } else { 0.0 };

                        let _ = include_slide_snapshots;

                        blocks.insert(
                            block_first,
                            Block::Heading {
                                block_index: block_first,
                                level: 1,
                                text: heading_text,
                                source: SourceSpan::default(),
                            },
                        );

                        // Renumber block_index fields deterministically after insertion.
                        for (i, b) in blocks.iter_mut().enumerate() {
                            let idx = i;
                            match b {
                                Block::Heading { block_index, .. }
                                | Block::Paragraph { block_index, .. }
                                | Block::List { block_index, .. }
                                | Block::Table { block_index, .. }
                                | Block::Image { block_index, .. }
                                | Block::Link { block_index, .. } => *block_index = idx,
                            }
                        }
                        block_index = blocks.len();

                        let block_last = blocks.len().saturating_sub(1);

                        slides_meta.push(serde_json::json!({
                            "index": slide_idx_1,
                            "title": slide_title,
                            "block_first": block_first,
                            "block_last": block_last,
                            "has_diagram": slide_has_diagram,
                            "text_chars": text_chars,
                            "text_boxes": nodes.len(),
                            "images": embedded_images,
                            "tables": tables,
                            "connectors": connectors,
                            "charts": charts,
                            "smartart": smartart,
                            "needs_semantics": needs_semantics,
                            "reasons": reasons,
                            "excluded_visuals": excluded_visuals,
                            "score": score,
                            "snapshot_image_id": snapshot_image_id,
                            "diagram_graph": graph_json,
                        }));

                        slide_name = None;
                        title_text.clear();
                        slide_has_diagram = false;
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow!("odp content.xml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    if slides_meta.is_empty() {
        return Err(anyhow!("no slides found in ODP"));
    }

    let metadata_json = serde_json::json!({
        "kind": "odp",
        "slides": slides_meta,
        "charts": [],
        "diagram_graphs": diagram_graphs_meta,
    });

    Ok(ParsedOfficeDocument {
        blocks,
        images,
        metadata_json,
    })
}

pub fn parse_odp_full(bytes: &[u8]) -> Result<ParsedOfficeDocument> {
    parse_odp_full_inner(bytes, true)
}

pub fn parse(bytes: &[u8]) -> crate::Result<crate::Document> {
    let parsed = parse_odp_full_inner(bytes, true)
        .map_err(|e| crate::Error::from_parse(crate::Format::Odp, e))?;
    Ok(super::finalize(crate::Format::Odp, parsed))
}

pub fn parse_with_options(
    bytes: &[u8],
    opts: crate::odp::ParseOptions,
) -> crate::Result<crate::Document> {
    let parsed = parse_odp_full_inner(bytes, opts.include_slide_snapshots)
        .map_err(|e| crate::Error::from_parse(crate::Format::Odp, e))?;
    Ok(super::finalize(crate::Format::Odp, parsed))
}
