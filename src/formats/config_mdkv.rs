use crate::document_ast::{Block, Cell, ListItem, SourceSpan};

fn push_heading(blocks: &mut Vec<Block>, next_idx: &mut usize, level: u8, text: String) {
    blocks.push(Block::Heading {
        block_index: *next_idx,
        level,
        text,
        source: SourceSpan::default(),
    });
    *next_idx += 1;
}

fn push_paragraph(blocks: &mut Vec<Block>, next_idx: &mut usize, text: String) {
    if text.trim().is_empty() {
        return;
    }
    blocks.push(Block::Paragraph {
        block_index: *next_idx,
        text,
        source: SourceSpan::default(),
    });
    *next_idx += 1;
}

fn push_list(blocks: &mut Vec<Block>, next_idx: &mut usize, items: Vec<String>) {
    if items.is_empty() {
        return;
    }
    blocks.push(Block::List {
        block_index: *next_idx,
        ordered: false,
        items: items
            .into_iter()
            .map(|t| ListItem {
                level: 0,
                text: t,
                source: SourceSpan::default(),
            })
            .collect(),
        source: SourceSpan::default(),
    });
    *next_idx += 1;
}

fn push_table(blocks: &mut Vec<Block>, next_idx: &mut usize, rows: Vec<Vec<Cell>>) {
    if rows.is_empty() {
        return;
    }
    blocks.push(Block::Table {
        block_index: *next_idx,
        rows,
        source: SourceSpan::default(),
    });
    *next_idx += 1;
}

fn scalar_to_kv_string(v: &serde_json::Value) -> String {
    match v {
        serde_json::Value::Null => "null".to_string(),
        serde_json::Value::Bool(b) => b.to_string(),
        serde_json::Value::Number(n) => n.to_string(),
        serde_json::Value::String(s) => {
            let t = s.trim();
            if t.is_empty() {
                "\"\"".to_string()
            } else if t.contains('\n') || t.contains('\r') {
                serde_json::to_string(s).unwrap_or_else(|_| format!("\"{}\"", t))
            } else {
                t.to_string()
            }
        }
        _ => serde_json::to_string(v).unwrap_or_else(|_| "<unrenderable>".to_string()),
    }
}

fn string_is_multiline(v: &serde_json::Value) -> Option<&str> {
    match v {
        serde_json::Value::String(s) if s.contains('\n') || s.contains('\r') => Some(s.as_str()),
        _ => None,
    }
}

fn is_scalar(v: &serde_json::Value) -> bool {
    matches!(
        v,
        serde_json::Value::Null
            | serde_json::Value::Bool(_)
            | serde_json::Value::Number(_)
            | serde_json::Value::String(_)
    )
}

fn flatten_object_paths(out: &mut Vec<String>, prefix: &str, v: &serde_json::Value) -> bool {
    // Returns false if we encounter non-scalar children (arrays/objects) at any level.
    match v {
        serde_json::Value::Object(m) => {
            for (k, child) in m {
                let next = if prefix.is_empty() {
                    k.clone()
                } else {
                    format!("{prefix}.{k}")
                };
                if !flatten_object_paths(out, &next, child) {
                    return false;
                }
            }
            true
        }
        serde_json::Value::Array(_) => false,
        _ => {
            out.push(format!("{prefix}: {}", scalar_to_kv_string(v)));
            true
        }
    }
}

fn array_to_table_rows(v: &serde_json::Value) -> Option<Vec<Vec<Cell>>> {
    let items = v.as_array()?;
    if items.is_empty() {
        return None;
    }
    if !items.iter().all(|it| it.is_object()) {
        return None;
    }

    // Require that values are scalars/null to stay table-like.
    if !items
        .iter()
        .all(|it| it.as_object().unwrap().values().all(|vv| is_scalar(vv)))
    {
        return None;
    }

    let mut cols: Vec<String> = Vec::new();
    let mut seen: std::collections::HashSet<String> = Default::default();
    for it in items {
        for k in it.as_object().unwrap().keys() {
            if seen.insert(k.clone()) {
                cols.push(k.clone());
            }
        }
    }

    if cols.is_empty() {
        return None;
    }
    if cols.len() > 20 {
        return None;
    }
    let mut rows: Vec<Vec<Cell>> = Vec::new();

    let mut hdr = Vec::new();
    hdr.push(Cell {
        text: "#".to_string(),
        colspan: 1,
        rowspan: 1,
    });
    for c in &cols {
        hdr.push(Cell {
            text: c.clone(),
            colspan: 1,
            rowspan: 1,
        });
    }
    rows.push(hdr);

    for (i, it) in items.iter().enumerate() {
        let obj = it.as_object().unwrap();
        let mut r = Vec::new();
        r.push(Cell {
            text: (i as u64).to_string(),
            colspan: 1,
            rowspan: 1,
        });
        for c in &cols {
            let txt = obj.get(c).map(scalar_to_kv_string).unwrap_or_default();
            r.push(Cell {
                text: txt,
                colspan: 1,
                rowspan: 1,
            });
        }
        rows.push(r);
    }

    Some(rows)
}

fn array_to_jsonl(v: &serde_json::Value) -> String {
    let Some(items) = v.as_array() else {
        return String::new();
    };
    let mut out = String::new();
    out.push_str("```jsonl\n");
    for it in items {
        out.push_str(&serde_json::to_string(it).unwrap_or_else(|_| "null".to_string()));
        out.push('\n');
    }
    out.push_str("```\n");
    out
}

fn render_value_at_level(
    key: Option<&str>,
    v: &serde_json::Value,
    level: u8,
    blocks: &mut Vec<Block>,
    next_idx: &mut usize,
) {
    // level is current heading level for this section.

    // Scalars: accumulate as KV lines.
    if is_scalar(v) {
        if let Some(s) = string_is_multiline(v) {
            let s = s.replace("\r", "").trim().to_string();
            if s.is_empty() {
                return;
            }
            if let Some(k) = key {
                push_paragraph(blocks, next_idx, format!("{k}:\n{s}"));
            } else {
                push_paragraph(blocks, next_idx, s);
            }
            return;
        }
        if let Some(k) = key {
            push_paragraph(blocks, next_idx, format!("{k}: {}", scalar_to_kv_string(v)));
        } else {
            push_paragraph(blocks, next_idx, scalar_to_kv_string(v));
        }
        return;
    }

    // Arrays.
    if let serde_json::Value::Array(arr) = v {
        if let Some(k) = key {
            push_paragraph(blocks, next_idx, format!("{k}:"));
        }
        if arr.iter().all(is_scalar) {
            push_list(
                blocks,
                next_idx,
                arr.iter().map(scalar_to_kv_string).collect(),
            );
        } else if let Some(rows) = array_to_table_rows(v) {
            push_table(blocks, next_idx, rows);
        } else if arr.iter().all(|it| it.is_object()) {
            // Prefer readability over JSONL for nested arrays-of-objects.
            for (i, it) in arr.iter().enumerate() {
                let next_level = (level + 1).min(4);
                push_heading(
                    blocks,
                    next_idx,
                    next_level,
                    if let Some(k) = key {
                        format!("{k} {i}")
                    } else {
                        format!("item {i}")
                    },
                );
                render_value_at_level(None, it, next_level, blocks, next_idx);
            }
        } else {
            push_paragraph(blocks, next_idx, array_to_jsonl(v));
        }
        return;
    }

    // Objects.
    if let serde_json::Value::Object(map) = v {
        if level >= 4 {
            // Can't add more headings. Render scalars as KV and complex values as payload blocks.
            let mut kv_lines: Vec<String> = Vec::new();
            for (k, child) in map {
                if is_scalar(child) {
                    kv_lines.push(format!("{k}: {}", scalar_to_kv_string(child)));
                    continue;
                }

                if !kv_lines.is_empty() {
                    push_paragraph(blocks, next_idx, kv_lines.join("\n"));
                    kv_lines.clear();
                }

                match child {
                    serde_json::Value::Object(_) => {
                        let mut flat: Vec<String> = Vec::new();
                        if flatten_object_paths(&mut flat, "", child) {
                            push_paragraph(
                                blocks,
                                next_idx,
                                format!(
                                    "{k}:\n{}",
                                    flat.into_iter()
                                        .map(|l| format!("  {l}"))
                                        .collect::<Vec<_>>()
                                        .join("\n")
                                ),
                            );
                        } else {
                            let payload =
                                serde_json::to_string(child).unwrap_or_else(|_| "null".to_string());
                            push_paragraph(
                                blocks,
                                next_idx,
                                format!("{k}:\n```json\n{payload}\n```"),
                            );
                        }
                    }
                    serde_json::Value::Array(_) => {
                        push_paragraph(blocks, next_idx, format!("{k}:"));
                        render_value_at_level(None, child, level, blocks, next_idx);
                    }
                    _ => {}
                }
            }

            if !kv_lines.is_empty() {
                push_paragraph(blocks, next_idx, kv_lines.join("\n"));
            }
            return;
        }

        // Normal structured recursion.
        let mut kv_lines: Vec<String> = Vec::new();
        for (k, child) in map {
            if is_scalar(child) {
                kv_lines.push(format!("{k}: {}", scalar_to_kv_string(child)));
                continue;
            }

            if !kv_lines.is_empty() {
                push_paragraph(blocks, next_idx, kv_lines.join("\n"));
                kv_lines.clear();
            }

            let next_level = (level + 1).min(4);
            push_heading(blocks, next_idx, next_level, k.clone());
            render_value_at_level(None, child, next_level, blocks, next_idx);
        }

        if !kv_lines.is_empty() {
            push_paragraph(blocks, next_idx, kv_lines.join("\n"));
        }
        return;
    }
}

pub(crate) fn build_blocks_max_depth_4(root: &serde_json::Value) -> Vec<Block> {
    let mut blocks: Vec<Block> = Vec::new();
    let mut next_idx: usize = 0;

    match root {
        serde_json::Value::Object(map) => {
            for (k, v) in map {
                push_heading(&mut blocks, &mut next_idx, 1, k.clone());
                render_value_at_level(None, v, 1, &mut blocks, &mut next_idx);
            }
        }
        _ => {
            push_heading(&mut blocks, &mut next_idx, 1, "value".to_string());
            render_value_at_level(None, root, 1, &mut blocks, &mut next_idx);
        }
    }

    blocks
}
