use crate::document_ast::{Block, Cell};
use crate::{Document, Result};
use base64::Engine;

#[derive(Clone, Debug)]
pub struct Chunk {
    pub content: String,
    pub block_first: usize,
    pub block_last: usize,
}

#[derive(Clone, Copy, Debug)]
pub struct JsonRenderOptions {
    pub include_image_bytes: bool,
}

impl Default for JsonRenderOptions {
    fn default() -> Self {
        Self {
            include_image_bytes: true,
        }
    }
}

pub fn to_markdown(doc: &Document) -> String {
    let mut blocks = doc.blocks.clone();
    crate::document_ast::render_blocks_to_extracted_text(&mut blocks)
}

fn escape_markdown_table_cell(s: &str) -> String {
    s.replace('|', "\\|").replace('\n', " ").replace('\r', " ")
}

fn render_block(b: &Block) -> String {
    match b {
        Block::Heading { level, text, .. } => {
            let lvl = (*level).max(1).min(6) as usize;
            format!("{} {}\n\n", "#".repeat(lvl), text.trim())
        }
        Block::Paragraph { text, .. } => format!("{}\n\n", text.trim()),
        Block::List { ordered, items, .. } => {
            let mut out = String::new();
            for it in items {
                let indent = "  ".repeat(it.level as usize);
                out.push_str(&indent);
                if *ordered {
                    out.push_str("1. ");
                } else {
                    out.push_str("- ");
                }
                out.push_str(it.text.trim());
                out.push('\n');
            }
            out.push('\n');
            out
        }
        Block::Table { rows, .. } => {
            let cols = rows.iter().map(|r| r.len()).max().unwrap_or(0).max(1);
            if rows.is_empty() {
                return "| |\n|---|\n\n".to_string();
            }

            let mut out = String::new();
            let row_to_line = |out: &mut String, row: &[Cell]| {
                out.push('|');
                for c in 0..cols {
                    let txt = row
                        .get(c)
                        .map(|cc| escape_markdown_table_cell(cc.text.as_str()))
                        .unwrap_or_default();
                    out.push(' ');
                    out.push_str(&txt);
                    out.push(' ');
                    out.push('|');
                }
                out.push('\n');
            };

            row_to_line(&mut out, &rows[0]);
            out.push('|');
            for _ in 0..cols {
                out.push_str("---|");
            }
            out.push('\n');
            for r in rows.iter().skip(1) {
                row_to_line(&mut out, r);
            }
            out.push('\n');
            out
        }
        Block::Image {
            id,
            filename,
            content_type,
            alt,
            ..
        } => {
            let alt_s = alt.as_deref().unwrap_or("");
            let file_s = filename.as_deref().unwrap_or("");
            let ct_s = content_type.as_deref().unwrap_or("");
            format!(
                "[image:{} filename=\"{}\" content_type=\"{}\" alt=\"{}\"]\n\n",
                id, file_s, ct_s, alt_s
            )
        }
        Block::Link { url, text, .. } => {
            let t = text.as_deref().unwrap_or("");
            if t.trim().is_empty() {
                format!("{}\n\n", url.trim())
            } else {
                format!("[{}]({})\n\n", t.trim(), url.trim())
            }
        }
    }
}

fn render_table_pieces(rows: &[Vec<Cell>], max_chars: usize) -> Vec<String> {
    let max_chars = max_chars.max(1);
    let cols = rows.iter().map(|r| r.len()).max().unwrap_or(0).max(1);
    if rows.is_empty() {
        return vec!["| |\n|---|\n\n".to_string()];
    }

    let row_to_line = |row: &[Cell]| {
        let mut out = String::new();
        out.push('|');
        for c in 0..cols {
            let txt = row
                .get(c)
                .map(|cc| escape_markdown_table_cell(cc.text.as_str()))
                .unwrap_or_default();
            out.push(' ');
            out.push_str(&txt);
            out.push(' ');
            out.push('|');
        }
        out.push('\n');
        out
    };

    let mut header = row_to_line(&rows[0]);
    header.push('|');
    for _ in 0..cols {
        header.push_str("---|");
    }
    header.push('\n');

    let full = {
        let mut out = String::new();
        out.push_str(&header);
        for r in rows.iter().skip(1) {
            out.push_str(&row_to_line(r));
        }
        out.push('\n');
        out
    };
    if full.chars().count() <= max_chars {
        return vec![full];
    }

    if header.chars().count() >= max_chars {
        return vec![full];
    }

    let mut pieces: Vec<String> = Vec::new();
    let mut cur = header.clone();

    for r in rows.iter().skip(1) {
        let line = row_to_line(r);
        if cur.chars().count() + line.chars().count() + 1 > max_chars && cur != header {
            cur.push('\n');
            pieces.push(std::mem::take(&mut cur));
            cur = header.clone();
        }
        cur.push_str(&line);
    }

    if cur != header {
        cur.push('\n');
        pieces.push(cur);
    } else {
        pieces.push(full);
    }

    pieces
}

fn render_block_pieces(b: &Block, max_chars: usize) -> Vec<String> {
    match b {
        Block::Table { rows, .. } => render_table_pieces(rows, max_chars),
        _ => vec![render_block(b)],
    }
}

pub fn to_chunks(doc: &Document, max_chars: usize) -> Vec<Chunk> {
    let max_chars = max_chars.max(1);
    let mut out: Vec<Chunk> = Vec::new();

    let mut cur = String::new();
    let mut cur_first: Option<usize> = None;
    let mut cur_last: Option<usize> = None;

    for (i, b) in doc.blocks.iter().enumerate() {
        for piece in render_block_pieces(b, max_chars) {
            let cur_len = cur.chars().count();
            let piece_len = piece.chars().count();

            if !cur.is_empty() && cur_len + piece_len > max_chars {
                out.push(Chunk {
                    content: std::mem::take(&mut cur),
                    block_first: cur_first.unwrap_or(i),
                    block_last: cur_last.unwrap_or(i),
                });
                cur_first = None;
            }

            if cur_first.is_none() {
                cur_first = Some(i);
            }
            cur_last = Some(i);
            cur.push_str(&piece);

            // If one block (or table chunk) is enormous, still emit it as its own chunk.
            if cur.chars().count() >= max_chars {
                out.push(Chunk {
                    content: std::mem::take(&mut cur),
                    block_first: cur_first.unwrap_or(i),
                    block_last: cur_last.unwrap_or(i),
                });
                cur_first = None;
            }
        }
    }

    if !cur.is_empty() {
        out.push(Chunk {
            content: cur,
            block_first: cur_first.unwrap_or(0),
            block_last: cur_last.unwrap_or(0),
        });
    }

    out
}

pub fn to_json(doc: &Document) -> Result<String> {
    to_json_with_options(doc, JsonRenderOptions::default())
}

pub fn to_json_with_options(doc: &Document, opts: JsonRenderOptions) -> Result<String> {
    let v = to_json_value_with_options(doc, opts);
    Ok(serde_json::to_string_pretty(&v).map_err(|e| crate::Error::Render(e.into()))?)
}

pub fn to_json_value(doc: &Document) -> serde_json::Value {
    to_json_value_with_options(doc, JsonRenderOptions::default())
}

pub fn to_json_value_with_options(doc: &Document, opts: JsonRenderOptions) -> serde_json::Value {
    let blocks = doc
        .blocks
        .iter()
        .map(|b| match b {
            Block::Heading { level, text, .. } => {
                serde_json::json!({"type":"heading","level":level,"text":text})
            }
            Block::Paragraph { text, .. } => serde_json::json!({"type":"paragraph","text":text}),
            Block::List { ordered, items, .. } => serde_json::json!({
                "type":"list",
                "ordered": ordered,
                "items": items.iter().map(|it| serde_json::json!({"level":it.level,"text":it.text})).collect::<Vec<_>>()
            }),
            Block::Table { rows, .. } => serde_json::json!({
                "type":"table",
                "rows": rows.iter().map(|r| r.iter().map(|c| serde_json::json!({"text":c.text,"colspan":c.colspan,"rowspan":c.rowspan})).collect::<Vec<_>>()).collect::<Vec<_>>()
            }),
            Block::Image { id, content_type, alt, .. } => {
                serde_json::json!({"type":"image","id":id,"mime":content_type,"alt":alt})
            }
            Block::Link { url, text, .. } => serde_json::json!({"type":"link","url":url,"text":text}),
        })
        .collect::<Vec<_>>();

    let images = doc
        .images
        .iter()
        .map(|img| {
            let mut obj = serde_json::Map::new();
            obj.insert("id".to_string(), serde_json::json!(img.id));
            obj.insert("mime".to_string(), serde_json::json!(img.mime_type));
            obj.insert("filename".to_string(), serde_json::json!(img.filename));
            obj.insert("source_ref".to_string(), serde_json::json!(img.source_ref));
            if opts.include_image_bytes {
                obj.insert(
                    "bytes_b64".to_string(),
                    serde_json::json!(
                        base64::engine::general_purpose::STANDARD.encode(&img.bytes)
                    ),
                );
            }
            serde_json::Value::Object(obj)
        })
        .collect::<Vec<_>>();

    serde_json::json!({
        "schema_version": 1,
        "metadata": {
            "format": doc.metadata.format.as_str(),
            "title": doc.metadata.title,
            "page_count": doc.metadata.page_count,
            "slide_count": doc.metadata.slide_count,
            "extra": doc.metadata.extra,
        },
        "blocks": blocks,
        "images": images,
    })
}

#[cfg(test)]
mod tests {
    use super::{JsonRenderOptions, to_json_value, to_json_value_with_options};
    use crate::document::{DocumentMetadata, ExtractedImage, Format};
    use crate::Document;

    #[test]
    fn json_render_includes_image_bytes_by_default() {
        let doc = Document {
            blocks: vec![],
            images: vec![ExtractedImage {
                bytes: vec![1, 2, 3],
                mime_type: "image/png".to_string(),
                filename: Some("asset/image.png".to_string()),
                source_ref: Some("slide:1:snapshot".to_string()),
                id: "sha256:abc".to_string(),
            }],
            metadata: DocumentMetadata {
                format: Format::Pptx,
                title: None,
                page_count: None,
                slide_count: Some(1),
                extra: serde_json::json!({}),
            },
        };

        let v = to_json_value(&doc);
        assert!(v["images"][0].get("bytes_b64").is_some());
    }

    #[test]
    fn json_render_can_omit_image_bytes() {
        let doc = Document {
            blocks: vec![],
            images: vec![ExtractedImage {
                bytes: vec![1, 2, 3],
                mime_type: "image/png".to_string(),
                filename: Some("asset/image.png".to_string()),
                source_ref: Some("slide:1:snapshot".to_string()),
                id: "sha256:abc".to_string(),
            }],
            metadata: DocumentMetadata {
                format: Format::Pptx,
                title: None,
                page_count: None,
                slide_count: Some(1),
                extra: serde_json::json!({}),
            },
        };

        let v = to_json_value_with_options(
            &doc,
            JsonRenderOptions {
                include_image_bytes: false,
            },
        );
        assert!(v["images"][0].get("bytes_b64").is_none());
        assert_eq!(v["images"][0]["id"], "sha256:abc");
    }
}
