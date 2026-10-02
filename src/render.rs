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
        Block::Heading {
            level,
            text,
            marks,
            ..
        } => {
            let lvl = (*level).max(1).min(6) as usize;
            format!(
                "{} {}\n\n",
                "#".repeat(lvl),
                serialize_inline(text.trim(), marks)
            )
        }
        Block::Paragraph { text, marks, .. } => format!(
            "{}\n\n",
            escape_morph_paragraph(&serialize_inline(text.trim(), marks))
        ),
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
                out.push_str(&serialize_inline(it.text.trim(), &it.marks));
                out.push('\n');
            }
            out.push('\n');
            out
        }
        Block::Table { rows, widths, .. } => {
            let cols = rows.iter().map(|r| r.len()).max().unwrap_or(0).max(1);
            if rows.is_empty() {
                return "| |\n|---|\n\n".to_string();
            }

            let mut out = String::new();
            // MorphEditor styled-table wrapper — only when the source table
            // carries column widths (an unstyled table never grows a wrapper).
            if !widths.is_empty() {
                let list = widths
                    .iter()
                    .map(|w| w.to_string())
                    .collect::<Vec<_>>()
                    .join(",");
                out.push_str(&format!("::: table {{widths=[{list}]}}\n"));
            }
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
            if !widths.is_empty() {
                out.push_str(":::\n");
            }
            out.push('\n');
            out
        }
        Block::Image { id, alt, .. } => {
            // MorphEditor media line: `![alt](path)` — kind is detected by
            // file extension. The CLI rewrites `office-image:{id}` to the
            // real asset path; alt brackets are escaped.
            let alt_s = alt
                .as_deref()
                .unwrap_or("")
                .replace('[', "\\[")
                .replace(']', "\\]");
            format!("![{alt_s}](office-image:{id})\n\n")
        }
        Block::Note { text, marks, .. } => {
            format!(
                "::: note\n{}\n:::\n\n",
                escape_morph_paragraph(&serialize_inline(text.trim(), marks))
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
            Block::Heading { level, text, marks, .. } => {
                serde_json::json!({"type":"heading","level":level,"text":text,"marks":marks})
            }
            Block::Paragraph { text, marks, .. } => serde_json::json!({"type":"paragraph","text":text,"marks":marks}),
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
            Block::Note { text, marks, .. } => serde_json::json!({"type":"note","text":text,"marks":marks}),
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
                    serde_json::json!(base64::engine::general_purpose::STANDARD.encode(&img.bytes)),
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
    use crate::Document;
    use crate::document::{DocumentMetadata, ExtractedImage, Format};

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
                description: None,
            }],
            metadata: DocumentMetadata {
                format: Format::Pptx,
                title: None,
                page_count: None,
                slide_count: Some(1),
                pdf_text_quality: None,
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
                description: None,
            }],
            metadata: DocumentMetadata {
                format: Format::Pptx,
                title: None,
                page_count: None,
                slide_count: Some(1),
                pdf_text_quality: None,
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

/// Escape one line that would re-detect as another block type in
/// MorphEditor's BlockModel (`PARAGRAPH_ESCAPE_RULES`): heading, bullet,
/// ordered item, quote, code fence, hr. Ordered items escape the dot,
/// CommonMark-style (`1\.`). Byte-mirrors `escapeParagraphSigils`.
pub(crate) fn escape_morph_line(line: &str) -> String {
    let indent_len = line.len() - line.trim_start().len();
    let (indent, rest) = line.split_at(indent_len);
    if rest.is_empty() {
        return line.to_string();
    }
    let b = rest.as_bytes();
    let hashes = b.iter().take_while(|&&c| c == b'#').count();
    let is_heading = (1..=6).contains(&hashes) && b.get(hashes) == Some(&b' ');
    let is_list =
        matches!(b.first(), Some(b'-') | Some(b'*') | Some(b'+')) && b.get(1) == Some(&b' ');
    let digits = b.iter().take_while(|c| c.is_ascii_digit()).count();
    let is_ordered = digits > 0 && b.get(digits) == Some(&b'.') && b.get(digits + 1) == Some(&b' ');
    let is_quote = b.first() == Some(&b'>');
    let is_fence = rest.starts_with("```") || rest.starts_with("~~~");
    let hr_trim = rest.trim_end();
    let is_hr = hr_trim.len() >= 3 && hr_trim.bytes().all(|c| c == b'-' || c == b'_' || c == b'*');
    if !(is_heading || is_list || is_ordered || is_quote || is_fence || is_hr) {
        return line.to_string();
    }
    if is_ordered {
        let mut s = String::with_capacity(line.len() + 1);
        s.push_str(indent);
        s.push_str(&rest[..digits]);
        s.push('\\');
        s.push_str(&rest[digits..]);
        return s;
    }
    format!("{indent}\\{rest}")
}

/// Paragraphs that ARE fenced diagrams (```mermaid … ```) stay unescaped
/// so MorphEditor parses them as code-fence blocks, not escaped text.
fn escape_morph_paragraph(text: &str) -> String {
    let t = text.trim_start();
    if t.starts_with("```") && t.trim_end().ends_with("```") {
        return text.to_string();
    }
    text.split('\n')
        .map(escape_morph_line)
        .collect::<Vec<_>>()
        .join("\n")
}

/// Render as MorphEditor-compatible markdown (the dialect BlockModel.js
/// parses): sigil-escaped paragraphs, fenced diagrams preserved, and for
/// presentations (PPTX/ODP) a `---` divider before every level-1 heading
/// — decks are hr-delimited sections in MorphEditor deck mode. Pairs with
/// the CLI's OKF frontmatter; `slides:` there selects deck rendering.
/// Column groups parsed from metadata (`multi_column_sections`):
/// (block_first, block_last, cols), non-overlapping, ascending.
fn column_sections(
    doc: &Document,
) -> Vec<(usize, usize, usize)> {
    doc.metadata
        .extra
        .get("multi_column_sections")
        .and_then(|v| v.as_array())
        .map(|arr| {
            arr.iter()
                .filter_map(|s| {
                    Some((
                        s.get("block_first")?.as_u64()? as usize,
                        s.get("block_last")?.as_u64()? as usize,
                        s.get("cols")?.as_u64()? as usize,
                    ))
                })
                .filter(|(f, l, c)| l >= f && *c >= 2)
                .collect()
        })
        .unwrap_or_default()
}

/// Split `len` blocks into `cols` even groups by cumulative text volume.
fn column_boundaries(first: usize, last: usize, cols: usize, doc: &Document) -> Vec<usize> {
    let len = last - first + 1;
    if cols >= len {
        // one block per column; extra columns stay empty
        return (0..=len).map(|k| first + k).collect();
    }
    let volumes: Vec<usize> = (first..=last)
        .map(|i| doc.blocks[i].text_len_chars().max(1))
        .collect();
    let total: usize = volumes.iter().sum();
    let mut cuts = vec![first];
    let mut acc = 0usize;
    let mut k = 1usize;
    for (i, v) in volumes.iter().enumerate() {
        acc += v;
        while k < cols && acc >= total * k / cols {
            let cut = first + i + 1;
            if cut > *cuts.last().unwrap() && cut <= last {
                cuts.push(cut);
            }
            k += 1;
        }
    }
    cuts.push(last + 1);
    cuts
}

pub fn to_morph_markdown(doc: &Document) -> String {
    let is_deck = matches!(doc.metadata.format, crate::Format::Pptx | crate::Format::Odp);
    let sections = column_sections(doc);
    let mut out = String::new();
    let mut seen_l1 = false;
    let mut idx = 0usize;
    while idx < doc.blocks.len() {
        // Multi-column section: emit a `::: columns` fence with even-volume
        // `::: column` groups (DOCX flowed columns carry no explicit split
        // points — the split is an approximation of the print layout).
        if let Some(&(f, l, c)) = sections.iter().find(|&&(f, _, _)| f == idx) {
            let cuts = column_boundaries(f, l, c, doc);
            out.push_str("::: columns\n");
            for w in cuts.windows(2) {
                out.push_str("::: column\n");
                for b in &doc.blocks[w[0]..w[1]] {
                    out.push_str(&render_block(b));
                }
                out.push_str(":::\n");
            }
            out.push_str(":::\n\n");
            idx = l + 1;
            continue;
        }
        let b = &doc.blocks[idx];
        if is_deck
            && let Block::Heading { level: 1, .. } = b
        {
            if seen_l1 {
                out.push_str("---\n\n");
            }
            seen_l1 = true;
        }
        out.push_str(&render_block(b));
        idx += 1;
    }
    out
}

#[cfg(test)]
mod morph_tests {
    use super::*;

    #[test]
    fn escape_morph_line_mirrors_blockmodel_rules() {
        // PARAGRAPH_ESCAPE_RULES: heading, bullet, ordered, quote, fence, hr
        assert_eq!(escape_morph_line("# Title"), "\\# Title");
        assert_eq!(escape_morph_line("## Sub"), "\\## Sub");
        assert_eq!(escape_morph_line("- item"), "\\- item");
        assert_eq!(escape_morph_line("* star"), "\\* star");
        assert_eq!(escape_morph_line("1. first"), "1\\. first");
        assert_eq!(escape_morph_line("10. tenth"), "10\\. tenth");
        assert_eq!(escape_morph_line("> quote"), "\\> quote");
        assert_eq!(escape_morph_line("```rust"), "\\```rust");
        assert_eq!(escape_morph_line("---"), "\\---");
        assert_eq!(escape_morph_line("___"), "\\___");
        // indented lines keep their indent
        assert_eq!(escape_morph_line("  - nested"), "  \\- nested");
        // grammar-mirroring, not blanket: these stay bare
        assert_eq!(escape_morph_line("plain text"), "plain text");
        assert_eq!(escape_morph_line("#hashtag"), "#hashtag");
        assert_eq!(escape_morph_line("-no-space"), "-no-space");
        assert_eq!(escape_morph_line("*bold* leader"), "*bold* leader");
        assert_eq!(escape_morph_line("1"), "1");
        assert_eq!(escape_morph_line(""), "");
    }

    #[test]
    fn fenced_diagram_paragraphs_stay_unescaped() {
        let md = "```mermaid\nflowchart LR\n  n1@{ shape: rect, label: \"a\" }\n```";
        assert_eq!(escape_morph_paragraph(md), md);
        // a lone fence line inside prose still escapes
        assert_eq!(escape_morph_paragraph("see:\n```"), "see:\n\\```");
    }
}

/// Render clean text + inline marks as MorphEditor markdown.
/// Port of BlockModel serializeMarks semantics: same-range bold+italic
/// clusters as `***`, adjacent same-kind marks merge, partial overlaps
/// nest (`**bo*ld*er**`), links wrap structurally (`[text](href)`).
pub(crate) fn serialize_inline(text: &str, marks: &[crate::document_ast::InlineMark]) -> String {
    use crate::document_ast::MarkKind;
    if marks.is_empty() {
        return text.to_string();
    }
    let chars: Vec<char> = text.chars().collect();
    let n = chars.len();
    // (start, end, open_marker, close_marker) — close markers may embed href
    let mut norm: Vec<(usize, usize, &'static str, String)> = marks
        .iter()
        .filter_map(|m| {
            if m.start >= m.end || m.end > n {
                return None;
            }
            let spec: (usize, &'static str, String) = match &m.kind {
                MarkKind::Bold => (0, "**", "**".into()),
                MarkKind::Italic => (1, "*", "*".into()),
                MarkKind::Code => (2, "`", "`".into()),
                MarkKind::Strikethrough => (3, "~~", "~~".into()),
                MarkKind::Link { url } => (4, "[", format!("]({})", url)),
            };
            Some((m.start, m.end, spec.1, spec.2))
        })
        .collect();
    norm.sort_by_key(|(s, e, open, close)| (*s, *e, open.len(), close.clone()));
    // merge adjacent same-kind marks
    let mut merged: Vec<(usize, usize, &'static str, String)> = Vec::new();
    for m in norm {
        if let Some(last) = merged.last_mut()
            && last.1 == m.0
            && last.2 == m.2
            && last.3 == m.3
        {
            last.1 = m.1;
            continue;
        }
        merged.push(m);
    }

    let mut points: Vec<usize> = merged.iter().flat_map(|m| [m.0, m.1]).collect();
    points.sort_unstable();
    points.dedup();

    let mut out = String::new();
    let mut prev = 0usize;
    for &pt in &points {
        if pt > prev {
            out.extend(chars[prev..pt].iter());
        }
        // close marks ending here (reverse order so nested wrappers close innermost-first)
        for (s, e, _open, close) in merged.iter().rev() {
            if *e == pt {
                let _ = s;
                out.push_str(close);
            }
        }
        // open marks starting here (order: links outermost, then **, *, `, ~~)
        for (s, e, open, close) in merged.iter() {
            if *s == pt {
                let _ = (e, close);
                out.push_str(open);
            }
        }
        prev = pt;
    }
    if prev < n {
        out.extend(chars[prev..].iter());
    }
    out
}

#[cfg(test)]
mod marks_tests {
    use super::*;
    use crate::document_ast::{InlineMark, MarkKind};

    fn mk(kind: MarkKind, start: usize, end: usize) -> InlineMark {
        InlineMark { kind, start, end }
    }

    #[test]
    fn serialize_inline_matches_blockmodel_semantics() {
        assert_eq!(serialize_inline("bold and more", &[mk(MarkKind::Bold, 0, 4)]), "**bold** and more");
        assert_eq!(
            serialize_inline("bold and it", &[mk(MarkKind::Bold, 0, 4), mk(MarkKind::Italic, 9, 11)]),
            "**bold** and *it*"
        );
        // full overlap clusters as ***
        assert_eq!(
            serialize_inline("both", &[mk(MarkKind::Bold, 0, 4), mk(MarkKind::Italic, 0, 4)]),
            "***both***"
        );
        // code + strike
        assert_eq!(
            serialize_inline("code here", &[mk(MarkKind::Code, 0, 4)]),
            "`code` here"
        );
        assert_eq!(
            serialize_inline("gone", &[mk(MarkKind::Strikethrough, 0, 4)]),
            "~~gone~~"
        );
        // link
        assert_eq!(
            serialize_inline(
                "link text tail",
                &[mk(MarkKind::Link { url: "https://x.y".into() }, 0, 9)]
            ),
            "[link text](https://x.y) tail"
        );
        // partial overlap nests
        assert_eq!(
            serialize_inline("border", &[mk(MarkKind::Bold, 0, 6), mk(MarkKind::Italic, 2, 4)]),
            "**bo*rd*er**"
        );
        // adjacent same-kind merges
        assert_eq!(
            serialize_inline("abcd", &[mk(MarkKind::Bold, 0, 2), mk(MarkKind::Bold, 2, 4)]),
            "**abcd**"
        );
        // empty text / no marks passthrough
        assert_eq!(serialize_inline("plain", &[]), "plain");
    }
}
