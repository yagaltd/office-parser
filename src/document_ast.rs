use serde::{Deserialize, Serialize};

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum SpanKind {
    /// Byte offsets into the rendered extracted text.
    ExtractedTextByte,
}

impl Default for SpanKind {
    fn default() -> Self {
        SpanKind::ExtractedTextByte
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct SourceSpan {
    pub kind: SpanKind,
    pub start: usize,
    pub end: usize,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct ListItem {
    pub level: u8,
    pub text: String,
    pub source: SourceSpan,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct Cell {
    pub text: String,
    pub colspan: usize,
    pub rowspan: usize,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum LinkKind {
    Web,
    Video,
    Audio,
    Unknown,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum Block {
    Heading {
        block_index: usize,
        level: u8,
        text: String,
        source: SourceSpan,
    },
    Paragraph {
        block_index: usize,
        text: String,
        source: SourceSpan,
    },
    List {
        block_index: usize,
        ordered: bool,
        items: Vec<ListItem>,
        source: SourceSpan,
    },
    Table {
        block_index: usize,
        rows: Vec<Vec<Cell>>,
        source: SourceSpan,
    },
    Image {
        block_index: usize,
        id: String,
        filename: Option<String>,
        content_type: Option<String>,
        alt: Option<String>,
        source: SourceSpan,
    },
    Link {
        block_index: usize,
        url: String,
        text: Option<String>,
        kind: LinkKind,
        source: SourceSpan,
    },
}

impl Block {
    pub fn block_index(&self) -> usize {
        match self {
            Block::Heading { block_index, .. }
            | Block::Paragraph { block_index, .. }
            | Block::List { block_index, .. }
            | Block::Table { block_index, .. }
            | Block::Image { block_index, .. }
            | Block::Link { block_index, .. } => *block_index,
        }
    }

    pub fn source(&self) -> SourceSpan {
        match self {
            Block::Heading { source, .. }
            | Block::Paragraph { source, .. }
            | Block::List { source, .. }
            | Block::Table { source, .. }
            | Block::Image { source, .. }
            | Block::Link { source, .. } => *source,
        }
    }

    pub fn set_source(&mut self, span: SourceSpan) {
        match self {
            Block::Heading { source, .. }
            | Block::Paragraph { source, .. }
            | Block::List { source, .. }
            | Block::Table { source, .. }
            | Block::Image { source, .. }
            | Block::Link { source, .. } => {
                *source = span;
            }
        }
    }

    pub fn text_len_chars(&self) -> usize {
        match self {
            Block::Heading { text, .. } => text.chars().count(),
            Block::Paragraph { text, .. } => text.chars().count(),
            Block::List { items, .. } => items.iter().map(|i| i.text.chars().count()).sum(),
            Block::Table { rows, .. } => rows
                .iter()
                .flat_map(|r| r.iter())
                .map(|c| c.text.chars().count())
                .sum(),
            Block::Image {
                id, filename, alt, ..
            } => {
                id.chars().count()
                    + filename.as_deref().unwrap_or("").chars().count()
                    + alt.as_deref().unwrap_or("").chars().count()
            }
            Block::Link { url, text, .. } => {
                url.chars().count() + text.as_deref().unwrap_or("").chars().count()
            }
        }
    }
}

fn classify_link_kind(url: &str) -> LinkKind {
    let u = url.trim().to_ascii_lowercase();
    if u.contains("youtube.com") || u.contains("youtu.be") || u.contains("vimeo.com") {
        return LinkKind::Video;
    }
    if u.ends_with(".mp3") || u.ends_with(".wav") || u.ends_with(".m4a") || u.ends_with(".flac") {
        return LinkKind::Audio;
    }
    if u.starts_with("http://") || u.starts_with("https://") {
        return LinkKind::Web;
    }
    LinkKind::Unknown
}

fn escape_markdown_table_cell(s: &str) -> String {
    s.replace('|', "\\|").replace('\n', " ").replace('\r', " ")
}

/// Deterministically render blocks into a markdown-ish extracted text string.
///
/// This function also assigns `SourceSpan` byte offsets for every block (and list item) into the
/// returned string.
pub fn render_blocks_to_extracted_text(blocks: &mut [Block]) -> String {
    let mut out = String::new();

    for b in blocks.iter_mut() {
        let start = out.len();

        match b {
            Block::Heading {
                level,
                text,
                source,
                ..
            } => {
                let lvl = (*level).max(1).min(6) as usize;
                out.push_str(&"#".repeat(lvl));
                out.push(' ');
                out.push_str(text.trim());
                out.push_str("\n\n");
                *source = SourceSpan {
                    kind: SpanKind::ExtractedTextByte,
                    start,
                    end: out.len(),
                };
            }
            Block::Paragraph { text, source, .. } => {
                out.push_str(text.trim());
                out.push_str("\n\n");
                *source = SourceSpan {
                    kind: SpanKind::ExtractedTextByte,
                    start,
                    end: out.len(),
                };
            }
            Block::List {
                ordered,
                items,
                source,
                ..
            } => {
                for it in items.iter_mut() {
                    let it_start = out.len();
                    let indent = "  ".repeat(it.level as usize);
                    out.push_str(&indent);
                    if *ordered {
                        out.push_str("1. ");
                    } else {
                        out.push_str("- ");
                    }
                    out.push_str(it.text.trim());
                    out.push('\n');
                    it.source = SourceSpan {
                        kind: SpanKind::ExtractedTextByte,
                        start: it_start,
                        end: out.len(),
                    };
                }
                out.push('\n');
                *source = SourceSpan {
                    kind: SpanKind::ExtractedTextByte,
                    start,
                    end: out.len(),
                };
            }
            Block::Table { rows, source, .. } => {
                let cols = rows.iter().map(|r| r.len()).max().unwrap_or(0).max(1);

                if rows.is_empty() {
                    out.push_str("| |\n|---|\n\n");
                } else {
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

                    // Header row.
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
                }

                *source = SourceSpan {
                    kind: SpanKind::ExtractedTextByte,
                    start,
                    end: out.len(),
                };
            }
            Block::Image {
                id,
                filename,
                content_type,
                alt,
                source,
                ..
            } => {
                let alt_s = alt.as_deref().unwrap_or("");
                let file_s = filename.as_deref().unwrap_or("");
                let ct_s = content_type.as_deref().unwrap_or("");
                out.push_str(&format!(
                    "[image:{} filename=\"{}\" content_type=\"{}\" alt=\"{}\"]\n\n",
                    id, file_s, ct_s, alt_s
                ));
                *source = SourceSpan {
                    kind: SpanKind::ExtractedTextByte,
                    start,
                    end: out.len(),
                };
            }
            Block::Link {
                url,
                text,
                kind,
                source,
                ..
            } => {
                *kind = classify_link_kind(url);
                let t = text.as_deref().unwrap_or("");
                if t.is_empty() {
                    out.push_str(url.trim());
                } else {
                    out.push_str(&format!("[{}]({})", t.trim(), url.trim()));
                }
                out.push_str("\n\n");
                *source = SourceSpan {
                    kind: SpanKind::ExtractedTextByte,
                    start,
                    end: out.len(),
                };
            }
        }
    }

    out
}

pub fn blocks_to_plain_text(blocks: &mut [Block]) -> String {
    render_blocks_to_extracted_text(blocks)
}
