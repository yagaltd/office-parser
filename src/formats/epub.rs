use anyhow::{Context, Result, anyhow};
use quick_xml::Reader;
use quick_xml::events::Event;

use super::{ParsedImage, ParsedOfficeDocument, mime_from_filename, sha256_hex};
use crate::document_ast::{Block, LinkKind, ListItem, SourceSpan};
use crate::{Document, Format};

pub fn parse(bytes: &[u8]) -> crate::Result<Document> {
    let parsed = parse_epub_full(bytes).map_err(|e| crate::Error::from_parse(Format::Epub, e))?;
    Ok(super::finalize(Format::Epub, parsed))
}

fn strip_fragment_and_query(s: &str) -> &str {
    let s = s.split('#').next().unwrap_or(s);
    s.split('?').next().unwrap_or(s)
}

fn resolve_path(base_file: &str, href: &str) -> String {
    let href = strip_fragment_and_query(href);
    let base = std::path::Path::new(base_file);
    let base_dir = base.parent().unwrap_or(std::path::Path::new(""));
    let joined = base_dir.join(href);
    let mut out = std::path::PathBuf::new();
    for c in joined.components() {
        use std::path::Component;
        match c {
            Component::CurDir => {}
            Component::ParentDir => {
                out.pop();
            }
            other => out.push(other.as_os_str()),
        }
    }
    out.to_string_lossy().replace('\\', "/")
}

fn normalize_ws(s: &str) -> String {
    s.replace('\r', "")
        .split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
        .trim()
        .to_string()
}

fn push_heading(blocks: &mut Vec<Block>, next_idx: &mut usize, level: u8, text: String) {
    let text = text.trim().to_string();
    if text.is_empty() {
        return;
    }
    blocks.push(Block::Heading {
        block_index: *next_idx,
        level,
        text,
        source: SourceSpan::default(),
    });
    *next_idx += 1;
}

fn push_paragraph(blocks: &mut Vec<Block>, next_idx: &mut usize, text: String) {
    let text = text.trim().to_string();
    if text.is_empty() {
        return;
    }
    blocks.push(Block::Paragraph {
        block_index: *next_idx,
        text,
        source: SourceSpan::default(),
    });
    *next_idx += 1;
}

fn push_list(blocks: &mut Vec<Block>, next_idx: &mut usize, ordered: bool, items: Vec<String>) {
    let items: Vec<ListItem> = items
        .into_iter()
        .map(|t| normalize_ws(&t))
        .filter(|t| !t.is_empty())
        .map(|t| ListItem {
            level: 0,
            text: t,
            source: SourceSpan::default(),
        })
        .collect();
    if items.is_empty() {
        return;
    }
    blocks.push(Block::List {
        block_index: *next_idx,
        ordered,
        items,
        source: SourceSpan::default(),
    });
    *next_idx += 1;
}

fn push_link(blocks: &mut Vec<Block>, next_idx: &mut usize, url: String, text: Option<String>) {
    let url = url.trim().to_string();
    if url.is_empty() {
        return;
    }
    let text = text.map(|s| normalize_ws(&s)).filter(|s| !s.is_empty());
    blocks.push(Block::Link {
        block_index: *next_idx,
        url: url.clone(),
        text,
        kind: LinkKind::Unknown,
        source: SourceSpan::default(),
    });
    *next_idx += 1;
}

fn parse_container_opf_path(container_xml: &str) -> Result<String> {
    let mut reader = Reader::from_str(container_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                if e.local_name().as_ref() == b"rootfile" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref().ends_with(b"full-path") {
                            let v = a.unescape_value().context("decode rootfile@full-path")?;
                            let v = v.trim();
                            if !v.is_empty() {
                                return Ok(v.to_string());
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(anyhow!("parse container.xml: {e}")),
        }
        buf.clear();
    }
    anyhow::bail!("container.xml missing rootfile@full-path")
}

#[derive(Clone, Debug)]
struct ManifestItem {
    href: String,
    media_type: String,
}

fn parse_opf(
    opf_xml: &str,
) -> Result<(
    Option<String>,
    std::collections::HashMap<String, ManifestItem>,
    Vec<String>,
)> {
    let mut reader = Reader::from_str(opf_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut title: Option<String> = None;
    let mut in_title = false;

    let mut manifest: std::collections::HashMap<String, ManifestItem> = Default::default();
    let mut spine: Vec<String> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let ln = e.local_name();
                let ln = ln.as_ref();
                if ln == b"title" {
                    in_title = true;
                }
                if ln == b"item" {
                    let mut id: Option<String> = None;
                    let mut href: Option<String> = None;
                    let mut mt: Option<String> = None;
                    for a in e.attributes().flatten() {
                        let k = a.key.as_ref();
                        let v = a.unescape_value().ok().map(|v| v.to_string());
                        if k.ends_with(b"id") {
                            id = v;
                        } else if k.ends_with(b"href") {
                            href = v;
                        } else if k.ends_with(b"media-type") {
                            mt = v;
                        }
                    }
                    if let (Some(id), Some(href), Some(mt)) = (id, href, mt) {
                        manifest.insert(
                            id,
                            ManifestItem {
                                href,
                                media_type: mt,
                            },
                        );
                    }
                }
                if ln == b"itemref" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref().ends_with(b"idref") {
                            if let Ok(v) = a.unescape_value() {
                                let v = v.trim();
                                if !v.is_empty() {
                                    spine.push(v.to_string());
                                }
                            }
                        }
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let ln = e.local_name();
                let ln = ln.as_ref();
                if ln == b"item" {
                    let mut id: Option<String> = None;
                    let mut href: Option<String> = None;
                    let mut mt: Option<String> = None;
                    for a in e.attributes().flatten() {
                        let k = a.key.as_ref();
                        let v = a.unescape_value().ok().map(|v| v.to_string());
                        if k.ends_with(b"id") {
                            id = v;
                        } else if k.ends_with(b"href") {
                            href = v;
                        } else if k.ends_with(b"media-type") {
                            mt = v;
                        }
                    }
                    if let (Some(id), Some(href), Some(mt)) = (id, href, mt) {
                        manifest.insert(
                            id,
                            ManifestItem {
                                href,
                                media_type: mt,
                            },
                        );
                    }
                }
                if ln == b"itemref" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref().ends_with(b"idref") {
                            if let Ok(v) = a.unescape_value() {
                                let v = v.trim();
                                if !v.is_empty() {
                                    spine.push(v.to_string());
                                }
                            }
                        }
                    }
                }
            }
            Ok(Event::Text(t)) => {
                if in_title && title.is_none() {
                    if let Ok(s) = t.decode() {
                        let s = normalize_ws(&s);
                        if !s.is_empty() {
                            title = Some(s);
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
            Ok(_) => {}
            Err(e) => return Err(anyhow!("parse opf: {e}")),
        }
        buf.clear();
    }

    Ok((title, manifest, spine))
}

fn append_text_glue(out: &mut String, s: &str) {
    if s.is_empty() {
        return;
    }
    let last = out.chars().rev().find(|c| !c.is_whitespace());
    let first = s.chars().find(|c| !c.is_whitespace());
    if matches!(last, Some(c) if c.is_ascii_alphanumeric())
        && matches!(first, Some(c) if c.is_ascii_alphanumeric())
        && !out.ends_with([' ', '\t', '\n'])
        && !s.starts_with([' ', '\t', '\n'])
    {
        out.push(' ');
    }
    out.push_str(s);
}

fn parse_xhtml_to_blocks(
    xhtml_bytes: &[u8],
    xhtml_path: &str,
    bytes: &[u8],
    manifest_href_to_media: &std::collections::HashMap<String, String>,
    blocks: &mut Vec<Block>,
    images: &mut Vec<ParsedImage>,
    next_idx: &mut usize,
    image_by_id: &mut std::collections::HashMap<String, usize>,
) -> Result<()> {
    let cur = std::io::Cursor::new(xhtml_bytes);
    let mut reader = Reader::from_reader(std::io::BufReader::new(cur));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut in_title = false;
    let mut title_buf = String::new();

    let heading_pos = blocks.len();
    blocks.push(Block::Heading {
        block_index: *next_idx,
        level: 1,
        text: "chapter".to_string(),
        source: SourceSpan::default(),
    });
    *next_idx += 1;

    let mut cur_p: Option<String> = None;
    let mut cur_heading_level: Option<u8> = None;
    let mut cur_heading_text = String::new();

    let mut list_stack: Vec<bool> = Vec::new(); // ordered?
    let mut cur_li: Option<String> = None;
    let mut cur_list_items: Vec<String> = Vec::new();

    let mut in_a: Option<String> = None;
    let mut a_text = String::new();

    let flush_p = |blocks: &mut Vec<Block>, next_idx: &mut usize, cur_p: &mut Option<String>| {
        if let Some(p) = cur_p.take() {
            let p = normalize_ws(&p);
            if !p.is_empty() {
                push_paragraph(blocks, next_idx, p);
            }
        }
    };

    let flush_heading = |blocks: &mut Vec<Block>,
                         next_idx: &mut usize,
                         cur_heading_level: &mut Option<u8>,
                         cur_heading_text: &mut String| {
        if let Some(lvl) = cur_heading_level.take() {
            let t = normalize_ws(cur_heading_text);
            cur_heading_text.clear();
            if !t.is_empty() {
                push_heading(blocks, next_idx, lvl, t);
            }
        }
    };

    let flush_list =
        |blocks: &mut Vec<Block>, next_idx: &mut usize, ordered: bool, items: &mut Vec<String>| {
            if !items.is_empty() {
                let items0 = std::mem::take(items);
                push_list(blocks, next_idx, ordered, items0);
            }
        };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = String::from_utf8_lossy(e.local_name().as_ref()).to_ascii_lowercase();
                match name.as_str() {
                    "title" => {
                        in_title = true;
                    }
                    "p" => {
                        flush_heading(
                            blocks,
                            next_idx,
                            &mut cur_heading_level,
                            &mut cur_heading_text,
                        );
                        flush_p(blocks, next_idx, &mut cur_p);
                        cur_p = Some(String::new());
                    }
                    "br" => {
                        if let Some(p) = cur_p.as_mut() {
                            p.push('\n');
                        } else if cur_heading_level.is_some() {
                            cur_heading_text.push('\n');
                        } else if let Some(li) = cur_li.as_mut() {
                            li.push('\n');
                        }
                    }
                    "h1" | "h2" | "h3" | "h4" | "h5" | "h6" => {
                        flush_p(blocks, next_idx, &mut cur_p);
                        flush_heading(
                            blocks,
                            next_idx,
                            &mut cur_heading_level,
                            &mut cur_heading_text,
                        );
                        let lvl = name[1..].parse::<u8>().unwrap_or(2).clamp(1, 6);
                        cur_heading_level = Some(lvl);
                        cur_heading_text.clear();
                    }
                    "ul" => {
                        flush_p(blocks, next_idx, &mut cur_p);
                        list_stack.push(false);
                    }
                    "ol" => {
                        flush_p(blocks, next_idx, &mut cur_p);
                        list_stack.push(true);
                    }
                    "li" => {
                        cur_li = Some(String::new());
                    }
                    "img" => {
                        // We handle image extraction on Empty too, but some XHTML uses Start+End.
                        let mut src: Option<String> = None;
                        let mut alt: Option<String> = None;
                        for a in e.attributes().flatten() {
                            let k = a.key.as_ref();
                            if k.ends_with(b"src") {
                                if let Ok(v) = a.unescape_value() {
                                    src = Some(v.to_string());
                                }
                            } else if k.ends_with(b"alt") {
                                if let Ok(v) = a.unescape_value() {
                                    alt = Some(v.to_string());
                                }
                            }
                        }
                        if let Some(src) = src {
                            let img_zip_path = resolve_path(xhtml_path, &src);
                            let img_bytes = super::read_zip_file(bytes, &img_zip_path)
                                .with_context(|| format!("read image {img_zip_path}"))?;
                            let hash = sha256_hex(&img_bytes);
                            let id = format!("sha256:{hash}");
                            let mime = manifest_href_to_media
                                .get(strip_fragment_and_query(&src).trim())
                                .cloned()
                                .unwrap_or_else(|| mime_from_filename(&img_zip_path));

                            if !image_by_id.contains_key(&id) {
                                let idx = images.len();
                                images.push(ParsedImage {
                                    id: id.clone(),
                                    bytes: img_bytes,
                                    mime_type: mime.clone(),
                                    filename: Some(img_zip_path.clone()),
                                });
                                image_by_id.insert(id.clone(), idx);
                            }

                            flush_p(blocks, next_idx, &mut cur_p);
                            flush_heading(
                                blocks,
                                next_idx,
                                &mut cur_heading_level,
                                &mut cur_heading_text,
                            );
                            blocks.push(Block::Image {
                                block_index: *next_idx,
                                id,
                                filename: Some(img_zip_path),
                                content_type: Some(mime),
                                alt: alt.map(|s| normalize_ws(&s)).filter(|s| !s.is_empty()),
                                source: SourceSpan::default(),
                            });
                            *next_idx += 1;
                        }
                    }
                    "a" => {
                        let mut href: Option<String> = None;
                        for a in e.attributes().flatten() {
                            if a.key.as_ref().ends_with(b"href") {
                                if let Ok(v) = a.unescape_value() {
                                    href = Some(v.to_string());
                                }
                            }
                        }
                        if let Some(href) = href {
                            in_a = Some(href);
                            a_text.clear();
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let name = String::from_utf8_lossy(e.local_name().as_ref()).to_ascii_lowercase();
                match name.as_str() {
                    "br" => {
                        if let Some(p) = cur_p.as_mut() {
                            p.push('\n');
                        } else if cur_heading_level.is_some() {
                            cur_heading_text.push('\n');
                        } else if let Some(li) = cur_li.as_mut() {
                            li.push('\n');
                        }
                    }
                    "img" => {
                        let mut src: Option<String> = None;
                        let mut alt: Option<String> = None;
                        for a in e.attributes().flatten() {
                            let k = a.key.as_ref();
                            if k.ends_with(b"src") {
                                if let Ok(v) = a.unescape_value() {
                                    src = Some(v.to_string());
                                }
                            } else if k.ends_with(b"alt") {
                                if let Ok(v) = a.unescape_value() {
                                    alt = Some(v.to_string());
                                }
                            }
                        }
                        if let Some(src) = src {
                            let img_zip_path = resolve_path(xhtml_path, &src);
                            let img_bytes = super::read_zip_file(bytes, &img_zip_path)
                                .with_context(|| format!("read image {img_zip_path}"))?;
                            let hash = sha256_hex(&img_bytes);
                            let id = format!("sha256:{hash}");
                            let mime = manifest_href_to_media
                                .get(strip_fragment_and_query(&src).trim())
                                .cloned()
                                .unwrap_or_else(|| mime_from_filename(&img_zip_path));

                            if !image_by_id.contains_key(&id) {
                                let idx = images.len();
                                images.push(ParsedImage {
                                    id: id.clone(),
                                    bytes: img_bytes,
                                    mime_type: mime.clone(),
                                    filename: Some(img_zip_path.clone()),
                                });
                                image_by_id.insert(id.clone(), idx);
                            }

                            flush_p(blocks, next_idx, &mut cur_p);
                            flush_heading(
                                blocks,
                                next_idx,
                                &mut cur_heading_level,
                                &mut cur_heading_text,
                            );
                            blocks.push(Block::Image {
                                block_index: *next_idx,
                                id,
                                filename: Some(img_zip_path),
                                content_type: Some(mime),
                                alt: alt.map(|s| normalize_ws(&s)).filter(|s| !s.is_empty()),
                                source: SourceSpan::default(),
                            });
                            *next_idx += 1;
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(t)) => {
                if let Ok(s) = t.decode() {
                    if in_title {
                        append_text_glue(&mut title_buf, &s);
                    } else if in_a.is_some() {
                        append_text_glue(&mut a_text, &s);
                    } else if let Some(li) = cur_li.as_mut() {
                        append_text_glue(li, &s);
                    } else if cur_heading_level.is_some() {
                        append_text_glue(&mut cur_heading_text, &s);
                    } else if let Some(p) = cur_p.as_mut() {
                        append_text_glue(p, &s);
                    }
                }
            }
            Ok(Event::CData(t)) => {
                if let Ok(s) = t.decode() {
                    if let Some(p) = cur_p.as_mut() {
                        append_text_glue(p, &s);
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = String::from_utf8_lossy(e.local_name().as_ref()).to_ascii_lowercase();
                match name.as_str() {
                    "title" => {
                        in_title = false;
                    }
                    "p" => {
                        flush_p(blocks, next_idx, &mut cur_p);
                    }
                    "h1" | "h2" | "h3" | "h4" | "h5" | "h6" => {
                        flush_heading(
                            blocks,
                            next_idx,
                            &mut cur_heading_level,
                            &mut cur_heading_text,
                        );
                    }
                    "li" => {
                        if let Some(li) = cur_li.take() {
                            let li = normalize_ws(&li);
                            if !li.is_empty() {
                                cur_list_items.push(li);
                            }
                        }
                    }
                    "ul" => {
                        let ordered = list_stack.pop().unwrap_or(false);
                        flush_list(blocks, next_idx, ordered, &mut cur_list_items);
                    }
                    "ol" => {
                        let ordered = list_stack.pop().unwrap_or(true);
                        flush_list(blocks, next_idx, ordered, &mut cur_list_items);
                    }
                    "a" => {
                        if let Some(href) = in_a.take() {
                            let text = normalize_ws(&a_text);
                            if !text.is_empty() {
                                // store as Link block; paragraph glue will keep context.
                                flush_p(blocks, next_idx, &mut cur_p);
                                push_link(blocks, next_idx, href, Some(text));
                            }
                            a_text.clear();
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(anyhow!("parse xhtml {xhtml_path}: {e}")),
        }
        buf.clear();
    }

    let title = {
        let t = normalize_ws(&title_buf);
        if !t.is_empty() {
            t
        } else {
            std::path::Path::new(xhtml_path)
                .file_name()
                .and_then(|s| s.to_str())
                .unwrap_or("chapter")
                .to_string()
        }
    };
    if let Some(Block::Heading { text, .. }) = blocks.get_mut(heading_pos) {
        *text = title;
    }

    Ok(())
}

fn parse_epub_full(bytes: &[u8]) -> Result<ParsedOfficeDocument> {
    let container_xml = super::read_zip_file_utf8(bytes, "META-INF/container.xml")
        .context("read META-INF/container.xml")?;
    let opf_path = parse_container_opf_path(&container_xml)?;
    let opf_xml = super::read_zip_file_utf8(bytes, &opf_path)
        .with_context(|| format!("read OPF {opf_path}"))?;

    let (title, manifest, spine) = parse_opf(&opf_xml)?;

    let mut manifest_href_to_media: std::collections::HashMap<String, String> = Default::default();
    for item in manifest.values() {
        // Use the literal href string for lookups from XHTML.
        manifest_href_to_media.insert(
            strip_fragment_and_query(&item.href).to_string(),
            item.media_type.clone(),
        );
    }

    let mut blocks: Vec<Block> = Vec::new();
    let mut images: Vec<ParsedImage> = Vec::new();
    let mut next_idx: usize = 0;
    let mut image_by_id: std::collections::HashMap<String, usize> = Default::default();

    for idref in spine {
        let Some(item) = manifest.get(&idref) else {
            continue;
        };
        if !item.media_type.contains("xhtml") && !item.media_type.contains("html") {
            continue;
        }
        let xhtml_path = resolve_path(&opf_path, &item.href);
        let xhtml_bytes = super::read_zip_file(bytes, &xhtml_path)
            .with_context(|| format!("read spine item {idref} at {xhtml_path}"))?;
        parse_xhtml_to_blocks(
            &xhtml_bytes,
            &xhtml_path,
            bytes,
            &manifest_href_to_media,
            &mut blocks,
            &mut images,
            &mut next_idx,
            &mut image_by_id,
        )?;
    }

    Ok(ParsedOfficeDocument {
        blocks,
        images,
        metadata_json: serde_json::json!({
            "kind": "epub",
            "title": title,
            "opf_path": opf_path,
        }),
    })
}
