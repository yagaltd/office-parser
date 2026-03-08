use anyhow::{Context, Result, anyhow};

use super::{ParsedOfficeDocument, config_mdkv};
use crate::{Document, Format};

pub fn parse(bytes: &[u8]) -> crate::Result<Document> {
    let parsed = parse_xml_full(bytes).map_err(|e| crate::Error::from_parse(Format::Xml, e))?;
    Ok(super::finalize(Format::Xml, parsed))
}

#[derive(Debug)]
struct BuildElem {
    name: String,
    attrs: Vec<(String, String)>,
    child_order: Vec<String>,
    children: std::collections::HashMap<String, Vec<serde_json::Value>>,
    text: String,
}

fn should_keep_attr(name: &str) -> bool {
    // Namespace declarations are almost always noise for ingestion.
    if name == "xmlns" || name.starts_with("xmlns:") {
        return false;
    }
    true
}

fn looks_like_html(s: &str) -> bool {
    let t = s.trim();
    if t.is_empty() {
        return false;
    }
    // Common in WXR exports.
    t.contains('<')
        && t.contains('>')
        && (t.contains("<p") || t.contains("<div") || t.contains("<h") || t.contains("</"))
}

fn html_fragment_to_markdown_best_effort(s: &str) -> String {
    // WordPress block comments (<!-- wp:... -->) are noise.
    let mut cleaned = String::with_capacity(s.len());
    let mut i = 0;
    let bytes = s.as_bytes();
    while i < bytes.len() {
        if i + 4 <= bytes.len() && &bytes[i..i + 4] == b"<!--" {
            if let Some(end) = s[i + 4..].find("-->") {
                i = i + 4 + end + 3;
                continue;
            }
        }
        cleaned.push(bytes[i] as char);
        i += 1;
    }

    let wrapped = format!("<root>{}</root>", cleaned);
    let mut reader = quick_xml::Reader::from_reader(wrapped.as_bytes());
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    enum ListMode {
        Ul,
        Ol,
    }

    let mut out = String::new();
    let mut list_stack: Vec<ListMode> = Vec::new();
    let mut ol_index_stack: Vec<usize> = Vec::new();
    let mut in_a: Option<String> = None;
    let mut in_blockquote = 0usize;

    let ensure_blank_line = |out: &mut String| {
        let t = out.trim_end_matches([' ', '\t']);
        let nl = t.chars().rev().take(2).filter(|c| *c == '\n').count();
        if nl < 2 {
            out.push('\n');
            out.push('\n');
        }
    };

    let push_text = |out: &mut String, t: &str, in_blockquote: usize| {
        let mut s = t.replace("\r", "");
        // Collapse runs of whitespace but preserve newlines.
        s = s
            .split('\n')
            .map(|line| line.split_whitespace().collect::<Vec<_>>().join(" "))
            .collect::<Vec<_>>()
            .join("\n");
        if s.trim().is_empty() {
            return;
        }
        if in_blockquote > 0 {
            for (i, line) in s.lines().enumerate() {
                if i > 0 {
                    out.push('\n');
                }
                out.push_str("> ");
                out.push_str(line);
            }
        } else {
            // Text nodes often split around inline tags (<b>, <a>, ...). Re-insert a space when
            // two adjacent chunks would otherwise glue together.
            let last = out.chars().rev().find(|c| !c.is_whitespace());
            let first = s.chars().find(|c| !c.is_whitespace());
            if matches!(last, Some(c) if c.is_ascii_alphanumeric())
                && matches!(first, Some(c) if c.is_ascii_alphanumeric())
                && !out.ends_with([' ', '\t', '\n'])
                && !s.starts_with([' ', '\t', '\n'])
            {
                out.push(' ');
            }
            out.push_str(&s);
        }
    };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(e)) => {
                let name = String::from_utf8_lossy(e.local_name().as_ref()).to_ascii_lowercase();
                match name.as_str() {
                    "p" | "div" => ensure_blank_line(&mut out),
                    "br" => out.push('\n'),
                    "blockquote" => {
                        ensure_blank_line(&mut out);
                        in_blockquote += 1;
                    }
                    "ul" => list_stack.push(ListMode::Ul),
                    "ol" => {
                        list_stack.push(ListMode::Ol);
                        ol_index_stack.push(0);
                    }
                    "li" => {
                        ensure_blank_line(&mut out);
                        match list_stack.last().copied() {
                            Some(ListMode::Ol) => {
                                if let Some(last) = ol_index_stack.last_mut() {
                                    *last += 1;
                                    out.push_str(&format!("{}. ", *last));
                                } else {
                                    out.push_str("1. ");
                                }
                            }
                            _ => out.push_str("- "),
                        }
                    }
                    "h1" | "h2" | "h3" | "h4" | "h5" | "h6" => {
                        ensure_blank_line(&mut out);
                        let lvl = name[1..].parse::<usize>().unwrap_or(2).clamp(1, 6);
                        out.push_str(&"#".repeat(lvl));
                        out.push(' ');
                    }
                    "a" => {
                        let mut href: Option<String> = None;
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"href" {
                                if let Ok(v) = a.unescape_value() {
                                    href = Some(v.to_string());
                                }
                            }
                        }
                        if let Some(href) = href {
                            out.push('[');
                            in_a = Some(href);
                        }
                    }
                    "script" | "style" => {
                        // Best-effort: ignore content; we'll just drop text by not handling it specially.
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Empty(e)) => {
                let name = String::from_utf8_lossy(e.local_name().as_ref()).to_ascii_lowercase();
                match name.as_str() {
                    "br" => out.push('\n'),
                    "hr" => {
                        ensure_blank_line(&mut out);
                        out.push_str("---\n\n");
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::End(e)) => {
                let name = String::from_utf8_lossy(e.local_name().as_ref()).to_ascii_lowercase();
                match name.as_str() {
                    "blockquote" => {
                        in_blockquote = in_blockquote.saturating_sub(1);
                        ensure_blank_line(&mut out);
                    }
                    "ul" => {
                        let _ = list_stack.pop();
                    }
                    "ol" => {
                        let _ = list_stack.pop();
                        let _ = ol_index_stack.pop();
                    }
                    "a" => {
                        if let Some(href) = in_a.take() {
                            out.push_str("](");
                            out.push_str(&href);
                            out.push(')');
                        }
                    }
                    "p" | "div" | "li" | "h1" | "h2" | "h3" | "h4" | "h5" | "h6" => {
                        ensure_blank_line(&mut out);
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Text(t)) => {
                if let Ok(s) = t.decode() {
                    push_text(&mut out, &s, in_blockquote);
                }
            }
            Ok(quick_xml::events::Event::CData(t)) => {
                if let Ok(s) = t.decode() {
                    push_text(&mut out, &s, in_blockquote);
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => {
                // Fallback: strip tags crudely.
                return cleaned
                    .replace('<', " ")
                    .replace('>', " ")
                    .split_whitespace()
                    .collect::<Vec<_>>()
                    .join(" ");
            }
            _ => {}
        }
        buf.clear();
    }

    out.trim().to_string()
}

fn finalize_elem(e: BuildElem) -> serde_json::Value {
    let mut map = serde_json::Map::new();

    for (k, v) in e.attrs {
        if should_keep_attr(&k) {
            map.insert(format!("@{k}"), serde_json::Value::String(v));
        }
    }

    let mut text = e.text;
    if looks_like_html(&text) {
        text = html_fragment_to_markdown_best_effort(&text);
    }
    let text = text.trim_matches(|c: char| c.is_whitespace()).to_string();

    // Children
    for k in e.child_order {
        let Some(items) = e.children.get(&k) else {
            continue;
        };
        if items.len() == 1 {
            map.insert(k, items[0].clone());
        } else {
            map.insert(k, serde_json::Value::Array(items.clone()));
        }
    }

    // Leaf: just return scalar text.
    if map.is_empty() {
        return serde_json::Value::String(text);
    }
    if !text.is_empty() {
        // Prefer text first (after attrs), before child keys.
        let mut with_text = serde_json::Map::new();

        // attrs
        for (k, v) in map.iter() {
            if k.starts_with('@') {
                with_text.insert(k.clone(), v.clone());
            }
        }

        with_text.insert("#text".to_string(), serde_json::Value::String(text));

        // non-attrs (children)
        for (k, v) in map.into_iter() {
            if !k.starts_with('@') {
                with_text.insert(k, v);
            }
        }

        return serde_json::Value::Object(with_text);
    }

    serde_json::Value::Object(map)
}

fn parse_xml_full(bytes: &[u8]) -> Result<ParsedOfficeDocument> {
    // `quick-xml` expects UTF-8; WXR exports are usually UTF-8.
    let mut reader = quick_xml::Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut stack: Vec<BuildElem> = Vec::new();
    let mut root: Option<serde_json::Value> = None;
    let mut root_name: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(e)) => {
                let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                let mut attrs: Vec<(String, String)> = Vec::new();
                for a in e.attributes().flatten() {
                    let key = String::from_utf8_lossy(a.key.as_ref()).to_string();
                    if let Ok(v) = a.unescape_value() {
                        attrs.push((key, v.to_string()));
                    }
                }
                if stack.is_empty() {
                    root_name = Some(name.clone());
                }
                stack.push(BuildElem {
                    name,
                    attrs,
                    child_order: Vec::new(),
                    children: Default::default(),
                    text: String::new(),
                });
            }
            Ok(quick_xml::events::Event::Empty(e)) => {
                let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                let mut attrs: Vec<(String, String)> = Vec::new();
                for a in e.attributes().flatten() {
                    let key = String::from_utf8_lossy(a.key.as_ref()).to_string();
                    if let Ok(v) = a.unescape_value() {
                        attrs.push((key, v.to_string()));
                    }
                }

                let v = finalize_elem(BuildElem {
                    name: name.clone(),
                    attrs,
                    child_order: Vec::new(),
                    children: Default::default(),
                    text: String::new(),
                });

                if let Some(parent) = stack.last_mut() {
                    if !parent.children.contains_key(&name) {
                        parent.child_order.push(name.clone());
                    }
                    parent.children.entry(name).or_default().push(v);
                } else {
                    root_name = Some(name.clone());
                    root = Some(v);
                }
            }
            Ok(quick_xml::events::Event::End(_e)) => {
                let Some(done) = stack.pop() else {
                    break;
                };
                let name = done.name.clone();
                let v = finalize_elem(done);
                if let Some(parent) = stack.last_mut() {
                    if !parent.children.contains_key(&name) {
                        parent.child_order.push(name.clone());
                    }
                    parent.children.entry(name).or_default().push(v);
                } else {
                    root = Some(v);
                }
            }
            Ok(quick_xml::events::Event::Text(t)) => {
                if let Some(cur) = stack.last_mut() {
                    if let Ok(s) = t.decode() {
                        cur.text.push_str(&s);
                    }
                }
            }
            Ok(quick_xml::events::Event::CData(t)) => {
                if let Some(cur) = stack.last_mut() {
                    if let Ok(s) = t.decode() {
                        cur.text.push_str(&s);
                    }
                }
            }
            Ok(quick_xml::events::Event::Comment(_)) => {}
            Ok(quick_xml::events::Event::Decl(_)) => {}
            Ok(quick_xml::events::Event::PI(_)) => {}
            Ok(quick_xml::events::Event::DocType(_)) => {}
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(anyhow!("parse xml: {e}")),
        }
        buf.clear();
    }

    let root = root.context("missing root element")?;
    let root_name = root_name.unwrap_or_else(|| "xml".to_string());

    let mut outer = serde_json::Map::new();
    outer.insert(root_name, root);
    let v = serde_json::Value::Object(outer);
    let blocks = config_mdkv::build_blocks_max_depth_4(&v);

    Ok(ParsedOfficeDocument {
        blocks,
        images: Vec::new(),
        metadata_json: serde_json::json!({
            "kind": "xml",
        }),
    })
}
