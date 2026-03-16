use anyhow::{Context, Result, anyhow};

use super::ParsedOfficeDocument;
use crate::document_ast::{Block, ListItem, SourceSpan};
use crate::{Document, Format};

#[derive(Clone, Debug)]
struct MindNode {
    title: String,
    collapsed: bool,
    children: Vec<MindNode>,
}

impl MindNode {
    fn new(title: String, collapsed: bool) -> Self {
        Self {
            title,
            collapsed,
            children: Vec::new(),
        }
    }
}

pub fn parse(bytes: &[u8]) -> crate::Result<Document> {
    let parsed = parse_xmind_full(bytes).map_err(|e| crate::Error::from_parse(Format::Xmind, e))?;
    Ok(super::finalize(Format::Xmind, parsed))
}

fn parse_xmind_full(bytes: &[u8]) -> Result<ParsedOfficeDocument> {
    let (root, sheet_index, source_kind) =
        if let Ok(v) = super::read_zip_file(bytes, "content.json") {
            let value: serde_json::Value =
                serde_json::from_slice(&v).context("parse xmind content.json")?;
            let (node, idx) = parse_content_json_root(&value)?;
            (node, idx, "content.json")
        } else {
            let xml = super::read_zip_file_utf8(bytes, "content.xml")
                .context("read xmind content.xml")?;
            let node = parse_content_xml_root(&xml)?;
            (node, 0usize, "content.xml")
        };

    let blocks = blocks_from_root(&root);
    let metadata_json = serde_json::json!({
        "kind": "xmind",
        "title": root.title,
        "mindmap": {
            "source": source_kind,
            "sheet_index": sheet_index,
            "node_count": node_count(&root),
            "max_depth": max_depth(&root),
            "root": node_to_json(&root),
        }
    });

    Ok(ParsedOfficeDocument {
        blocks,
        images: Vec::new(),
        metadata_json,
    })
}

fn parse_content_json_root(v: &serde_json::Value) -> Result<(MindNode, usize)> {
    let sheets = if let Some(a) = v.as_array() {
        a
    } else {
        return Err(anyhow!(
            "xmind content.json root must be an array of sheets"
        ));
    };

    let (idx, sheet) = sheets
        .iter()
        .enumerate()
        .find(|(_, s)| s.get("rootTopic").is_some() || s.get("topic").is_some())
        .ok_or_else(|| anyhow!("xmind content.json missing sheet root topic"))?;

    let root_topic = sheet
        .get("rootTopic")
        .or_else(|| sheet.get("topic"))
        .ok_or_else(|| anyhow!("xmind content.json sheet missing root topic"))?;

    Ok((node_from_json_topic(root_topic), idx))
}

fn node_from_json_topic(v: &serde_json::Value) -> MindNode {
    let title = v
        .get("title")
        .and_then(|x| x.as_str())
        .map(|s| s.to_string())
        .filter(|s| !s.trim().is_empty())
        .unwrap_or_else(|| "Untitled".to_string());

    let collapsed = v
        .get("branch")
        .and_then(|x| x.as_str())
        .map(|s| s.eq_ignore_ascii_case("folded"))
        .unwrap_or(false);

    let mut node = MindNode::new(title, collapsed);

    if let Some(children_obj) = v.get("children").and_then(|x| x.as_object()) {
        if let Some(attached) = children_obj.get("attached").and_then(|x| x.as_array()) {
            for child in attached {
                node.children.push(node_from_json_topic(child));
            }
        }
    }

    node
}

fn local_name(tag: &[u8]) -> String {
    let raw = String::from_utf8_lossy(tag);
    if let Some((_, rhs)) = raw.rsplit_once(':') {
        return rhs.to_string();
    }
    raw.to_string()
}

fn parse_content_xml_root(xml: &str) -> Result<MindNode> {
    let mut reader = quick_xml::Reader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);

    #[derive(Debug)]
    struct TopicBuilder {
        title: String,
        collapsed: bool,
        children: Vec<MindNode>,
    }

    let mut buf = Vec::new();
    let mut seen_first_sheet = false;
    let mut in_first_sheet = false;
    let mut topic_stack: Vec<TopicBuilder> = Vec::new();
    let mut in_title = false;
    let mut root: Option<MindNode> = None;

    let decode_general_ref = |r: quick_xml::events::BytesRef<'_>| -> Option<String> {
        if let Ok(Some(ch)) = r.resolve_char_ref() {
            return Some(ch.to_string());
        }
        let name = r.decode().ok()?;
        let s = match name.as_ref() {
            "amp" => "&",
            "apos" => "'",
            "quot" => "\"",
            "lt" => "<",
            "gt" => ">",
            _ => return Some(format!("&{};", name)),
        };
        Some(s.to_string())
    };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(e)) => {
                let n = local_name(e.name().as_ref());
                if n == "sheet" {
                    if !seen_first_sheet {
                        seen_first_sheet = true;
                        in_first_sheet = true;
                    }
                    buf.clear();
                    continue;
                }

                if !in_first_sheet {
                    buf.clear();
                    continue;
                }

                if n == "topic" {
                    let collapsed = e.attributes().flatten().any(|a| {
                        local_name(a.key.as_ref()) == "branch"
                            && a.unescape_value()
                                .map(|v| v.eq_ignore_ascii_case("folded"))
                                .unwrap_or(false)
                    });
                    topic_stack.push(TopicBuilder {
                        title: String::new(),
                        collapsed,
                        children: Vec::new(),
                    });
                } else if n == "title" && !topic_stack.is_empty() {
                    in_title = true;
                }
            }
            Ok(quick_xml::events::Event::Text(t)) => {
                if in_first_sheet && in_title {
                    let s = t
                        .decode()
                        .ok()
                        .and_then(|decoded| {
                            quick_xml::escape::unescape(&decoded)
                                .ok()
                                .map(|v| v.into_owned())
                        })
                        .unwrap_or_default();
                    if let Some(cur) = topic_stack.last_mut() {
                        cur.title.push_str(&s);
                    }
                }
            }
            Ok(quick_xml::events::Event::CData(t)) => {
                if in_first_sheet && in_title {
                    let s = t.decode().map(|v| v.to_string()).unwrap_or_default();
                    if let Some(cur) = topic_stack.last_mut() {
                        cur.title.push_str(&s);
                    }
                }
            }
            Ok(quick_xml::events::Event::GeneralRef(r)) => {
                if in_first_sheet && in_title {
                    if let Some(s) = decode_general_ref(r) {
                        if let Some(cur) = topic_stack.last_mut() {
                            cur.title.push_str(&s);
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::End(e)) => {
                let n = local_name(e.name().as_ref());

                if n == "sheet" && in_first_sheet {
                    in_first_sheet = false;
                    if root.is_some() {
                        break;
                    }
                }

                if !seen_first_sheet {
                    buf.clear();
                    continue;
                }

                if n == "title" {
                    in_title = false;
                } else if n == "topic" && in_first_sheet {
                    if let Some(top) = topic_stack.pop() {
                        let title = if top.title.trim().is_empty() {
                            "Untitled".to_string()
                        } else {
                            top.title
                        };
                        let node = MindNode {
                            title,
                            collapsed: top.collapsed,
                            children: top.children,
                        };
                        if let Some(parent) = topic_stack.last_mut() {
                            parent.children.push(node);
                        } else if root.is_none() {
                            root = Some(node);
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(anyhow!(e)).context("parse xmind content.xml"),
        }
        buf.clear();
    }

    root.ok_or_else(|| anyhow!("xmind content.xml missing root topic in first sheet"))
}

fn blocks_from_root(root: &MindNode) -> Vec<Block> {
    let mut blocks = Vec::new();
    let mut block_index = 0usize;

    blocks.push(Block::Heading {
        block_index,
        level: 1,
        text: root.title.clone(),
        source: SourceSpan::default(),
    });
    block_index += 1;

    let mut items: Vec<ListItem> = Vec::new();
    for child in &root.children {
        flatten_list(child, 0, &mut items);
    }

    if !items.is_empty() {
        blocks.push(Block::List {
            block_index,
            ordered: false,
            items,
            source: SourceSpan::default(),
        });
    }

    blocks
}

fn flatten_list(node: &MindNode, level: usize, out: &mut Vec<ListItem>) {
    out.push(ListItem {
        level: level.min(u8::MAX as usize) as u8,
        text: node.title.clone(),
        source: SourceSpan::default(),
    });

    for child in &node.children {
        flatten_list(child, level + 1, out);
    }
}

fn node_to_json(n: &MindNode) -> serde_json::Value {
    serde_json::json!({
        "title": n.title,
        "collapsed": n.collapsed,
        "children": n.children.iter().map(node_to_json).collect::<Vec<_>>()
    })
}

fn node_count(n: &MindNode) -> usize {
    1 + n.children.iter().map(node_count).sum::<usize>()
}

fn max_depth(n: &MindNode) -> usize {
    if n.children.is_empty() {
        0
    } else {
        1 + n.children.iter().map(max_depth).max().unwrap_or(0)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_zip(entries: &[(&str, &str)]) -> Vec<u8> {
        use std::io::Write;

        let mut cur = std::io::Cursor::new(Vec::<u8>::new());
        {
            let mut w = zip::ZipWriter::new(&mut cur);
            let opts = zip::write::FileOptions::default();
            for (name, body) in entries {
                w.start_file(*name, opts).expect("start file");
                w.write_all(body.as_bytes()).expect("write body");
            }
            w.finish().expect("finish zip");
        }
        cur.into_inner()
    }

    #[test]
    fn xmind_content_json_parses_to_heading_and_nested_list() {
        let content_json = r#"[
  {
    "id": "sheet-1",
    "title": "Main",
    "rootTopic": {
      "id": "root",
      "title": "Root",
      "children": {
        "attached": [
          {
            "id": "a",
            "title": "A",
            "children": { "attached": [ { "id": "a1", "title": "A1", "branch": "folded" } ] }
          },
          { "id": "b", "title": "B" }
        ]
      }
    }
  }
]"#;
        let bytes = make_zip(&[("content.json", content_json)]);

        let doc = parse(&bytes).expect("parse xmind content.json");
        assert_eq!(doc.metadata.format, Format::Xmind);
        assert_eq!(doc.metadata.title.as_deref(), Some("Root"));

        let md = crate::render::to_markdown(&doc);
        assert!(md.contains("# Root"), "{md}");
        assert!(md.contains("- A"), "{md}");
        assert!(md.contains("  - A1"), "{md}");
        assert!(md.contains("- B"), "{md}");

        let mindmap = doc.metadata.extra.get("mindmap").expect("mindmap metadata");
        assert_eq!(mindmap.get("node_count").and_then(|v| v.as_u64()), Some(4));
    }

    #[test]
    fn xmind_content_xml_parses_first_sheet() {
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<xmap-content xmlns="urn:xmind:xmap:xmlns:content:2.0" version="2.0">
  <sheet id="s1">
    <topic id="r1">
      <title>Root XML</title>
      <children>
        <topics type="attached">
          <topic id="c1"><title>Child</title></topic>
        </topics>
      </children>
    </topic>
  </sheet>
  <sheet id="s2">
    <topic id="r2"><title>Ignored</title></topic>
  </sheet>
</xmap-content>
"#;
        let bytes = make_zip(&[("content.xml", content_xml)]);

        let doc = parse(&bytes).expect("parse xmind content.xml");
        assert_eq!(doc.metadata.title.as_deref(), Some("Root XML"));
        let md = crate::render::to_markdown(&doc);
        assert!(md.contains("# Root XML"), "{md}");
        assert!(md.contains("- Child"), "{md}");
        assert!(!md.contains("Ignored"), "{md}");
    }

    #[test]
    fn xmind_content_xml_unescapes_entities_in_titles() {
        let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<xmap-content xmlns="urn:xmind:xmap:xmlns:content:2.0" version="2.0">
  <sheet id="s1">
    <topic id="r1">
      <title>Root &amp; Team</title>
      <children>
        <topics type="attached">
          <topic id="c1"><title>Door doesn&apos;t close</title></topic>
          <topic id="c2"><title>Venturi 4&apos;&apos; finished</title></topic>
        </topics>
      </children>
    </topic>
  </sheet>
</xmap-content>
"#;
        let bytes = make_zip(&[("content.xml", content_xml)]);

        let doc = parse(&bytes).expect("parse xmind content.xml");
        let md = crate::render::to_markdown(&doc);
        assert!(md.contains("# Root & Team"), "{md}");
        assert!(md.contains("- Door doesn't close"), "{md}");
        assert!(md.contains("- Venturi 4'' finished"), "{md}");
    }
}
