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

pub fn parse(bytes: &[u8]) -> crate::Result<Document> {
    let parsed = parse_mmap_full(bytes).map_err(|e| crate::Error::from_parse(Format::Mmap, e))?;
    Ok(super::finalize(Format::Mmap, parsed))
}

fn parse_mmap_full(bytes: &[u8]) -> Result<ParsedOfficeDocument> {
    let xml = super::read_zip_file_utf8(bytes, "Document.xml").context("read mmap Document.xml")?;
    let root = parse_document_xml_root(&xml)?;

    let blocks = blocks_from_root(&root);
    let metadata_json = serde_json::json!({
        "kind": "mmap",
        "title": root.title,
        "mindmap": {
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

fn local_name(tag: &[u8]) -> String {
    let raw = String::from_utf8_lossy(tag);
    if let Some((_, rhs)) = raw.rsplit_once(':') {
        return rhs.to_string();
    }
    raw.to_string()
}

fn parse_document_xml_root(xml: &str) -> Result<MindNode> {
    let mut reader = quick_xml::Reader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(true);

    #[derive(Debug)]
    struct TopicBuilder {
        title: String,
        collapsed: bool,
        children: Vec<MindNode>,
    }

    let mut buf = Vec::new();
    let mut in_one_topic = false;
    let mut topic_stack: Vec<TopicBuilder> = Vec::new();
    let mut in_text = false;
    let mut root: Option<MindNode> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(e)) => {
                let n = local_name(e.name().as_ref());
                if n == "OneTopic" {
                    in_one_topic = true;
                    buf.clear();
                    continue;
                }

                if !in_one_topic {
                    buf.clear();
                    continue;
                }

                if n == "Topic" {
                    topic_stack.push(TopicBuilder {
                        title: String::new(),
                        collapsed: false,
                        children: Vec::new(),
                    });
                } else if n == "Collapsed" {
                    let is_collapsed = e.attributes().flatten().any(|a| {
                        local_name(a.key.as_ref()) == "Collapsed"
                            && a.unescape_value()
                                .map(|v| v.eq_ignore_ascii_case("true"))
                                .unwrap_or(false)
                    });
                    if let Some(top) = topic_stack.last_mut() {
                        top.collapsed = is_collapsed;
                    }
                } else if n == "Text" {
                    if let Some(top) = topic_stack.last_mut() {
                        let plain = e.attributes().flatten().find_map(|a| {
                            if local_name(a.key.as_ref()) == "PlainText" {
                                a.unescape_value().ok().map(|v| v.to_string())
                            } else {
                                None
                            }
                        });
                        if let Some(t) = plain {
                            top.title = t;
                        } else {
                            in_text = true;
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Empty(e)) => {
                if !in_one_topic {
                    buf.clear();
                    continue;
                }
                let n = local_name(e.name().as_ref());
                if n == "Collapsed" {
                    let is_collapsed = e.attributes().flatten().any(|a| {
                        local_name(a.key.as_ref()) == "Collapsed"
                            && a.unescape_value()
                                .map(|v| v.eq_ignore_ascii_case("true"))
                                .unwrap_or(false)
                    });
                    if let Some(top) = topic_stack.last_mut() {
                        top.collapsed = is_collapsed;
                    }
                } else if n == "Text" {
                    if let Some(top) = topic_stack.last_mut() {
                        let plain = e.attributes().flatten().find_map(|a| {
                            if local_name(a.key.as_ref()) == "PlainText" {
                                a.unescape_value().ok().map(|v| v.to_string())
                            } else {
                                None
                            }
                        });
                        if let Some(t) = plain {
                            top.title = t;
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Text(t)) => {
                if in_one_topic && in_text {
                    let s = t.decode().map(|v| v.to_string()).unwrap_or_default();
                    if let Some(top) = topic_stack.last_mut() {
                        top.title.push_str(&s);
                    }
                }
            }
            Ok(quick_xml::events::Event::CData(t)) => {
                if in_one_topic && in_text {
                    let s = t.decode().map(|v| v.to_string()).unwrap_or_default();
                    if let Some(top) = topic_stack.last_mut() {
                        top.title.push_str(&s);
                    }
                }
            }
            Ok(quick_xml::events::Event::End(e)) => {
                let n = local_name(e.name().as_ref());
                if n == "Text" {
                    in_text = false;
                } else if n == "Topic" && in_one_topic {
                    if let Some(top) = topic_stack.pop() {
                        let title = if top.title.trim().is_empty() {
                            "Untitled".to_string()
                        } else {
                            top.title.trim().to_string()
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
                } else if n == "OneTopic" && in_one_topic {
                    in_one_topic = false;
                    if root.is_some() {
                        break;
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(anyhow!(e)).context("parse mmap Document.xml"),
        }
        buf.clear();
    }

    root.ok_or_else(|| anyhow!("mmap Document.xml missing root topic"))
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
    fn mmap_document_xml_parses_topic_tree() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<ap:Map xmlns:ap="http://schemas.mindjet.com/MindManager/Application/2003">
  <ap:OneTopic>
    <ap:Topic>
      <ap:Text PlainText="Weekly Meeting"/>
      <ap:SubTopics>
        <ap:Topic>
          <ap:Text PlainText="Agenda"/>
          <ap:SubTopics>
            <ap:Topic>
              <ap:Text PlainText="Status"/>
              <ap:TopicViewGroup><ap:Collapsed Collapsed="true"/></ap:TopicViewGroup>
            </ap:Topic>
          </ap:SubTopics>
        </ap:Topic>
        <ap:Topic>
          <ap:Text PlainText="Decisions"/>
        </ap:Topic>
      </ap:SubTopics>
    </ap:Topic>
  </ap:OneTopic>
</ap:Map>
"#;

        let bytes = make_zip(&[("Document.xml", document_xml)]);
        let doc = parse(&bytes).expect("parse mmap");

        assert_eq!(doc.metadata.format, Format::Mmap);
        assert_eq!(doc.metadata.title.as_deref(), Some("Weekly Meeting"));

        let md = crate::render::to_markdown(&doc);
        assert!(md.contains("# Weekly Meeting"), "{md}");
        assert!(md.contains("- Agenda"), "{md}");
        assert!(md.contains("  - Status"), "{md}");
        assert!(md.contains("- Decisions"), "{md}");

        let node_count = doc
            .metadata
            .extra
            .get("mindmap")
            .and_then(|m| m.get("node_count"))
            .and_then(|v| v.as_u64());
        assert_eq!(node_count, Some(4));
    }
}
