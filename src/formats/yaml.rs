use anyhow::{Context, Result, anyhow};

use super::{ParsedOfficeDocument, config_mdkv};
use crate::{Document, Format};

pub fn parse(bytes: &[u8]) -> crate::Result<Document> {
    let parsed = parse_yaml_full(bytes).map_err(|e| crate::Error::from_parse(Format::Yaml, e))?;
    Ok(super::finalize(Format::Yaml, parsed))
}

fn yaml_to_json(v: serde_yaml::Value) -> serde_json::Value {
    match v {
        serde_yaml::Value::Null => serde_json::Value::Null,
        serde_yaml::Value::Bool(b) => serde_json::Value::Bool(b),
        serde_yaml::Value::Number(n) => {
            if let Some(i) = n.as_i64() {
                serde_json::Value::Number(i.into())
            } else if let Some(u) = n.as_u64() {
                serde_json::Value::Number(u.into())
            } else if let Some(f) = n.as_f64() {
                serde_json::Number::from_f64(f)
                    .map(serde_json::Value::Number)
                    .unwrap_or(serde_json::Value::String(f.to_string()))
            } else {
                serde_json::Value::String(n.to_string())
            }
        }
        serde_yaml::Value::String(s) => serde_json::Value::String(s),
        serde_yaml::Value::Sequence(seq) => {
            serde_json::Value::Array(seq.into_iter().map(yaml_to_json).collect())
        }
        serde_yaml::Value::Mapping(map) => {
            let mut out = serde_json::Map::new();
            for (k, v) in map {
                let key = match k {
                    serde_yaml::Value::String(s) => s,
                    other => serde_yaml::to_string(&other).unwrap_or_else(|_| "<key>".to_string()),
                };
                out.insert(key.trim().to_string(), yaml_to_json(v));
            }
            serde_json::Value::Object(out)
        }
        serde_yaml::Value::Tagged(t) => yaml_to_json(t.value),
    }
}

fn parse_yaml_full(bytes: &[u8]) -> Result<ParsedOfficeDocument> {
    let s = std::str::from_utf8(bytes).context("decode yaml as utf-8")?;
    let v: serde_yaml::Value = serde_yaml::from_str(s).map_err(|e| anyhow!("parse yaml: {e}"))?;
    let v = yaml_to_json(v);
    let blocks = config_mdkv::build_blocks_max_depth_4(&v);
    Ok(ParsedOfficeDocument {
        blocks,
        images: Vec::new(),
        metadata_json: serde_json::json!({
            "kind": "yaml",
        }),
    })
}
