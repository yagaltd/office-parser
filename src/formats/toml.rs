use anyhow::{Context, Result, anyhow};

use super::{ParsedOfficeDocument, config_mdkv};
use crate::{Document, Format};

pub fn parse(bytes: &[u8]) -> crate::Result<Document> {
    let parsed = parse_toml_full(bytes).map_err(|e| crate::Error::from_parse(Format::Toml, e))?;
    Ok(super::finalize(Format::Toml, parsed))
}

fn toml_to_json(v: toml::Value) -> serde_json::Value {
    match v {
        toml::Value::String(s) => serde_json::Value::String(s),
        toml::Value::Integer(i) => serde_json::Value::Number(i.into()),
        toml::Value::Float(f) => serde_json::Number::from_f64(f)
            .map(serde_json::Value::Number)
            .unwrap_or_else(|| serde_json::Value::String(f.to_string())),
        toml::Value::Boolean(b) => serde_json::Value::Bool(b),
        toml::Value::Datetime(dt) => serde_json::Value::String(dt.to_string()),
        toml::Value::Array(a) => {
            serde_json::Value::Array(a.into_iter().map(toml_to_json).collect())
        }
        toml::Value::Table(t) => {
            let mut out = serde_json::Map::new();
            // toml::Table is a BTreeMap; iter is deterministic.
            for (k, v) in t {
                out.insert(k, toml_to_json(v));
            }
            serde_json::Value::Object(out)
        }
    }
}

fn parse_toml_full(bytes: &[u8]) -> Result<ParsedOfficeDocument> {
    let s = std::str::from_utf8(bytes).context("decode toml as utf-8")?;
    let v: toml::Value = toml::from_str(s).map_err(|e| anyhow!("parse toml: {e}"))?;
    let v = toml_to_json(v);
    let blocks = config_mdkv::build_blocks_max_depth_4(&v);
    Ok(ParsedOfficeDocument {
        blocks,
        images: Vec::new(),
        metadata_json: serde_json::json!({
            "kind": "toml",
        }),
    })
}
