use anyhow::{Context, Result};

use super::{ParsedOfficeDocument, config_mdkv};
use crate::{Document, Format};

pub fn parse(bytes: &[u8]) -> crate::Result<Document> {
    let parsed = parse_json_full(bytes).map_err(|e| crate::Error::from_parse(Format::Json, e))?;
    Ok(super::finalize(Format::Json, parsed))
}

fn parse_json_full(bytes: &[u8]) -> Result<ParsedOfficeDocument> {
    let v: serde_json::Value = serde_json::from_slice(bytes).context("parse json")?;
    let blocks = config_mdkv::build_blocks_max_depth_4(&v);
    Ok(ParsedOfficeDocument {
        blocks,
        images: Vec::new(),
        metadata_json: serde_json::json!({
            "kind": "json",
        }),
    })
}
