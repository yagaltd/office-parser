use anyhow::{Context, Result};

use crate::document_ast::Block;
use crate::{Document, DocumentMetadata, ExtractedImage, Format};

#[derive(Clone, Debug)]
pub struct ParsedImage {
    pub id: String,
    pub bytes: Vec<u8>,
    pub mime_type: String,
    pub filename: Option<String>,
}

#[derive(Clone, Debug)]
pub struct ParsedOfficeDocument {
    pub blocks: Vec<Block>,
    pub images: Vec<ParsedImage>,
    pub metadata_json: serde_json::Value,
}

fn mime_from_filename(name: &str) -> String {
    let lower = name.to_ascii_lowercase();
    if lower.ends_with(".png") {
        return "image/png".to_string();
    }
    if lower.ends_with(".jpg") || lower.ends_with(".jpeg") {
        return "image/jpeg".to_string();
    }
    if lower.ends_with(".webp") {
        return "image/webp".to_string();
    }
    if lower.ends_with(".bmp") {
        return "image/bmp".to_string();
    }
    if lower.ends_with(".tif") || lower.ends_with(".tiff") {
        return "image/tiff".to_string();
    }
    "application/octet-stream".to_string()
}

#[cfg(any(
    feature = "docx",
    feature = "pptx",
    feature = "odt",
    feature = "odp",
    feature = "epub",
    feature = "xlsx",
    feature = "ods_sheet"
))]
fn read_zip_file(bytes: &[u8], path: &str) -> Result<Vec<u8>> {
    use std::io::{Cursor, Read};
    let cur = Cursor::new(bytes);
    let mut zip = zip::ZipArchive::new(cur).context("open zip")?;
    let mut f = zip
        .by_name(path)
        .with_context(|| format!("zip missing entry: {path}"))?;
    let mut out = Vec::with_capacity(f.size() as usize);
    f.read_to_end(&mut out)
        .with_context(|| format!("read zip entry: {path}"))?;
    Ok(out)
}

#[cfg(not(any(
    feature = "docx",
    feature = "pptx",
    feature = "odt",
    feature = "odp",
    feature = "epub",
    feature = "xlsx",
    feature = "ods_sheet"
)))]
fn read_zip_file(_bytes: &[u8], _path: &str) -> Result<Vec<u8>> {
    anyhow::bail!("zip-based formats are disabled")
}

#[cfg(any(
    feature = "docx",
    feature = "pptx",
    feature = "odt",
    feature = "odp",
    feature = "epub",
    feature = "xlsx",
    feature = "ods_sheet"
))]
fn read_zip_file_utf8(bytes: &[u8], path: &str) -> Result<String> {
    let v = read_zip_file(bytes, path)?;
    Ok(String::from_utf8(v).with_context(|| format!("decode {path} as utf-8"))?)
}

#[cfg(not(any(
    feature = "docx",
    feature = "pptx",
    feature = "odt",
    feature = "odp",
    feature = "epub",
    feature = "xlsx",
    feature = "ods_sheet"
)))]
fn read_zip_file_utf8(_bytes: &[u8], _path: &str) -> Result<String> {
    anyhow::bail!("zip-based formats are disabled")
}

fn sha256_hex(bytes: &[u8]) -> String {
    use sha2::Digest;
    let mut h = sha2::Sha256::new();
    h.update(bytes);
    let out = h.finalize();

    let mut s = String::with_capacity(out.len() * 2);
    for b in out {
        const HEX: &[u8; 16] = b"0123456789abcdef";
        s.push(HEX[(b >> 4) as usize] as char);
        s.push(HEX[(b & 0x0f) as usize] as char);
    }
    s
}

#[derive(Clone, Debug, Default)]
struct ImageSlideMeta {
    slide_indexes: Vec<usize>,
    slide_titles: Vec<Option<String>>,
}

fn slide_meta_for_image_ids(
    metadata_json: &serde_json::Value,
    blocks: &[Block],
) -> std::collections::HashMap<String, ImageSlideMeta> {
    let mut out: std::collections::HashMap<String, ImageSlideMeta> = Default::default();

    let Some(slides) = metadata_json.get("slides").and_then(|v| v.as_array()) else {
        return out;
    };

    #[derive(Clone, Debug)]
    struct SlideRange {
        idx: usize,
        title: Option<String>,
        first: usize,
        last: usize,
    }

    let mut ranges: Vec<SlideRange> = Vec::new();
    for s in slides {
        let idx = s.get("index").and_then(|v| v.as_u64()).unwrap_or(0) as usize;
        let title = s
            .get("title")
            .and_then(|v| v.as_str())
            .map(|s| s.to_string());
        let first = s.get("block_first").and_then(|v| v.as_u64()).unwrap_or(0) as usize;
        let last = s
            .get("block_last")
            .and_then(|v| v.as_u64())
            .unwrap_or(first as u64) as usize;
        ranges.push(SlideRange {
            idx,
            title,
            first,
            last,
        });
    }

    for b in blocks {
        let (img_id, bidx) = match b {
            Block::Image {
                id, block_index, ..
            } => (id.as_str(), *block_index),
            _ => continue,
        };

        let mut slide_idx: Option<usize> = None;
        let mut slide_title: Option<String> = None;
        for r in &ranges {
            if bidx >= r.first && bidx <= r.last {
                slide_idx = Some(r.idx);
                slide_title = r.title.clone();
                break;
            }
        }

        let Some(slide_idx) = slide_idx else { continue };
        let e = out.entry(img_id.to_string()).or_default();
        if !e.slide_indexes.contains(&slide_idx) {
            e.slide_indexes.push(slide_idx);
            e.slide_titles.push(slide_title);
        }
    }

    for v in out.values_mut() {
        let mut zipped: Vec<(usize, Option<String>)> = v
            .slide_indexes
            .iter()
            .copied()
            .zip(v.slide_titles.iter().cloned())
            .collect();
        zipped.sort_by_key(|(i, _)| *i);
        v.slide_indexes = zipped.iter().map(|(i, _)| *i).collect();
        v.slide_titles = zipped.into_iter().map(|(_, t)| t).collect();
    }

    out
}

fn finalize(format: Format, parsed: ParsedOfficeDocument) -> Document {
    let slide_meta = match format {
        Format::Pptx | Format::Odp => {
            slide_meta_for_image_ids(&parsed.metadata_json, &parsed.blocks)
        }
        _ => Default::default(),
    };

    let slide_count = parsed
        .metadata_json
        .get("slides")
        .and_then(|v| v.as_array())
        .map(|a| a.len());

    let page_count = parsed
        .metadata_json
        .get("page_count")
        .and_then(|v| v.as_u64())
        .map(|n| n as usize);

    let images = parsed
        .images
        .into_iter()
        .map(|img| {
            let source_ref = if matches!(format, Format::Pptx | Format::Odp) {
                img.filename
                    .as_deref()
                    .and_then(|f| f.strip_prefix("slide_"))
                    .and_then(|f| f.strip_suffix(".png"))
                    .and_then(|n| n.parse::<usize>().ok())
                    .map(|idx| format!("slide:{idx}:snapshot"))
                    .or_else(|| {
                        slide_meta
                            .get(&img.id)
                            .and_then(|m| m.slide_indexes.first().copied())
                            .map(|idx| format!("slide:{idx}"))
                    })
            } else {
                slide_meta
                    .get(&img.id)
                    .and_then(|m| m.slide_indexes.first().copied())
                    .map(|idx| format!("slide:{idx}"))
            };
            ExtractedImage {
                bytes: img.bytes,
                mime_type: img.mime_type,
                filename: img.filename,
                source_ref,
                id: img.id,
            }
        })
        .collect();

    let title = parsed
        .metadata_json
        .get("title")
        .and_then(|v| v.as_str())
        .map(|s| s.trim())
        .filter(|s| !s.is_empty())
        .map(|s| s.to_string())
        .or_else(|| {
            if matches!(format, Format::Pptx | Format::Odp) {
                parsed
                    .metadata_json
                    .get("slides")
                    .and_then(|v| v.as_array())
                    .and_then(|slides| {
                        slides.iter().find_map(|s| {
                            s.get("title")
                                .and_then(|t| t.as_str())
                                .map(|t| t.trim())
                                .filter(|t| !t.is_empty())
                                .map(|t| t.to_string())
                        })
                    })
            } else {
                None
            }
        });

    Document {
        blocks: parsed.blocks,
        images,
        metadata: DocumentMetadata {
            format,
            title,
            page_count,
            slide_count,
            extra: parsed.metadata_json,
        },
    }
}

mod config_mdkv;
mod spreadsheet;

#[cfg(feature = "docx")]
pub mod docx;
#[cfg(not(feature = "docx"))]
pub mod docx {
    pub fn parse(_bytes: &[u8]) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("docx"))
    }
}

#[cfg(feature = "odt")]
pub mod odt;
#[cfg(not(feature = "odt"))]
pub mod odt {
    pub fn parse(_bytes: &[u8]) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("odt"))
    }
}

#[cfg(feature = "rtf")]
pub mod rtf;
#[cfg(not(feature = "rtf"))]
pub mod rtf {
    pub fn parse(_bytes: &[u8]) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("rtf"))
    }
}

#[cfg(feature = "pptx")]
pub mod pptx;
#[cfg(not(feature = "pptx"))]
pub mod pptx {
    pub fn parse(_bytes: &[u8]) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("pptx"))
    }
    pub fn parse_with_options(
        _bytes: &[u8],
        _opts: crate::pptx::ParseOptions,
    ) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("pptx"))
    }
}

#[cfg(feature = "odp")]
pub mod odp;
#[cfg(not(feature = "odp"))]
pub mod odp {
    pub fn parse(_bytes: &[u8]) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("odp"))
    }

    pub fn parse_with_options(
        _bytes: &[u8],
        _opts: crate::odp::ParseOptions,
    ) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("odp"))
    }
}

#[cfg(feature = "pdf")]
pub mod pdf;
#[cfg(not(feature = "pdf"))]
pub mod pdf {
    pub fn parse(_bytes: &[u8]) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("pdf"))
    }
}

#[cfg(feature = "json")]
pub mod json;
#[cfg(not(feature = "json"))]
pub mod json {
    pub fn parse(_bytes: &[u8]) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("json"))
    }
}

#[cfg(feature = "yaml")]
pub mod yaml;
#[cfg(not(feature = "yaml"))]
pub mod yaml {
    pub fn parse(_bytes: &[u8]) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("yaml"))
    }
}

#[cfg(feature = "toml")]
pub mod toml;
#[cfg(not(feature = "toml"))]
pub mod toml {
    pub fn parse(_bytes: &[u8]) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("toml"))
    }
}

#[cfg(feature = "xml")]
pub mod xml;
#[cfg(not(feature = "xml"))]
pub mod xml {
    pub fn parse(_bytes: &[u8]) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("xml"))
    }
}

#[cfg(feature = "epub")]
pub mod epub;
#[cfg(not(feature = "epub"))]
pub mod epub {
    pub fn parse(_bytes: &[u8]) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("epub"))
    }
}

#[cfg(feature = "xlsx")]
pub mod xlsx;
#[cfg(not(feature = "xlsx"))]
pub mod xlsx {
    pub fn parse(_bytes: &[u8]) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("xlsx"))
    }

    pub fn parse_with_options(
        _bytes: &[u8],
        _opts: crate::spreadsheet::ParseOptions,
    ) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("xlsx"))
    }
}

#[cfg(feature = "ods_sheet")]
pub mod ods_sheet;
#[cfg(not(feature = "ods_sheet"))]
pub mod ods_sheet {
    pub fn parse(_bytes: &[u8]) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("ods_sheet"))
    }

    pub fn parse_with_options(
        _bytes: &[u8],
        _opts: crate::spreadsheet::ParseOptions,
    ) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("ods_sheet"))
    }
}

#[cfg(feature = "csv")]
pub mod csv;
#[cfg(not(feature = "csv"))]
pub mod csv {
    pub fn parse(_bytes: &[u8]) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("csv"))
    }

    pub fn parse_with_options(
        _bytes: &[u8],
        _opts: crate::spreadsheet::ParseOptions,
    ) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("csv"))
    }
}

#[cfg(feature = "tsv")]
pub mod tsv;
#[cfg(not(feature = "tsv"))]
pub mod tsv {
    pub fn parse(_bytes: &[u8]) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("tsv"))
    }

    pub fn parse_with_options(
        _bytes: &[u8],
        _opts: crate::spreadsheet::ParseOptions,
    ) -> crate::Result<crate::Document> {
        Err(crate::Error::FeatureDisabled("tsv"))
    }
}

// Submodules access helpers via `super::...`.
