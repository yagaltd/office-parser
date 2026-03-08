pub mod docx;
pub mod odp;
pub mod odt;
pub mod pptx;
#[cfg(feature = "lopdf_pdf")]
pub mod pdf;
pub mod rtf;

use anyhow::{Context, Result};

#[derive(Clone, Debug)]
pub struct ParsedImage {
    pub bytes: Vec<u8>,
    pub mime_type: String,
    pub filename: Option<String>,
}

#[derive(Clone, Debug)]
pub struct ParsedOfficeDocument {
    pub blocks: Vec<crate::document_ast::Block>,
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

fn read_zip_file_utf8(bytes: &[u8], path: &str) -> Result<String> {
    let v = read_zip_file(bytes, path)?;
    Ok(String::from_utf8(v).with_context(|| format!("decode {path} as utf-8"))?)
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
