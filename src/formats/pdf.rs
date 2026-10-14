use std::collections::{BTreeMap, HashMap};

use anyhow::{Context, Result};

use lopdf::Object;

use crate::document_ast::{Block, Cell, LinkKind, ListItem, SourceSpan};
use crate::PdfTextQuality;

use super::{ParsedImage, sha256_hex};

fn cleanup_pdf_text(s: &str) -> String {
    fn is_hyphen(ch: char) -> bool {
        matches!(ch, '-' | '‐' | '‑' | '‒' | '–' | '—')
    }

    // 1) Dehyphenate word breaks across line wraps: "hyphen-\nated" => "hyphenated".
    let mut tmp = String::with_capacity(s.len());
    let mut it = s.chars().peekable();
    while let Some(ch) = it.next() {
        if is_hyphen(ch) {
            let mut look = it.clone();
            let mut consumed = 0usize;
            // allow optional \r before \n
            let mut saw_break = false;
            if let Some('\r') = look.peek().copied() {
                look.next();
                consumed += 1;
            }
            if let Some('\n') = look.peek().copied() {
                look.next();
                consumed += 1;
                saw_break = true;
            }

            // Some PDFs preserve indentation/spaces at the start of wrapped lines.
            if saw_break {
                while let Some(n) = look.peek().copied() {
                    if n == ' ' || n == '\t' {
                        look.next();
                        consumed += 1;
                    } else {
                        break;
                    }
                }
            }

            if saw_break {
                // If next is a lowercase letter, treat as hyphenation.
                if let Some(nxt) = look.peek().copied() {
                    if nxt.is_alphabetic() && nxt.is_lowercase() {
                        for _ in 0..consumed {
                            it.next();
                        }
                        continue; // drop hyphen + newline
                    }
                }
            }
        }
        tmp.push(ch);
    }

    // 2) Collapse all whitespace (including newlines) into single spaces.
    let mut ws = String::with_capacity(tmp.len());
    let mut last_space = false;
    for ch in tmp.chars() {
        let is_space = ch.is_whitespace();
        if is_space {
            if !last_space {
                ws.push(' ');
                last_space = true;
            }
        } else {
            ws.push(ch);
            last_space = false;
        }
    }

    // 3) Fix common PDF artifact: glued year + citation digits (e.g. 203070 => "2030 70").
    // Also normalize YYYYMM -> YYYY-MM and YYYYMMDD -> YYYY-MM-DD when plausible.
    fn year_ok(y: u32) -> bool {
        (1900..=2100).contains(&y)
    }

    let mut out = String::with_capacity(ws.len());
    let mut chars = ws.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch.is_ascii_digit() {
            let mut digits = String::new();
            digits.push(ch);
            while let Some(n) = chars.peek().copied() {
                if !n.is_ascii_digit() {
                    break;
                }
                digits.push(n);
                chars.next();
            }

            if digits.len() >= 6 {
                let year = digits.get(0..4).and_then(|x| x.parse::<u32>().ok());
                if let Some(y) = year.filter(|y| year_ok(*y)) {
                    if digits.len() == 6 {
                        let mm = digits
                            .get(4..6)
                            .and_then(|x| x.parse::<u32>().ok())
                            .unwrap_or(0);
                        if (1..=12).contains(&mm) {
                            out.push_str(&format!("{y}-{mm:02}"));
                            continue;
                        }
                    }
                    if digits.len() == 8 {
                        let mm = digits
                            .get(4..6)
                            .and_then(|x| x.parse::<u32>().ok())
                            .unwrap_or(0);
                        let dd = digits
                            .get(6..8)
                            .and_then(|x| x.parse::<u32>().ok())
                            .unwrap_or(0);
                        if (1..=12).contains(&mm) && (1..=31).contains(&dd) {
                            out.push_str(&format!("{y}-{mm:02}-{dd:02}"));
                            continue;
                        }
                    }

                    // Split YYYY + suffix for citation-like glue.
                    let suffix = &digits[4..];
                    if (2..=5).contains(&suffix.len()) {
                        out.push_str(&digits[0..4]);
                        out.push(' ');
                        out.push_str(suffix);
                        continue;
                    }
                }
            }

            out.push_str(&digits);
            continue;
        }

        out.push(ch);
    }

    out.trim().to_string()
}

fn cleanup_pdf_blocks(blocks: &mut [Block]) {
    for b in blocks {
        match b {
            Block::Heading { text, .. } | Block::Paragraph { text, .. } => {
                *text = cleanup_pdf_text(text);
            }
            Block::Note { text, .. } => {
                *text = cleanup_pdf_text(text);
            }
            Block::List { items, .. } => {
                for it in items {
                    it.text = cleanup_pdf_text(&it.text);
                }
            }
            Block::Table { rows, .. } => {
                for r in rows {
                    for c in r {
                        c.text = cleanup_pdf_text(&c.text);
                    }
                }
            }
            Block::Image { alt, .. } => {
                if let Some(a) = alt.as_mut() {
                    *a = cleanup_pdf_text(a);
                }
            }
            Block::Link { text, .. } => {
                if let Some(t) = text.as_mut() {
                    *t = cleanup_pdf_text(t);
                }
            }
        }
    }
}

#[derive(Clone, Debug)]
pub struct PdfPageBlockRange {
    pub page_number: u32,
    pub block_start: usize,
    pub block_end: usize, // exclusive
}

#[derive(Clone, Debug)]
pub struct ParsedPdfDocument {
    pub blocks: Vec<Block>,
    pub images: Vec<ParsedImage>,
    pub metadata_json: serde_json::Value,
    pub page_block_ranges: Vec<PdfPageBlockRange>,
}

#[derive(Clone, Debug)]
struct TextAtom {
    seq: u32,
    x: f32,
    y: f32,
    font_size: f32,
    text: String,
}

#[derive(Clone, Debug, Default)]
struct PdfTextState {
    in_text: bool,
    leading: f32,
    font_alias: Option<Vec<u8>>,
    font_size: f32,

    // Graphics/text matrices (2D affine). We only use them to approximate reading-order.
    ctm: Matrix2D,
    tm: Matrix2D,
}

#[derive(Clone, Copy, Debug)]
struct Matrix2D {
    a: f32,
    b: f32,
    c: f32,
    d: f32,
    e: f32,
    f: f32,
}

impl Default for Matrix2D {
    fn default() -> Self {
        Self::identity()
    }
}

impl Matrix2D {
    fn identity() -> Self {
        Self {
            a: 1.0,
            b: 0.0,
            c: 0.0,
            d: 1.0,
            e: 0.0,
            f: 0.0,
        }
    }

    fn from_components(a: f32, b: f32, c: f32, d: f32, e: f32, f: f32) -> Self {
        Self { a, b, c, d, e, f }
    }

    fn multiply(&self, other: &Self) -> Self {
        // Same as mcat's pdf_state: self * other
        Self {
            a: self.a * other.a + self.c * other.b,
            b: self.b * other.a + self.d * other.b,
            c: self.a * other.c + self.c * other.d,
            d: self.b * other.c + self.d * other.d,
            e: self.a * other.e + self.c * other.f + self.e,
            f: self.b * other.e + self.d * other.f + self.f,
        }
    }

    fn apply_to_origin(&self) -> (f32, f32) {
        (self.e, self.f)
    }

    fn scale_hint(&self) -> f32 {
        // Approximate scalar scale for font-size projection (rotation-safe).
        let sx = (self.a * self.a + self.b * self.b).sqrt();
        let sy = (self.c * self.c + self.d * self.d).sqrt();
        sx.max(sy).max(1.0)
    }
}

impl PdfTextState {
    fn bt(&mut self) {
        self.in_text = true;
        self.tm = Matrix2D::identity();
    }

    fn et(&mut self) {
        self.in_text = false;
    }

    fn current_position(&self) -> (f32, f32) {
        // Best-effort: apply text matrix in current graphics state.
        let combined = self.ctm.multiply(&self.tm);
        combined.apply_to_origin()
    }

    fn effective_font_size(&self) -> f32 {
        (self.font_size.abs() * self.tm.scale_hint()).max(1.0)
    }
}

fn obj_to_f32(o: &Object) -> Option<f32> {
    match o {
        Object::Integer(i) => Some(*i as f32),
        Object::Real(f) => Some(*f),
        _ => None,
    }
}

fn decode_pdf_string(bytes: &[u8]) -> String {
    // PDF string objects can be UTF-16 with BOM, but in the wild you'll also find UTF-16BE/LE
    // without a BOM (especially for metadata-like strings).
    if bytes.starts_with(b"\xFE\xFF") {
        let mut u16s = Vec::with_capacity((bytes.len().saturating_sub(2) + 1) / 2);
        let mut i = 2usize;
        while i + 1 < bytes.len() {
            u16s.push(u16::from_be_bytes([bytes[i], bytes[i + 1]]));
            i += 2;
        }
        return String::from_utf16_lossy(&u16s);
    }
    if bytes.starts_with(b"\xFF\xFE") {
        let mut u16s = Vec::with_capacity((bytes.len().saturating_sub(2) + 1) / 2);
        let mut i = 2usize;
        while i + 1 < bytes.len() {
            u16s.push(u16::from_le_bytes([bytes[i], bytes[i + 1]]));
            i += 2;
        }
        return String::from_utf16_lossy(&u16s);
    }

    // Heuristic: UTF-16 without BOM (common pattern: 0x00 xx 0x00 yy ...).
    if bytes.len() >= 4 && bytes.len() % 2 == 0 {
        let pairs = bytes.len() / 2;
        let mut even_nul = 0usize;
        let mut odd_nul = 0usize;
        for i in 0..pairs {
            if bytes[i * 2] == 0 {
                even_nul += 1;
            }
            if bytes[i * 2 + 1] == 0 {
                odd_nul += 1;
            }
        }

        let thr = (pairs * 7) / 10; // 70%
        if even_nul >= thr || odd_nul >= thr {
            let mut u16s = Vec::with_capacity(pairs);
            let mut i = 0usize;
            while i + 1 < bytes.len() {
                let u = if even_nul >= thr {
                    u16::from_be_bytes([bytes[i], bytes[i + 1]])
                } else {
                    u16::from_le_bytes([bytes[i], bytes[i + 1]])
                };
                u16s.push(u);
                i += 2;
            }
            return String::from_utf16_lossy(&u16s);
        }
    }

    if let Ok(s) = std::str::from_utf8(bytes) {
        return s.to_string();
    }

    // PDFDocEncoding is not UTF-8; for Tier-B, fall back to latin-1-ish.
    bytes.iter().map(|b| *b as char).collect()
}

fn collect_fonts_from_resources(
    doc: &lopdf::Document,
    resources: &lopdf::Dictionary,
) -> HashMap<Vec<u8>, lopdf::Dictionary> {
    let mut out = HashMap::new();

    let Ok(font_obj) = resources.get(b"Font") else {
        return out;
    };
    let dict: lopdf::Dictionary = match font_obj {
        Object::Dictionary(d) => d.clone(),
        Object::Reference(oid) => match doc.get_object(*oid).ok() {
            Some(Object::Dictionary(d)) => d.clone(),
            _ => return out,
        },
        _ => return out,
    };

    for (k, v) in dict.iter() {
        let fd = match v {
            Object::Dictionary(d) => Some(d.clone()),
            Object::Reference(oid) => match doc.get_object(*oid).ok() {
                Some(Object::Dictionary(d)) => Some(d.clone()),
                _ => None,
            },
            _ => None,
        };
        if let Some(fd) = fd {
            out.insert(k.clone(), fd);
        }
    }

    out
}

fn decode_font_string(
    encodings: &HashMap<Vec<u8>, lopdf::Encoding>,
    font_alias: Option<&[u8]>,
    bytes: &[u8],
) -> String {
    fn strip_nul(s: &str) -> String {
        if !s.contains('\u{0000}') {
            return s.to_string();
        }
        s.chars().filter(|c| *c != '\u{0000}').collect()
    }

    if let Some(alias) = font_alias {
        if let Some(enc) = encodings.get(alias) {
            if let Ok(s) = lopdf::Document::decode_text(enc, bytes) {
                // Some PDFs (esp. multi-byte / CID-ish) can yield strings with interleaved NULs.
                // Prefer a NUL-stripped version, and if the decoded output is mostly NUL,
                // fall back to direct string decoding.
                let s_stripped = strip_nul(&s);
                if s_stripped.is_empty() {
                    return strip_nul(&decode_pdf_string(bytes));
                }
                let nul_ratio = (s.chars().filter(|c| *c == '\u{0000}').count() as f32)
                    / (s.chars().count().max(1) as f32);
                if nul_ratio >= 0.20 {
                    let alt = strip_nul(&decode_pdf_string(bytes));
                    if alt.chars().count() >= s_stripped.chars().count() {
                        return alt;
                    }
                }
                return s_stripped;
            }
        }
    }

    // Best-effort fallback when the PDF doesn't provide a usable font encoding.
    strip_nul(&decode_pdf_string(bytes))
}

fn decode_text_object(
    encodings: &HashMap<Vec<u8>, lopdf::Encoding>,
    font_alias: Option<&[u8]>,
    o: &Object,
) -> Option<String> {
    match o {
        Object::String(bytes, _) => Some(decode_font_string(encodings, font_alias, bytes)),
        _ => None,
    }
}

fn normalize_atom_text(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut last_space = false;
    for ch in s.chars() {
        if ch == '\u{0000}' {
            continue;
        }
        let is_space = ch.is_whitespace() && ch != '\n';
        if is_space {
            if !last_space {
                out.push(' ');
                last_space = true;
            }
            continue;
        }
        out.push(ch);
        last_space = false;
    }
    out.trim().to_string()
}

fn estimate_text_advance(text: &str, font_size: f32) -> f32 {
    if !font_size.is_finite() || font_size <= 0.0 {
        return 0.0;
    }
    let mut units = 0.0f32;
    for ch in text.chars() {
        if ch == ' ' {
            units += 0.33;
        } else if ch.is_ascii_punctuation() {
            units += 0.35;
        } else {
            units += 0.55;
        }
    }
    (units * font_size).max(0.0)
}
fn is_filter_name(n: &[u8], want: &str) -> bool {
    if let Ok(s) = std::str::from_utf8(n) {
        return s.trim() == want;
    }
    false
}

fn stream_filter_contains(stream: &lopdf::Stream, want: &str) -> bool {
    let Ok(f) = stream.dict.get(b"Filter") else {
        return false;
    };
    match f {
        Object::Name(n) => is_filter_name(n, want),
        Object::Array(arr) => arr.iter().any(|o| match o {
            Object::Name(n) => is_filter_name(n, want),
            _ => false,
        }),
        _ => false,
    }
}

fn extract_xobject_image(
    doc: &lopdf::Document,
    stream: &lopdf::Stream,
    images: &mut Vec<ParsedImage>,
    image_by_hash: &mut HashMap<String, usize>,
) -> Option<(String, String)> {
    fn dict_get_u32(dict: &lopdf::Dictionary, key: &[u8]) -> Option<u32> {
        match dict.get(key).ok()? {
            Object::Integer(i) => (*i).try_into().ok(),
            Object::Real(f) => {
                if *f >= 0.0 {
                    Some(*f as u32)
                } else {
                    None
                }
            }
            _ => None,
        }
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    enum Cs {
        Rgb,
        Gray,
    }

    fn parse_colorspace(o: &Object) -> Option<Cs> {
        match o {
            Object::Name(n) => match n.as_slice() {
                b"DeviceRGB" => Some(Cs::Rgb),
                b"DeviceGray" => Some(Cs::Gray),
                _ => None,
            },
            Object::Array(arr) => {
                // Commonly: [/DeviceRGB] or [/ICCBased ...]
                let first = arr.first()?;
                parse_colorspace(first)
            }
            _ => None,
        }
    }

    fn encode_png(width: u32, height: u32, cs: Cs, raw: &[u8]) -> Result<Vec<u8>> {
        use png::{BitDepth, ColorType, Encoder};

        let mut out = Vec::new();
        let mut enc = Encoder::new(&mut out, width, height);
        match cs {
            Cs::Rgb => enc.set_color(ColorType::Rgb),
            Cs::Gray => enc.set_color(ColorType::Grayscale),
        }
        enc.set_depth(BitDepth::Eight);

        {
            let mut w = enc.write_header().context("png write header")?;
            w.write_image_data(raw).context("png write image data")?;
        }
        Ok(out)
    }

    let (mime, bytes) = if stream_filter_contains(stream, "DCTDecode") {
        ("image/jpeg".to_string(), stream.content.clone())
    } else if stream_filter_contains(stream, "JPXDecode") {
        ("image/jp2".to_string(), stream.content.clone())
    } else {
        // Best-effort decode for common raw pixel images (Flate/LZW/ASCII85, etc.).
        let width = dict_get_u32(&stream.dict, b"Width")?;
        let height = dict_get_u32(&stream.dict, b"Height")?;
        let bpc = dict_get_u32(&stream.dict, b"BitsPerComponent").unwrap_or(8);
        if bpc != 8 {
            return None;
        }

        let Ok(cs_obj) = stream.dict.get(b"ColorSpace") else {
            return None;
        };
        let cs = parse_colorspace(cs_obj)?;

        let raw = stream.decompressed_content().ok()?;
        let want_len = match cs {
            Cs::Rgb => (width as usize)
                .saturating_mul(height as usize)
                .saturating_mul(3),
            Cs::Gray => (width as usize).saturating_mul(height as usize),
        };
        if raw.len() != want_len {
            return None;
        }

        let png = encode_png(width, height, cs, &raw).ok()?;
        ("image/png".to_string(), png)
    };

    if bytes.is_empty() {
        return None;
    }

    // Ensure the image bytes are the raw encoded stream content (not the decoded pixels).
    // Some PDFs might have their stream objects stored indirectly; be conservative.
    let _ = doc; // reserved for future decoding enhancements

    let hash = sha256_hex(&bytes);
    let _idx = *image_by_hash.entry(hash.clone()).or_insert_with(|| {
        images.push(ParsedImage {
            id: format!("sha256:{hash}"),
            bytes,
            mime_type: mime.clone(),
            filename: None,
        });
        images.len() - 1
    });

    Some((format!("sha256:{hash}"), mime))
}

fn resources_get_xobject_stream(
    doc: &lopdf::Document,
    resources: &lopdf::Dictionary,
    name: &[u8],
) -> Result<Option<lopdf::Stream>> {
    let Ok(xobj) = resources.get(b"XObject") else {
        return Ok(None);
    };

    let dict: lopdf::Dictionary = match xobj {
        Object::Dictionary(d) => d.clone(),
        Object::Reference(oid) => {
            let o = doc.get_object(*oid).context("resolve /XObject")?;
            match o {
                Object::Dictionary(d) => d.clone(),
                _ => return Ok(None),
            }
        }
        _ => return Ok(None),
    };

    let Ok(obj) = dict.get(name) else {
        return Ok(None);
    };
    let oid = match obj {
        Object::Reference(oid) => *oid,
        _ => return Ok(None),
    };
    let o = doc.get_object(oid).context("resolve XObject")?;
    let stream = match o {
        Object::Stream(s) => s,
        _ => return Ok(None),
    };
    Ok(Some(stream.clone()))
}

fn stream_subtype(stream: &lopdf::Stream) -> Option<Vec<u8>> {
    let o = stream.dict.get(b"Subtype").ok()?;
    match o {
        Object::Name(n) => Some(n.clone()),
        _ => None,
    }
}

fn resolve_dict<'a>(doc: &'a lopdf::Document, o: &'a Object) -> Option<lopdf::Dictionary> {
    match o {
        Object::Dictionary(d) => Some(d.clone()),
        Object::Reference(oid) => match doc.get_object(*oid).ok()? {
            Object::Dictionary(d) => Some(d.clone()),
            _ => None,
        },
        _ => None,
    }
}

fn extract_page_links(doc: &lopdf::Document, page_id: lopdf::ObjectId) -> Vec<String> {
    let Ok(page_obj) = doc.get_object(page_id) else {
        return Vec::new();
    };
    let Some(page_dict) = resolve_dict(doc, page_obj) else {
        return Vec::new();
    };

    let Ok(annots_obj) = page_dict.get(b"Annots") else {
        return Vec::new();
    };
    let annots = match annots_obj {
        Object::Array(arr) => arr.clone(),
        Object::Reference(oid) => match doc.get_object(*oid).ok() {
            Some(Object::Array(arr)) => arr.clone(),
            _ => return Vec::new(),
        },
        _ => return Vec::new(),
    };

    let mut urls = Vec::new();
    for a in &annots {
        let Some(d) = resolve_dict(doc, a) else {
            continue;
        };
        let Ok(Object::Name(sub)) = d.get(b"Subtype") else {
            continue;
        };
        if sub.as_slice() != b"Link" {
            continue;
        }

        let Ok(a_obj) = d.get(b"A") else {
            continue;
        };
        let Some(ad) = resolve_dict(doc, a_obj) else {
            continue;
        };
        let Ok(Object::Name(s)) = ad.get(b"S") else {
            continue;
        };
        if s.as_slice() != b"URI" {
            continue;
        }

        let Ok(Object::String(uri, _)) = ad.get(b"URI") else {
            continue;
        };
        let u = decode_pdf_string(uri).trim().to_string();
        if !u.is_empty() {
            urls.push(u);
        }
    }
    urls.sort();
    urls.dedup();
    urls
}

#[derive(Clone, Debug)]
struct PdfOutlineEntry {
    title: String,
    level: u8,
    page_number: u32,
}

fn parse_pdf_outline(
    doc: &lopdf::Document,
    pages: &BTreeMap<u32, lopdf::ObjectId>,
) -> Vec<PdfOutlineEntry> {
    let mut out = Vec::new();

    let page_num_by_id: HashMap<lopdf::ObjectId, u32> =
        pages.iter().map(|(n, oid)| (*oid, *n)).collect();

    let Ok(root_obj) = doc.trailer.get(b"Root") else {
        return out;
    };
    let Some(catalog) = resolve_dict(doc, root_obj) else {
        return out;
    };

    let Ok(outlines_obj) = catalog.get(b"Outlines") else {
        return out;
    };
    let Some(outlines) = resolve_dict(doc, outlines_obj) else {
        return out;
    };

    let Ok(first_obj) = outlines.get(b"First") else {
        return out;
    };

    fn resolve_dest(doc: &lopdf::Document, item: &lopdf::Dictionary) -> Option<Object> {
        if let Ok(d) = item.get(b"Dest") {
            return Some(d.clone());
        }
        let Ok(a_obj) = item.get(b"A") else {
            return None;
        };
        let a = resolve_dict(doc, a_obj)?;
        let Ok(s) = a.get(b"S") else {
            return None;
        };
        if let Object::Name(n) = s {
            if n.as_slice() != b"GoTo" {
                return None;
            }
        }
        let Ok(d) = a.get(b"D") else {
            return None;
        };
        Some(d.clone())
    }

    fn resolve_dest_array(doc: &lopdf::Document, d: &Object) -> Option<Vec<Object>> {
        match d {
            Object::Array(arr) => Some(arr.clone()),
            Object::Reference(oid) => match doc.get_object(*oid).ok()? {
                Object::Array(arr) => Some(arr.clone()),
                _ => None,
            },
            _ => None,
        }
    }

    fn walk(
        doc: &lopdf::Document,
        page_num_by_id: &HashMap<lopdf::ObjectId, u32>,
        item_ref: Object,
        level: u8,
        out: &mut Vec<PdfOutlineEntry>,
        visited: &mut std::collections::HashSet<lopdf::ObjectId>,
    ) {
        if out.len() >= 2_000 {
            return;
        }

        let Object::Reference(oid) = item_ref else {
            return;
        };

        if !visited.insert(oid) {
            return;
        }

        let Ok(obj) = doc.get_object(oid) else {
            return;
        };
        let Some(item) = resolve_dict(doc, obj) else {
            return;
        };

        let title = item
            .get(b"Title")
            .ok()
            .and_then(|o| match o {
                Object::String(b, _) => Some(decode_pdf_string(b)),
                _ => None,
            })
            .unwrap_or_default()
            .trim()
            .to_string();

        let page_number = resolve_dest(doc, &item)
            .and_then(|d| resolve_dest_array(doc, &d))
            .and_then(|arr| arr.first().cloned())
            .and_then(|o| match o {
                Object::Reference(pid) => page_num_by_id.get(&pid).copied(),
                _ => None,
            });

        if !title.is_empty() {
            if let Some(page_number) = page_number {
                out.push(PdfOutlineEntry {
                    title,
                    level: level.max(1),
                    page_number,
                });
            }
        }

        if let Ok(child) = item.get(b"First") {
            walk(
                doc,
                page_num_by_id,
                child.clone(),
                level.saturating_add(1),
                out,
                visited,
            );
        }
        if let Ok(next) = item.get(b"Next") {
            walk(doc, page_num_by_id, next.clone(), level, out, visited);
        }
    }

    let mut visited = std::collections::HashSet::new();
    walk(
        doc,
        &page_num_by_id,
        first_obj.clone(),
        1,
        &mut out,
        &mut visited,
    );
    out
}

fn handle_content_ops(
    doc: &lopdf::Document,
    ops: &[lopdf::content::Operation],
    resources: &lopdf::Dictionary,
    fonts: &HashMap<Vec<u8>, lopdf::Dictionary>,
    encodings: &HashMap<Vec<u8>, lopdf::Encoding>,
    page_number: u32,
    state: &mut PdfTextState,
    state_stack: &mut Vec<PdfTextState>,
    atoms: &mut Vec<TextAtom>,
    images: &mut Vec<ParsedImage>,
    image_by_hash: &mut HashMap<String, usize>,
    image_blocks: &mut Vec<(u32, String, String)>,
) -> Result<()> {
    for op in ops {
        let operator = op.operator.as_str();
        let operands = &op.operands;

        match operator {
            "BT" => {
                state.bt();
            }
            "q" => {
                state_stack.push(state.clone());
            }
            "Q" => {
                if let Some(prev) = state_stack.pop() {
                    *state = prev;
                }
            }
            "ET" => {
                state.et();
            }
            "Tf" => {
                if operands.len() >= 2 {
                    if let Object::Name(n) = &operands[0] {
                        state.font_alias = Some(n.clone());
                    }
                    if let Some(sz) = obj_to_f32(&operands[1]) {
                        state.font_size = sz.abs();
                        if state.leading <= 0.0 {
                            state.leading = (sz.abs() * 1.2).max(1.0);
                        }
                    }
                }
            }
            "Tm" => {
                // a b c d e f
                if operands.len() >= 6 {
                    if let (Some(a), Some(b), Some(c), Some(d), Some(e), Some(f)) = (
                        obj_to_f32(&operands[0]),
                        obj_to_f32(&operands[1]),
                        obj_to_f32(&operands[2]),
                        obj_to_f32(&operands[3]),
                        obj_to_f32(&operands[4]),
                        obj_to_f32(&operands[5]),
                    ) {
                        state.tm = Matrix2D::from_components(a, b, c, d, e, f);
                    }
                }
            }
            "cm" => {
                // a b c d e f
                if operands.len() >= 6 {
                    if let (Some(a), Some(b), Some(c), Some(d), Some(e), Some(f)) = (
                        obj_to_f32(&operands[0]),
                        obj_to_f32(&operands[1]),
                        obj_to_f32(&operands[2]),
                        obj_to_f32(&operands[3]),
                        obj_to_f32(&operands[4]),
                        obj_to_f32(&operands[5]),
                    ) {
                        let m = Matrix2D::from_components(a, b, c, d, e, f);
                        state.ctm = state.ctm.multiply(&m);
                    }
                }
            }
            "Td" => {
                if operands.len() >= 2 {
                    if let (Some(tx), Some(ty)) =
                        (obj_to_f32(&operands[0]), obj_to_f32(&operands[1]))
                    {
                        let t = Matrix2D::from_components(1.0, 0.0, 0.0, 1.0, tx, ty);
                        state.tm = state.tm.multiply(&t);
                    }
                }
            }
            "TD" => {
                if operands.len() >= 2 {
                    if let (Some(tx), Some(ty)) =
                        (obj_to_f32(&operands[0]), obj_to_f32(&operands[1]))
                    {
                        state.leading = (-ty).abs().max(1.0);
                        let t = Matrix2D::from_components(1.0, 0.0, 0.0, 1.0, tx, ty);
                        state.tm = state.tm.multiply(&t);
                    }
                }
            }
            "TL" => {
                if let Some(tl) = operands.get(0).and_then(obj_to_f32) {
                    state.leading = tl.abs().max(1.0);
                }
            }
            "T*" => {
                let t = Matrix2D::from_components(1.0, 0.0, 0.0, 1.0, 0.0, -state.leading);
                state.tm = state.tm.multiply(&t);
            }
            "Tj" => {
                if !state.in_text {
                    continue;
                }
                let alias = state.font_alias.as_deref();
                if let Some(s) = operands
                    .get(0)
                    .and_then(|o| decode_text_object(encodings, alias, o))
                {
                    let t = normalize_atom_text(&s);
                    if !t.is_empty() {
                        let seq = atoms.len() as u32;
                        let fs = state.effective_font_size();
                        let (x, y) = state.current_position();
                        atoms.push(TextAtom {
                            seq,
                            x,
                            y,
                            font_size: fs,
                            text: t,
                        });
                    }
                }
            }
            "TJ" => {
                if !state.in_text {
                    continue;
                }
                if let Some(Object::Array(arr)) = operands.get(0) {
                    let mut s = String::new();
                    let alias = state.font_alias.as_deref();
                    let fs = state.effective_font_size();
                    for o in arr {
                        if let Some(t) = decode_text_object(encodings, alias, o) {
                            s.push_str(&t);
                        }
                    }
                    let t = normalize_atom_text(&s);
                    if !t.is_empty() {
                        let seq = atoms.len() as u32;
                        let (x, y) = state.current_position();
                        atoms.push(TextAtom {
                            seq,
                            x,
                            y,
                            font_size: fs,
                            text: t,
                        });
                    }
                }
            }
            "'" => {
                // Move to next line then show text.
                let tmove = Matrix2D::from_components(1.0, 0.0, 0.0, 1.0, 0.0, -state.leading);
                state.tm = state.tm.multiply(&tmove);
                if !state.in_text {
                    continue;
                }
                let alias = state.font_alias.as_deref();
                if let Some(s) = operands
                    .get(0)
                    .and_then(|o| decode_text_object(encodings, alias, o))
                {
                    let t = normalize_atom_text(&s);
                    if !t.is_empty() {
                        let seq = atoms.len() as u32;
                        let fs = state.effective_font_size();
                        let (x, y) = state.current_position();
                        atoms.push(TextAtom {
                            seq,
                            x,
                            y,
                            font_size: fs,
                            text: t,
                        });
                    }
                }
            }
            "\"" => {
                // wordspace charspace string
                let tmove = Matrix2D::from_components(1.0, 0.0, 0.0, 1.0, 0.0, -state.leading);
                state.tm = state.tm.multiply(&tmove);
                if !state.in_text {
                    continue;
                }
                let alias = state.font_alias.as_deref();
                if let Some(s) = operands
                    .get(2)
                    .and_then(|o| decode_text_object(encodings, alias, o))
                {
                    let t = normalize_atom_text(&s);
                    if !t.is_empty() {
                        let seq = atoms.len() as u32;
                        let fs = state.effective_font_size();
                        let (x, y) = state.current_position();
                        atoms.push(TextAtom {
                            seq,
                            x,
                            y,
                            font_size: fs,
                            text: t,
                        });
                    }
                }
            }
            "BDC" => {
                // tag, properties
                if !state.in_text {
                    continue;
                }
                if operands.len() >= 2 {
                    let props = &operands[1];
                    let dict = match props {
                        Object::Dictionary(d) => Some(d),
                        Object::Reference(oid) => {
                            let o = doc.get_object(*oid).ok();
                            match o {
                                Some(Object::Dictionary(d)) => Some(d),
                                _ => None,
                            }
                        }
                        _ => None,
                    };
                    if let Some(d) = dict {
                        if let Ok(Object::String(b, _)) = d.get(b"ActualText") {
                            let t = normalize_atom_text(&decode_pdf_string(b));
                            if !t.is_empty() {
                                let seq = atoms.len() as u32;
                                let fs = state.effective_font_size();
                                let (x, y) = state.current_position();
                                atoms.push(TextAtom {
                                    seq,
                                    x,
                                    y,
                                    font_size: fs,
                                    text: t,
                                });
                            }
                        }
                    }
                }
            }
            "Do" => {
                if operands.len() >= 1 {
                    let Object::Name(name) = &operands[0] else {
                        continue;
                    };
                    let Some(stream) = resources_get_xobject_stream(doc, resources, name)? else {
                        continue;
                    };
                    let subtype = stream_subtype(&stream).unwrap_or_default();

                    if subtype == b"Form".to_vec() {
                        let mut form_resources = resources.clone();
                        if let Ok(obj) = stream.dict.get(b"Resources") {
                            if let Some(r) = resolve_dict(doc, obj) {
                                merge_resources_into(doc, &mut form_resources, &r);
                            }
                        }

                        let mut form_fonts = fonts.clone();
                        form_fonts.extend(collect_fonts_from_resources(doc, &form_resources));

                        let form_encodings: HashMap<Vec<u8>, lopdf::Encoding> = form_fonts
                            .iter()
                            .filter_map(|(name, font)| {
                                font.get_font_encoding(doc)
                                    .ok()
                                    .map(|enc| (name.clone(), enc))
                            })
                            .collect();

                        let bytes = stream
                            .decompressed_content()
                            .unwrap_or_else(|_| stream.content.clone());
                        let content =
                            lopdf::content::Content::decode(&bytes).with_context(|| {
                                format!("decode Form XObject content (page {page_number})")
                            })?;

                        let saved = state.clone();
                        let saved_stack = state_stack.clone();
                        handle_content_ops(
                            doc,
                            &content.operations,
                            &form_resources,
                            &form_fonts,
                            &form_encodings,
                            page_number,
                            state,
                            state_stack,
                            atoms,
                            images,
                            image_by_hash,
                            image_blocks,
                        )?;
                        *state = saved;
                        *state_stack = saved_stack;
                    } else if subtype == b"Image".to_vec() {
                        if let Some((id, mime)) =
                            extract_xobject_image(doc, &stream, images, image_by_hash)
                        {
                            image_blocks.push((page_number, id, mime));
                        }
                    }
                }
            }
            _ => {}
        }
    }
    Ok(())
}

fn merge_resources_into(
    doc: &lopdf::Document,
    dst: &mut lopdf::Dictionary,
    src: &lopdf::Dictionary,
) {
    fn dict_from_obj(doc: &lopdf::Document, o: &Object) -> Option<lopdf::Dictionary> {
        match o {
            Object::Dictionary(d) => Some(d.clone()),
            Object::Reference(oid) => match doc.get_object(*oid).ok()? {
                Object::Dictionary(d) => Some(d.clone()),
                _ => None,
            },
            _ => None,
        }
    }

    fn merge_named_dict(
        doc: &lopdf::Document,
        dst: &mut lopdf::Dictionary,
        key: &[u8],
        src_val: &Object,
    ) {
        let Some(src_dict) = dict_from_obj(doc, src_val) else {
            return;
        };

        let merged = if let Ok(dst_val) = dst.get(key) {
            if let Some(dst_dict) = dict_from_obj(doc, dst_val) {
                let mut out = dst_dict;
                for (k, v) in src_dict.iter() {
                    out.set(k.clone(), v.clone());
                }
                out
            } else {
                src_dict
            }
        } else {
            src_dict
        };

        dst.set(key.to_vec(), Object::Dictionary(merged));
    }

    for (k, v) in src.iter() {
        match k.as_slice() {
            b"Font" => merge_named_dict(doc, dst, b"Font", v),
            b"XObject" => merge_named_dict(doc, dst, b"XObject", v),
            _ => {
                dst.set(k.clone(), v.clone());
            }
        }
    }
}

fn build_page_resources(
    doc: &lopdf::Document,
    page_id: lopdf::ObjectId,
    page_number: u32,
) -> Result<lopdf::Dictionary> {
    let (res_opt, resource_ids) = doc
        .get_page_resources(page_id)
        .with_context(|| format!("get page resources (page {page_number})"))?;

    let mut out = lopdf::Dictionary::new();

    // get_page_resources collects resource references from the page node upwards; reverse so
    // parents are merged first and the page-level resources take precedence.
    for rid in resource_ids.iter().rev() {
        if let Ok(d) = doc.get_dictionary(*rid) {
            merge_resources_into(doc, &mut out, d);
        }
    }
    if let Some(d) = res_opt {
        merge_resources_into(doc, &mut out, d);
    }
    Ok(out)
}

#[derive(Clone, Debug)]
struct Line {
    y: f32,
    font_size: f32,
    text: String,
}

#[derive(Clone, Debug)]
struct Paragraph {
    y_top: f32,
    font_size: f32,
    lines: Vec<String>,
}

fn median_f32(mut v: Vec<f32>, default: f32) -> f32 {
    v.retain(|x| x.is_finite() && *x > 0.0);
    if v.is_empty() {
        return default;
    }
    v.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let n = v.len();
    if n % 2 == 1 {
        v[n / 2]
    } else {
        let a = v[(n / 2).saturating_sub(1)];
        let b = v[n / 2];
        (a + b) / 2.0
    }
}

fn cluster_atoms_to_lines(atoms: &[TextAtom]) -> Vec<Line> {
    let sizes = atoms.iter().map(|a| a.font_size).collect::<Vec<_>>();
    let med = median_f32(sizes, 12.0);
    let eps_y = (0.5 * med).max(0.5);

    let mut sorted = atoms.to_vec();
    sorted.sort_by(|a, b| {
        // y desc, x asc, font_size desc, seq asc (deterministic, preserves content-stream order)
        let y_cmp = b.y.partial_cmp(&a.y).unwrap_or(std::cmp::Ordering::Equal);
        if y_cmp != std::cmp::Ordering::Equal {
            return y_cmp;
        }
        let x_cmp = a.x.partial_cmp(&b.x).unwrap_or(std::cmp::Ordering::Equal);
        if x_cmp != std::cmp::Ordering::Equal {
            return x_cmp;
        }
        let fs_cmp = b
            .font_size
            .partial_cmp(&a.font_size)
            .unwrap_or(std::cmp::Ordering::Equal);
        if fs_cmp != std::cmp::Ordering::Equal {
            return fs_cmp;
        }
        a.seq.cmp(&b.seq)
    });

    let mut lines: Vec<(f32, Vec<TextAtom>)> = Vec::new();
    for a in sorted {
        if let Some((ly, group)) = lines.last_mut() {
            if (a.y - *ly).abs() <= eps_y {
                group.push(a);
                continue;
            }
        }
        lines.push((a.y, vec![a]));
    }

    let mut out = Vec::new();
    for (y, mut group) in lines {
        group.sort_by(|a, b| {
            let x_cmp = a.x.partial_cmp(&b.x).unwrap_or(std::cmp::Ordering::Equal);
            if x_cmp != std::cmp::Ordering::Equal {
                return x_cmp;
            }
            a.seq.cmp(&b.seq)
        });
        let font_size = median_f32(group.iter().map(|a| a.font_size).collect(), 12.0);
        let space_thr = (0.65 * font_size).max(1.0);
        let word_thr = (1.4 * font_size).max(space_thr);
        let mut text = String::new();
        let mut prev_end_x: Option<f32> = None;
        let mut prev_last_ch: Option<char> = None;
        for (i, a) in group.iter().enumerate() {
            let s = a.text.trim();
            if s.is_empty() {
                continue;
            }
            let adv = estimate_text_advance(s, font_size);
            if i > 0 {
                if let Some(pe) = prev_end_x {
                    let dx = a.x - pe;
                    let first = s.chars().next();
                    let thr = match (prev_last_ch, first) {
                        (Some(p), Some(c)) if p.is_alphanumeric() && c.is_alphanumeric() => {
                            word_thr
                        }
                        _ => space_thr,
                    };
                    if dx.is_finite() && dx > thr && !text.ends_with(' ') {
                        text.push(' ');
                    }
                } else if !text.ends_with(' ') {
                    text.push(' ');
                }
            }
            text.push_str(s);
            prev_end_x = Some(a.x + adv);
            prev_last_ch = s.chars().last();
        }
        let t = text.trim().to_string();
        if !t.is_empty() {
            out.push(Line {
                y,
                font_size,
                text: t,
            });
        }
    }
    out
}

fn cluster_lines_to_paragraphs(lines: &[Line]) -> Vec<Paragraph> {
    if lines.is_empty() {
        return Vec::new();
    }
    let med = median_f32(lines.iter().map(|l| l.font_size).collect(), 12.0);
    let gap_thr = (1.5 * med).max(6.0);

    let mut sorted = lines.to_vec();
    sorted.sort_by(|a, b| b.y.partial_cmp(&a.y).unwrap_or(std::cmp::Ordering::Equal));

    let mut out: Vec<Paragraph> = Vec::new();
    let mut cur: Option<Paragraph> = None;
    let mut prev_y: Option<f32> = None;

    for l in sorted {
        if let Some(py) = prev_y {
            let gap = (py - l.y).abs();
            if gap > gap_thr {
                if let Some(p) = cur.take() {
                    out.push(p);
                }
            }
        }
        prev_y = Some(l.y);

        if cur.is_none() {
            cur = Some(Paragraph {
                y_top: l.y,
                font_size: l.font_size,
                lines: vec![l.text],
            });
            continue;
        }

        let p = cur.as_mut().unwrap();
        p.font_size = p.font_size.max(l.font_size);
        p.lines.push(l.text);
    }

    if let Some(p) = cur.take() {
        out.push(p);
    }

    out
}

fn looks_like_list_item(s: &str) -> Option<(bool, String)> {
    let t = s.trim_start();
    for bullet in ["• ", "- ", "* ", "– ", "— "] {
        if let Some(rest) = t.strip_prefix(bullet) {
            return Some((false, rest.trim().to_string()));
        }
    }

    // 1. item / 1) item
    let mut digits = 0usize;
    for ch in t.chars().take(8) {
        if ch.is_ascii_digit() {
            digits += 1;
            continue;
        }
        break;
    }
    if digits > 0 {
        let rest = &t[digits..];
        let rest = rest.strip_prefix(".").or_else(|| rest.strip_prefix(")"));
        if let Some(rest) = rest {
            let rest = rest.trim_start();
            if !rest.is_empty() {
                return Some((true, rest.to_string()));
            }
        }
    }
    None
}

fn split_tableish_columns(line: &str) -> Option<Vec<String>> {
    let t = line.trim();
    if t.is_empty() {
        return None;
    }

    let mut cols: Vec<String> = Vec::new();
    let mut cur = String::new();
    let mut pending_spaces = 0usize;
    let mut saw_delim = false;

    let flush = |cols: &mut Vec<String>, cur: &mut String| {
        let c = cur.trim();
        if !c.is_empty() {
            cols.push(c.to_string());
        }
        cur.clear();
    };

    for ch in t.chars() {
        if ch == '\t' {
            saw_delim = true;
            pending_spaces = 0;
            flush(&mut cols, &mut cur);
            continue;
        }
        if ch == ' ' {
            pending_spaces += 1;
            continue;
        }

        if pending_spaces == 1 {
            cur.push(' ');
        } else if pending_spaces >= 2 {
            saw_delim = true;
            flush(&mut cols, &mut cur);
        }
        pending_spaces = 0;
        cur.push(ch);
    }

    // Trailing spaces: ignore.
    flush(&mut cols, &mut cur);

    if !saw_delim {
        return None;
    }
    if cols.len() < 2 {
        return None;
    }
    Some(cols)
}

fn paragraph_as_table(lines: &[String]) -> Option<Vec<Vec<Cell>>> {
    if lines.len() < 3 {
        return None;
    }
    let mut rows_raw: Vec<Vec<String>> = Vec::new();
    for l in lines {
        rows_raw.push(split_tableish_columns(l)?);
    }

    let cols_n = rows_raw.first()?.len();
    if cols_n < 2 || cols_n > 10 {
        return None;
    }
    if rows_raw.iter().any(|r| r.len() != cols_n) {
        return None;
    }

    Some(
        rows_raw
            .into_iter()
            .map(|r| {
                r.into_iter()
                    .map(|t| Cell {
                        text: t,
                        colspan: 1,
                        rowspan: 1,
                    })
                    .collect::<Vec<_>>()
            })
            .collect::<Vec<_>>(),
    )
}

fn heading_level_from_ratio(r: f32) -> u8 {
    if r >= 2.0 {
        1
    } else if r >= 1.6 {
        2
    } else if r >= 1.3 {
        3
    } else if r >= 1.15 {
        4
    } else {
        5
    }
}

fn looks_like_heading_text(s: &str) -> bool {
    let s = s.trim();
    if s.is_empty() {
        return false;
    }
    if s.chars().count() > 180 {
        return false;
    }
    if s.ends_with('.') {
        return false;
    }

    const STOP: &[&str] = &[
        "and", "or", "of", "the", "a", "an", "to", "in", "on", "for", "with", "at", "by", "from",
        "vs", "via", "as", "is", "are", "into", "over", "under",
    ];

    let mut sig = 0usize;
    let mut caps = 0usize;
    let mut letters = 0usize;
    let mut upper = 0usize;

    for raw in s.split_whitespace() {
        let w = raw.trim_matches(|c: char| !c.is_ascii_alphanumeric());
        if w.is_empty() {
            continue;
        }
        let lw = w.to_ascii_lowercase();
        if STOP.contains(&lw.as_str()) {
            continue;
        }

        let mut has_alpha = false;
        for ch in w.chars() {
            if ch.is_ascii_alphabetic() {
                has_alpha = true;
                letters += 1;
                if ch.is_ascii_uppercase() {
                    upper += 1;
                }
            }
        }
        if !has_alpha {
            continue;
        }

        sig += 1;

        // TitleCase or ALLCAPS word.
        let mut chars = w.chars().filter(|c| c.is_ascii_alphabetic());
        let first = chars.next();
        let rest: String = chars.collect();
        let is_all_caps = w
            .chars()
            .filter(|c| c.is_ascii_alphabetic())
            .all(|c| c.is_ascii_uppercase());
        let is_title = first.map(|c| c.is_ascii_uppercase()).unwrap_or(false)
            && rest.chars().all(|c| c.is_ascii_lowercase());
        if is_all_caps || is_title {
            caps += 1;
        }
    }

    if sig < 2 {
        return false;
    }

    let caps_ratio = caps as f32 / sig as f32;
    let upper_ratio = if letters == 0 {
        0.0
    } else {
        upper as f32 / letters as f32
    };

    caps_ratio >= 0.70 || upper_ratio >= 0.85
}

pub fn parse_pdf_full(bytes: &[u8]) -> Result<ParsedPdfDocument> {
    let doc = lopdf::Document::load_mem(bytes).context("lopdf load_mem")?;
    let pages: BTreeMap<u32, lopdf::ObjectId> = doc.get_pages();

    let outline = parse_pdf_outline(&doc, &pages);

    let mut images: Vec<ParsedImage> = Vec::new();
    let mut image_by_hash: HashMap<String, usize> = HashMap::new();

    let mut blocks: Vec<Block> = Vec::new();
    let mut page_block_ranges: Vec<PdfPageBlockRange> = Vec::new();
    let mut next_block_index: usize = 0;

    for (page_number, page_id) in pages {
        let resources = build_page_resources(&doc, page_id, page_number)?;
        let content_bytes = doc
            .get_page_content(page_id)
            .with_context(|| format!("get page content (page {page_number})"))?;
        let content = lopdf::content::Content::decode(&content_bytes)
            .with_context(|| format!("decode page content (page {page_number})"))?;

        let mut fonts: HashMap<Vec<u8>, lopdf::Dictionary> = HashMap::new();
        if let Ok(page_fonts) = doc.get_page_fonts(page_id) {
            for (name, dict) in page_fonts {
                fonts.insert(name, dict.clone());
            }
        }
        fonts.extend(collect_fonts_from_resources(&doc, &resources));

        let mut atoms: Vec<TextAtom> = Vec::new();
        let mut image_blocks: Vec<(u32, String, String)> = Vec::new();
        let mut state = PdfTextState::default();
        let mut state_stack: Vec<PdfTextState> = Vec::new();

        let encodings: HashMap<Vec<u8>, lopdf::Encoding> = fonts
            .iter()
            .filter_map(|(name, font)| {
                font.get_font_encoding(&doc)
                    .ok()
                    .map(|enc| (name.clone(), enc))
            })
            .collect();
        handle_content_ops(
            &doc,
            &content.operations,
            &resources,
            &fonts,
            &encodings,
            page_number,
            &mut state,
            &mut state_stack,
            &mut atoms,
            &mut images,
            &mut image_by_hash,
            &mut image_blocks,
        )?;

        let lines = cluster_atoms_to_lines(&atoms);
        let paragraphs = cluster_lines_to_paragraphs(&lines);

        let y_min = lines
            .iter()
            .map(|l| l.y)
            .fold(f32::INFINITY, |a, b| a.min(b));
        let y_max = lines
            .iter()
            .map(|l| l.y)
            .fold(f32::NEG_INFINITY, |a, b| a.max(b));
        let y_span = (y_max - y_min).abs().max(1.0);

        // Estimate body font-size by ignoring the top area (where titles/headers live).
        // This improves heading detection on cover-like pages.
        let para_sizes_all = paragraphs.iter().map(|p| p.font_size).collect::<Vec<_>>();
        let med_para_all = median_f32(para_sizes_all, 12.0);
        let body_sizes = paragraphs
            .iter()
            .filter(|p| p.y_top <= (y_max - 0.35 * y_span))
            .map(|p| p.font_size)
            .collect::<Vec<_>>();
        let med_para = median_f32(body_sizes, med_para_all);

        let page_block_start = blocks.len();

        let mut para_idx: usize = 0;

        // Special-case: attempt to extract a cover/title block at the very top of page 1.
        // PDFs do not have semantics; this is a best-effort heuristic.
        if page_number == 1 {
            let top_thr = y_max - (0.35 * y_span);
            let max_fs = paragraphs
                .iter()
                .map(|p| p.font_size)
                .fold(0.0f32, |a, b| a.max(b));
            let title_min_fs = (max_fs * 0.65).max(med_para * 1.35);
            let mut title_parts: Vec<String> = Vec::new();
            while para_idx < paragraphs.len() {
                let p = &paragraphs[para_idx];
                let joined = p.lines.join("\n").trim().to_string();
                if joined.is_empty() {
                    para_idx += 1;
                    continue;
                }

                let is_bodyish = joined.chars().count() > 200 && p.font_size <= (med_para * 1.20);
                if is_bodyish {
                    break;
                }

                let is_title_like = p.y_top >= top_thr && p.font_size >= title_min_fs;
                if !is_title_like {
                    break;
                }

                title_parts.push(joined.split_whitespace().collect::<Vec<_>>().join(" "));
                para_idx += 1;

                if title_parts.len() >= 20 {
                    break;
                }
            }

            if !title_parts.is_empty() {
                blocks.push(Block::Heading {
                    block_index: next_block_index,
                    level: 1,
                    text: title_parts.join(" "),
                    source: SourceSpan::default(),
                    marks: Vec::new(),
    });
                next_block_index += 1;
            }
        }

        while para_idx < paragraphs.len() {
            let p = &paragraphs[para_idx];
            let joined = p.lines.join("\n").trim().to_string();
            if joined.is_empty() {
                para_idx += 1;
                continue;
            }
            let ratio = (p.font_size / med_para).max(0.0);

            if let Some(rows) = paragraph_as_table(&p.lines) {
                blocks.push(Block::Table {
                    block_index: next_block_index,
                    rows,
                    source: SourceSpan::default(),
                    widths: Vec::new(),
                });
                next_block_index += 1;
                para_idx += 1;
                continue;
            }

            // List detection: treat multi-line paragraphs with list markers as a single list.
            let mut list_items: Vec<(bool, String)> = Vec::new();
            for l in &p.lines {
                if let Some(li) = looks_like_list_item(l) {
                    list_items.push(li);
                } else {
                    list_items.clear();
                    break;
                }
            }
            if !list_items.is_empty() {
                let ordered = list_items.iter().all(|(o, _)| *o);
                let items = list_items
                    .into_iter()
                    .map(|(_ord, t)| ListItem {
                        level: 0,
                        text: t,
                        source: SourceSpan::default(),
                        marks: Vec::new(),
    })
                    .collect::<Vec<_>>();
                blocks.push(Block::List {
                    block_index: next_block_index,
                    ordered,
                    items,
                    source: SourceSpan::default(),
                });
                next_block_index += 1;
                para_idx += 1;
                continue;
            }

            let joined_chars = joined.chars().count();
            let mut is_heading = ratio >= 1.30 && joined_chars <= 120;

            // PDFs often split headings across lines/paragraphs and use only slightly larger fonts.
            // Add a conservative, position-aware fallback.
            if !is_heading {
                let near_top = p.y_top >= (y_max - 0.55 * y_span);
                let shortish = joined_chars <= 180 && p.lines.len() <= 4;
                let fontish = ratio >= 1.12;
                let semantic = looks_like_heading_text(&joined);
                let sentence_like = joined.trim_end().ends_with('.') && joined_chars > 80;
                // Allow semantic/titlecase headings that are the same size as body text, but only
                // near the top of the page (common in PDFs that don't increase font size for heads).
                if near_top
                    && shortish
                    && !sentence_like
                    && (fontish || (ratio >= 0.98 && semantic))
                {
                    is_heading = true;
                }

                // Some PDFs use body-sized text for headings, and the heading can appear mid-page.
                // Detect the common "2-line heading then body" pattern using local neighborhood.
                if !is_heading && shortish && !sentence_like && semantic && ratio >= 0.90 {
                    if let Some(p2) = paragraphs.get(para_idx + 1) {
                        let t2 = p2.lines.join("\n").trim().to_string();
                        let t2_chars = t2.chars().count();
                        let shortish2 = t2_chars <= 120 && p2.lines.len() <= 4;
                        let semantic2 = looks_like_heading_text(&t2);
                        let rel_fs = (p2.font_size - p.font_size).abs() / p.font_size.max(1.0);
                        if !t2.is_empty() && shortish2 && semantic2 && rel_fs <= 0.20 {
                            let bodyish_after = paragraphs
                                .get(para_idx + 2)
                                .map(|p3| {
                                    let t3 = p3.lines.join("\n").trim().to_string();
                                    let n = t3.chars().count();
                                    n > 140 || t3.ends_with('.')
                                })
                                .unwrap_or(true);
                            if bodyish_after {
                                is_heading = true;
                            }
                        }
                    }
                }
            }
            if is_heading {
                let mut level = heading_level_from_ratio(ratio).max(1).min(6);
                if page_number == 1 && ratio < 1.30 {
                    level = 2;
                }

                let mut text = joined.split_whitespace().collect::<Vec<_>>().join(" ");

                // Merge multi-line headings that ended up split into adjacent paragraphs.
                let mut j = para_idx + 1;
                while j < paragraphs.len() {
                    let p2 = &paragraphs[j];
                    let t2 = p2.lines.join("\n").trim().to_string();
                    if t2.is_empty() {
                        j += 1;
                        continue;
                    }
                    let ratio2 = (p2.font_size / med_para).max(0.0);
                    let t2_chars = t2.chars().count();
                    let mut is_heading2 = ratio2 >= 1.30 && t2_chars <= 120;
                    if !is_heading2 {
                        let near_top = p2.y_top >= (y_max - 0.55 * y_span);
                        let shortish = t2_chars <= 180 && p2.lines.len() <= 4;
                        let fontish = ratio2 >= 1.12;
                        let semantic = looks_like_heading_text(&t2);
                        let sentence_like = t2.trim_end().ends_with('.') && t2_chars > 80;
                        if (near_top
                            && shortish
                            && !sentence_like
                            && (fontish || (ratio2 >= 0.98 && semantic)))
                            || (shortish && !sentence_like && semantic && ratio2 >= 0.90)
                        {
                            is_heading2 = true;
                        }
                    }
                    if !is_heading2 {
                        break;
                    }
                    let mut level2 = heading_level_from_ratio(ratio2).max(1).min(6);
                    if page_number == 1 && ratio2 < 1.30 {
                        level2 = 2;
                    }
                    if level2 != level {
                        break;
                    }
                    let rel = (p2.font_size - p.font_size).abs() / p.font_size.max(1.0);
                    if rel > 0.20 {
                        break;
                    }
                    if text.chars().count() + 1 + t2.chars().count() > 160 {
                        break;
                    }

                    text.push(' ');
                    text.push_str(&t2.split_whitespace().collect::<Vec<_>>().join(" "));
                    j += 1;
                }

                blocks.push(Block::Heading {
                    block_index: next_block_index,
                    level,
                    text,
                    source: SourceSpan::default(),
                    marks: Vec::new(),
    });
                next_block_index += 1;
                para_idx = j;
            } else {
                blocks.push(Block::Paragraph {
                    block_index: next_block_index,
                    text: joined,
                    source: SourceSpan::default(),
                    marks: Vec::new(),
    });
                next_block_index += 1;
                para_idx += 1;
            }
        }

        // Append image blocks for this page deterministically after text blocks.
        image_blocks.sort_by(|a, b| (a.0, &a.1).cmp(&(b.0, &b.1)));
        for (_pn, id, mime) in image_blocks {
            blocks.push(Block::Image {
                block_index: next_block_index,
                id,
                filename: None,
                content_type: Some(mime),
                alt: None,
                source: SourceSpan::default(),
            });
            next_block_index += 1;
        }

        // Link annotations (best-effort).
        for url in extract_page_links(&doc, page_id) {
            blocks.push(Block::Link {
                block_index: next_block_index,
                url,
                text: None,
                kind: LinkKind::Unknown,
                source: SourceSpan::default(),
            });
            next_block_index += 1;
        }

        let page_block_end = blocks.len();
        page_block_ranges.push(PdfPageBlockRange {
            page_number,
            block_start: page_block_start,
            block_end: page_block_end,
        });
    }

    // Ensure monotonic block indices.
    for (idx, b) in blocks.iter_mut().enumerate() {
        match b {
            Block::Heading { block_index, .. }
            | Block::Paragraph { block_index, .. }
            | Block::List { block_index, .. }
            | Block::Table { block_index, .. }
            | Block::Image { block_index, .. }
            | Block::Link { block_index, .. }
            | Block::Note { block_index, .. } => *block_index = idx,
        }
    }

    cleanup_pdf_blocks(&mut blocks);

    fn extract_pdf_title(doc: &lopdf::Document) -> Option<String> {
        use lopdf::Object;

        let info = doc.trailer.get(b"Info").ok()?;
        let info = match info {
            Object::Reference(oid) => doc.get_object(*oid).ok()?,
            other => other,
        };
        let Object::Dictionary(dict) = info else {
            return None;
        };
        let title_obj = dict.get(b"Title").ok()?;
        match title_obj {
            Object::String(b, _) => {
                let t = decode_pdf_string(b).trim().to_string();
                if t.is_empty() { None } else { Some(t) }
            }
            _ => None,
        }
    }

    let title = extract_pdf_title(&doc);

    let metadata_json = serde_json::json!({
        "parser": "office-parser-pdf-v1",
        "title": title,
        "page_count": page_block_ranges.len(),
        "page_block_ranges": page_block_ranges.iter().map(|r| {
            serde_json::json!({
                "page_number": r.page_number,
                "block_start": r.block_start,
                "block_end": r.block_end,
            })
        }).collect::<Vec<_>>(),
        "outline": outline.iter().map(|e| {
            serde_json::json!({
                "title": e.title.clone(),
                "level": e.level,
                "page_number": e.page_number,
            })
        }).collect::<Vec<_>>(),
    });

    Ok(ParsedPdfDocument {
        blocks,
        images,
        metadata_json,
        page_block_ranges,
    })
}

/// Count alphanumeric characters across all blocks after cleanup.
fn count_alphanumeric(blocks: &[Block]) -> usize {
    blocks
        .iter()
        .map(|b| {
            let text: String = match b {
                Block::Heading { text, .. } => text.clone(),
                Block::Paragraph { text, .. } => text.clone(),
                Block::List { items, .. } => {
                    items.iter().map(|i| i.text.as_str()).collect::<Vec<_>>().join(" ")
                }
                Block::Table { rows, .. } => rows
                    .iter()
                    .flat_map(|r| r.iter())
                    .map(|c| c.text.as_str())
                    .collect::<Vec<_>>()
                    .join(" "),
                Block::Image { alt, .. } => alt.as_deref().unwrap_or("").to_string(),
                Block::Link { text, url, .. } => {
                    let t = text.as_deref().unwrap_or("");
                    format!("{t}{url}")
                }
                Block::Note { text, .. } => text.clone(),
            };
            text.chars().filter(|c| c.is_alphanumeric()).count()
        })
        .sum()
}

/// Classify PDF text extraction quality based on alphanumeric character count.
fn classify_pdf_text_quality(blocks: &[Block]) -> PdfTextQuality {
    let n = count_alphanumeric(blocks);
    if n > 500 {
        PdfTextQuality::TextRich
    } else if n >= 50 {
        PdfTextQuality::TextSparse
    } else {
        PdfTextQuality::ImageOnly
    }
}

/// Rasterize a PDF's pages to PNG bytes via poppler's `pdftoppm` (system dep).
///
/// OCR needs the rendered *page*, not the PDF's embedded images:
/// [`extract_xobject_image`] pulls logos/photos out of the file, while
/// text + layout as pixels only exist once the page is rasterized. ~100 dpi
/// keeps OCR content ≤ ~0.5 Mpx per page — the PaddleOCR-VL reading window
/// (see typellama `docs/BACKENDS.md`).
///
fn pdftoppm_available() -> bool {
    std::process::Command::new("pdftoppm")
        .arg("-v")
        .output()
        .map(|o| o.status.success())
        .unwrap_or(false)
}

/// Where a bundled Pdfium dynamic library would live (env PDFIUM_PATH, else
/// ~/.cache/office-parser/pdfium/ — see [`install_pdfium`]).
pub fn pdfium_lib_path() -> Option<std::path::PathBuf> {
    if let Ok(p) = std::env::var("PDFIUM_PATH") {
        let p = std::path::PathBuf::from(p);
        if p.exists() {
            return Some(p);
        }
    }
    let cache = std::env::var("HOME")
        .ok()
        .map(|h| std::path::PathBuf::from(h).join(".cache/office-parser/pdfium/libpdfium.so"))?;
    cache.exists().then_some(cache)
}

/// Download the Pdfium dynamic library for this platform into
/// `~/.cache/office-parser/pdfium/` (bblanchon/pdfium-binaries latest release;
/// Linux/macOS via curl+tar — same system-tool policy as pdftoppm; on Windows
/// download manually from that releases page). Requires the `render-pdfium`
/// feature to take effect at render time.
pub fn install_pdfium() -> Result<std::path::PathBuf> {
    use std::process::Command;
    let (asset, lib) = if cfg!(target_os = "macos") {
        ("pdfium-mac-arm64.tgz", "lib/libpdfium.dylib")
    } else if cfg!(target_os = "linux") {
        ("pdfium-linux-x64.tgz", "lib/libpdfium.so")
    } else {
        anyhow::bail!("automatic Pdfium install supports Linux/macOS; download manually from https://github.com/bblanchon/pdfium-binaries/releases");
    };
    let dir = std::env::temp_dir().join(format!("office-parser-pdfium-{}", std::process::id()));
    std::fs::create_dir_all(&dir)?;
    let result = (|| -> Result<std::path::PathBuf> {
        let tgz = dir.join(asset);
        let url = format!(
            "https://github.com/bblanchon/pdfium-binaries/releases/latest/download/{asset}"
        );
        let status = Command::new("curl")
            .args(["-sL", "--fail", "-o"])
            .arg(&tgz)
            .arg(&url)
            .status()
            .map_err(|e| anyhow::anyhow!("curl: {e}"))?;
        anyhow::ensure!(status.success(), "pdfium download failed: {status}");
        let status = Command::new("tar")
            .args(["-xzf"])
            .arg(&tgz)
            .arg("-C")
            .arg(&dir)
            .status()
            .map_err(|e| anyhow::anyhow!("tar: {e}"))?;
        anyhow::ensure!(status.success(), "pdfium extract failed");
        let home = std::env::var("HOME").map_err(|_| anyhow::anyhow!("HOME not set"))?;
        let dest_dir = std::path::PathBuf::from(home).join(".cache/office-parser/pdfium");
        std::fs::create_dir_all(&dest_dir)?;
        let lib_src = dir.join(lib);
        anyhow::ensure!(lib_src.exists(), "pdfium archive missing {lib}");
        let dest = dest_dir.join(lib.rsplit('/').next().unwrap());
        std::fs::copy(&lib, &dest)?;
        Ok(dest)
    })();
    let _ = std::fs::remove_dir_all(&dir);
    result
}

#[cfg(feature = "render-pdfium")]
/// pdfium allows ONE library init per process (its bindings live in a global
/// OnceLock inside pdfium-render), so the handle is cached and shared.
#[cfg(feature = "render-pdfium")]
fn pdfium_handle() -> Option<&'static std::sync::Mutex<pdfium_render::prelude::Pdfium>> {
    use pdfium_render::prelude::*;
    static PDFIUM: std::sync::OnceLock<Option<std::sync::Mutex<Pdfium>>> =
        std::sync::OnceLock::new();
    PDFIUM
        .get_or_init(|| {
            let bindings = match pdfium_lib_path() {
                Some(lib) => Pdfium::bind_to_library(lib)
                    .or_else(|_| Pdfium::bind_to_system_library())
                    .ok()?,
                None => Pdfium::bind_to_system_library().ok()?,
            };
            Some(std::sync::Mutex::new(Pdfium::new(bindings)))
        })
        .as_ref()
}

#[cfg(feature = "render-pdfium")]
fn render_pages_png_pdfium(bytes: &[u8], dpi: u32, clamp_budget: bool) -> Result<Vec<Vec<u8>>> {
    use pdfium_render::prelude::*;
    static SEQ: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
    let dir = std::env::temp_dir().join(format!(
        "office-parser-pdfium-doc-{}-{}",
        std::process::id(),
        SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
    ));
    std::fs::create_dir_all(&dir)?;
    let result = (|| -> Result<Vec<Vec<u8>>> {
        let src = dir.join("doc.pdf");
        std::fs::write(&src, bytes)?;
        let Some(pdfium) = pdfium_handle() else {
            anyhow::bail!("pdfium bindings unavailable");
        };
        let pdfium = pdfium.lock().map_err(|_| anyhow::anyhow!("pdfium mutex poisoned"))?;
        let document = pdfium
            .load_pdf_from_file(&src, None)
            .map_err(|e| anyhow::anyhow!("pdfium load: {e:?}"))?;
        // clamped: the legacy ~0.4 MP OCR window (longest side ≤ ~750).
        // unclamped: true dpi rendering (Indra resizes to 1024 itself).
        let config = if clamp_budget {
            PdfRenderConfig::new().set_target_width(530).set_maximum_height(750)
        } else {
            PdfRenderConfig::new().scale_page_by_factor(dpi as f32 / 72.0)
        };
        let mut pages = Vec::new();
        for (index, page) in document.pages().iter().enumerate() {
            let image = page
                .render_with_config(&config)?
                .as_image()?
                .into_rgb8();
            let mut png = std::io::Cursor::new(Vec::new());
            image
                .write_to(&mut png, ::image::ImageFormat::Png)
                .map_err(|e| anyhow::anyhow!("png encode page {index}: {e}"))?;
            pages.push(png.into_inner());
        }
        Ok(pages)
    })();
    let _ = std::fs::remove_dir_all(&dir);
    result
}

/// OCR megapixel budget: PaddleOCR-VL's reading window (empirical: ≤ ~0.4 MP
/// content; denser pages yield empty grids — typellama docs/BACKENDS.md).
const OCR_MEGAPIXEL_BUDGET: f64 = 0.4;

/// sqrt back-off: pixels scale with dpi², so dpi scales with √(budget/mp).
fn clamp_dpi_to_budget(page_pts_w: f64, page_pts_h: f64, dpi: u32) -> u32 {
    let scale = dpi as f64 / 72.0;
    let mp = (page_pts_w * scale) * (page_pts_h * scale) / 1_000_000.0;
    if mp <= OCR_MEGAPIXEL_BUDGET {
        return dpi;
    }
    let factor = (OCR_MEGAPIXEL_BUDGET / mp).sqrt();
    ((dpi as f64 * factor).floor() as u32).max(72)
}

/// Page size in PDF points via `pdfinfo` (same poppler-utils install as
/// pdftoppm). None if the probe is unavailable or unparseable.
fn page_size_pts(bytes: &[u8]) -> Option<(f64, f64)> {
    use std::process::Command;
    let dir = std::env::temp_dir().join(format!("office-parser-probe-{}", std::process::id()));
    std::fs::create_dir_all(&dir).ok()?;
    let src = dir.join("doc.pdf");
    std::fs::write(&src, bytes).ok()?;
    let out = Command::new("pdfinfo")
        .arg("-f").arg("1").arg("-l").arg("1")
        .arg(&src)
        .output()
        .ok()?;
    let _ = std::fs::remove_dir_all(&dir);
    let text = String::from_utf8_lossy(&out.stdout);
    for line in text.lines() {
        if let Some(size) = line.strip_prefix("Page") {
            if let Some(rest) = size.trim().strip_prefix("size:") {
                let mut nums = rest.split_whitespace().filter_map(|t| t.parse::<f64>().ok());
                if let (Some(w), Some(h)) = (nums.next(), nums.next()) {
                    return Some((w, h));
                }
            }
        }
    }
    None
}

/// Errors if `pdftoppm` (poppler-utils) is not on PATH.
pub fn render_pages_png(bytes: &[u8], dpi: u32) -> Result<Vec<Vec<u8>>> {
    render_pages_png_opts(bytes, dpi, true)
}

/// `clamp_budget=false` skips the 0.4 MP clamp — that budget was tuned for
/// PaddleOCR-VL's reading window; Indra resizes to a 1024-long-side input
/// itself and expects ~200 dpi source pixels (its own pipeline renders via
/// pdfium at dpi/72). Clamping starves it (~64 dpi on A4).
pub fn render_pages_png_opts(bytes: &[u8], dpi: u32, clamp_budget: bool) -> Result<Vec<Vec<u8>>> {
    use std::process::Command;
    if dpi == 0 || dpi > 600 {
        anyhow::bail!("dpi must be 1..=600, got {dpi}");
    }
    // backend selection: the pdfium library first when installed (engine-
    // faithful pixels — Indra's own pipeline renders via pdfium), then
    // poppler (pdftoppm), then an actionable error naming every fix.
    #[cfg(feature = "render-pdfium")]
    if pdfium_lib_path().is_some() {
        if let Ok(pages) = render_pages_png_pdfium(bytes, dpi, clamp_budget) {
            return Ok(pages);
        }
        // pdfium unavailable/failed: fall through to poppler
    }
    if !pdftoppm_available() {
        anyhow::bail!(
            "no PDF renderer available. Install one of:\n  \
             • poppler-utils   (apt install poppler-utils / brew install poppler)\n  \
             • pdfium library  (office_parser::formats::pdf::install_pdfium()? — auto-download, \
               then build with --features render-pdfium)"
        );
    }
    // megapixel clamp (dev.to/serhii_kalyna_730b636889c pattern, budget ours):
    // probe the page size in points, back the dpi off by sqrt(budget/mp) —
    // pixels scale with dpi² — so dense pages land inside the OCR window
    // instead of yielding empty grids (PaddleOCR-VL, 2026-10-06).
    let dpi = if clamp_budget {
        match page_size_pts(bytes) {
            Some((w, h)) => clamp_dpi_to_budget(w, h, dpi),
            None => dpi, // probe failed: caller's dpi stands, downstream errors if fatal
        }
    } else {
        dpi
    };
    static RENDER_SEQ: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
    let dir = std::env::temp_dir().join(format!(
        "office-parser-render-{}-{}",
        std::process::id(),
        RENDER_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
    ));
    std::fs::create_dir_all(&dir)?;
    let result = (|| -> Result<Vec<Vec<u8>>> {
        let src = dir.join("doc.pdf");
        std::fs::write(&src, bytes)?;
        let stem = dir.join("page");
        let status = Command::new("pdftoppm")
            .args(["-png", "-r", &dpi.to_string()])
            .arg(&src)
            .arg(&stem)
            .status()
            .map_err(|e| anyhow::anyhow!("pdftoppm not found (install poppler-utils): {e}"))?;
        if !status.success() {
            anyhow::bail!("pdftoppm failed: {status}");
        }
        let mut pages: Vec<(String, Vec<u8>)> = std::fs::read_dir(&dir)?
            .filter_map(|entry| entry.ok())
            .filter_map(|entry| {
                let name = entry.file_name().into_string().ok()?;
                (entry.path().extension()?.to_str()? == "png" && name.starts_with("page-"))
                    .then(|| std::fs::read(entry.path()).ok().map(|b| (name, b)))?
            })
            .collect();
        pages.sort_by(|a, b| a.0.cmp(&b.0)); // pdftoppm zero-pads to the page-count width
        Ok(pages.into_iter().map(|(_, bytes)| bytes).collect())
    })();
    let _ = std::fs::remove_dir_all(&dir);
    result
}

pub fn parse(bytes: &[u8]) -> crate::Result<crate::Document> {
    let parsed =
        parse_pdf_full(bytes).map_err(|e| crate::Error::from_parse(crate::Format::Pdf, e))?;
    let quality = classify_pdf_text_quality(&parsed.blocks);
    let office = super::ParsedOfficeDocument {
        blocks: parsed.blocks,
        images: parsed.images,
        metadata_json: parsed.metadata_json,
    };
    let mut doc = super::finalize(crate::Format::Pdf, office);
    doc.metadata.pdf_text_quality = Some(quality);
    Ok(doc)
}

/// Alphanumeric-character count below which a page counts as image-like
/// (same threshold as the document-level `ImageOnly` classification, applied
/// per page).
pub const PAGE_TEXT_THRESHOLD: usize = 50;

/// Per-page OCR: an engine that turns one rendered page PNG into blocks.
pub trait PageOcrEngine {
    fn ocr_page_png(&mut self, page_png: &[u8]) -> Result<(Vec<Block>, Option<f64>), String>;
    fn name(&self) -> &'static str;
}

/// Router with default thresholds (50 chars, no force).
pub fn parse_with_page_ocr(
    bytes: &[u8],
    engine: &mut dyn PageOcrEngine,
    dpi: u32,
) -> crate::Result<crate::Document> {
    parse_with_page_ocr_opts(bytes, engine, dpi, PAGE_TEXT_THRESHOLD, false)
}

/// textlayer.py trust verdicts per page: Some(reason) = untrust the text
/// layer. Invisible render-mode-3 text = OCR layer over a scan; bad
/// characters (replacement / private-use / control) = broken font map.
/// Requires the pdfium feature; returns None (skip checks) without it.
#[cfg(feature = "render-pdfium")]
fn pdfium_text_layer_reasons(bytes: &[u8]) -> Option<Vec<Option<String>>> {
    use pdfium_render::prelude::*;
    let pdfium = pdfium_handle()?;
    let pdfium = pdfium.lock().ok()?;
    let doc = pdfium.load_pdf_from_byte_slice(bytes, None).ok()?;
    let mut out = Vec::new();
    for page in doc.pages().iter() {
        let text = match page.text() {
            Ok(t) => t,
            Err(_) => {
                out.push(Some("no text layer".to_string()));
                continue;
            }
        };
        let mut total = 0usize;
        let mut invisible = 0usize;
        let mut bad = 0usize;
        {
            for ch in text.chars().iter() {
                total += 1;
                match ch.unicode_char() {
                    Some(c) => {
                        let bad_char = c == '\u{FFFD}'
                            || ('\u{E000}'..='\u{F8FF}').contains(&c)
                            || ('\u{FFF0}'..='\u{FFFF}').contains(&c)
                            || (c.is_control() && c != '\t' && c != '\n' && c != '\r');
                        if bad_char {
                            bad += 1;
                        }
                    }
                    None => bad += 1,
                }
                if ch.render_mode().ok() == Some(PdfPageTextRenderMode::Invisible) {
                    invisible += 1;
                }
            }
        }
        let verdict = if total == 0 {
            Some("no text layer".to_string())
        } else if invisible as f32 / total as f32 > 0.05 {
            Some(format!("invisible text {}% (OCR layer)", invisible * 100 / total))
        } else if bad as f32 / total as f32 > 0.005 {
            Some(format!("broken font map ({}% bad chars)", bad * 100 / total))
        } else {
            None
        };
        out.push(verdict);
    }
    Some(out)
}

/// Per-page alphanumeric char counts (extraction order == page order).
fn page_char_counts(parsed: &ParsedPdfDocument) -> Vec<usize> {
    parsed
        .page_block_ranges
        .iter()
        .map(|r| count_alphanumeric(&parsed.blocks[r.block_start..r.block_end]))
        .collect()
}

/// Parse a PDF and route image-like pages through a per-page OCR engine while
/// keeping the text layer for text-rich pages — mixed documents (a digital
/// compilation with scanned inserts) get every page recovered.
pub fn parse_with_page_ocr_opts(
    bytes: &[u8],
    engine: &mut dyn PageOcrEngine,
    dpi: u32,
    min_chars: usize,
    force: bool,
) -> crate::Result<crate::Document> {
    parse_with_page_ocr_inner(bytes, &mut Some(engine), dpi, min_chars, force).map(|(doc, _)| doc)
}

/// Split a PDF into one [`crate::Document`] per page (1 page = 1 document).
/// `engine` enables OCR routing for image-like pages; `None` keeps the text
/// layer only. The combined document is returned alongside the pages.
pub fn parse_pdf_split(
    bytes: &[u8],
    mut engine: Option<&mut dyn PageOcrEngine>,
    dpi: u32,
    min_chars: usize,
    force: bool,
) -> crate::Result<(crate::Document, Vec<(u32, crate::Document)>)> {
    parse_with_page_ocr_inner(bytes, &mut engine, dpi, min_chars, force)
}

fn parse_with_page_ocr_inner<'e>(
    bytes: &[u8],
    engine: &mut Option<&'e mut (dyn PageOcrEngine + 'e)>,
    dpi: u32,
    min_chars: usize,
    force: bool,
) -> crate::Result<(
    crate::Document,
    Vec<(u32, crate::Document)>,
)> {
    let mut parsed =
        parse_pdf_full(bytes).map_err(|e| crate::Error::from_parse(crate::Format::Pdf, e))?;
    let quality = classify_pdf_text_quality(&parsed.blocks);
    let counts = page_char_counts(&parsed);
    // Text-layer trust checks (port of indra_ocr textlayer.py): invisible
    // render-mode-3 text = OCR layer over a scan; replacement/private-use/
    // control chars = broken font map. Either verdict sends the page to OCR
    // even when the char count looks rich. (no-ink check pending: needs
    // char-box geometry.)
    let trust_reasons: Vec<Option<String>> = if force {
        Vec::new()
    } else {
        pdfium_text_layer_reasons(bytes).unwrap_or_default()
    };
    let mut page_reasons: Vec<serde_json::Value> = Vec::new();
    let ocr_pages: Vec<usize> = counts
        .iter()
        .enumerate()
        .filter(|(i, c)| {
            if force || **c < min_chars {
                return true;
            }
            match trust_reasons.get(*i).and_then(|r| r.as_ref()) {
                Some(reason) => {
                    page_reasons.push(serde_json::json!({"page": i + 1, "reason": reason}));
                    true
                }
                None => false,
            }
        })
        .map(|(i, _)| i)
        .collect();

    
        #[allow(unused_assignments)]
    let mut ocr_meta = serde_json::json!({
        "engine": engine.as_ref().map(|e| e.name()).unwrap_or("none"),
        "pages_ocr": 0,
    });

    let mut pages_from_merge: Vec<Vec<Block>> = Vec::new();
    if !ocr_pages.is_empty() && engine.is_some() {
        // Indra wants full-resolution source pixels (see render_pages_png_opts)
        eprintln!(
            "office-parser: rendering {} page(s) at {} dpi...",
            counts.len(),
            dpi
        );
        let rendered = render_pages_png_opts(bytes, dpi, false)?;
        let engine_name = engine.as_deref_mut().map(|e| e.name()).unwrap_or("none").to_string();
        eprintln!(
            "office-parser: OCR routing {} page(s) of {} through {}",
            ocr_pages.len(),
            counts.len(),
            engine_name
        );
        let mut merged: Vec<Block> = Vec::new();
        let mut low_conf = Vec::new();
        for (page_pos, range) in parsed.page_block_ranges.iter().enumerate() {
            let mut page_vec: Vec<Block> = Vec::new();
            if ocr_pages.contains(&page_pos) {
                let png = rendered.get(page_pos).ok_or_else(|| {
                    crate::Error::from_parse(
                        crate::Format::Pdf,
                        anyhow::anyhow!("page {} was not rendered", page_pos + 1),
                    )
                })?;
                eprintln!(
                    "office-parser: ocr page {}/{}",
                    page_pos + 1,
                    counts.len()
                );
                let Some(engine_ref) = engine.as_deref_mut() else {
                    unreachable!("engine present when ocr_pages non-empty")
                };
                let (mut blocks, conf) = engine_ref
                    .ocr_page_png(png)
                    .map_err(|e| crate::Error::from_parse(crate::Format::Pdf, anyhow::anyhow!(e)))?;
                if let Some(conf) = conf {
                    if conf < 0.60 {
                        low_conf.push(serde_json::json!(page_pos));
                    }
                }
                page_vec.append(&mut blocks);
                ocr_meta["pages_ocr"] = serde_json::json!(ocr_meta["pages_ocr"].as_u64().unwrap_or(0) + 1);
            } else {
                page_vec.extend(parsed.blocks[range.block_start..range.block_end].iter().cloned());
            }
            let page_start = merged.len();
            let mut local: Vec<Block> = Vec::new();
            for (j, mut b) in page_vec.into_iter().enumerate() {
                set_block_index(&mut b, page_start + j);
                merged.push(b.clone());
                local.push(b);
            }
            pages_from_merge.push(local);
        }
        parsed.blocks = merged;
        ocr_meta["low_confidence_pages"] = serde_json::Value::Array(low_conf);

        if !page_reasons.is_empty() {
            ocr_meta["page_reasons"] = serde_json::Value::Array(page_reasons);
        }
    }

    let images = parsed.images;
    let metadata_json = parsed.metadata_json;
    let office = super::ParsedOfficeDocument {
        blocks: parsed.blocks,
        images: images.clone(),
        metadata_json: metadata_json.clone(),
    };
    let mut doc = super::finalize(crate::Format::Pdf, office);
    doc.metadata.pdf_text_quality = Some(quality);
    if ocr_meta["pages_ocr"].as_u64().unwrap_or(0) > 0 {
        doc.metadata.extra["ocr"] = ocr_meta;
    }

    // 1 page = 1 document (page-tagged via metadata.extra["page"])
    let mut pages = Vec::new();
    for (pos, r) in parsed.page_block_ranges.iter().enumerate() {
        let page_blocks: Vec<Block> = if pages_from_merge.is_empty() {
            doc.blocks[r.block_start..r.block_end].to_vec()
        } else {
            pages_from_merge.get(pos).cloned().unwrap_or_default()
        };
        let office = super::ParsedOfficeDocument {
            blocks: page_blocks,
            images: images.clone(),
            metadata_json: metadata_json.clone(),
        };
        let mut page_doc = super::finalize(crate::Format::Pdf, office);
        page_doc.metadata.pdf_text_quality = Some(quality);
        page_doc.metadata.extra["page"] = serde_json::json!(r.page_number);
        if doc.metadata.extra.get("ocr").is_some() {
            page_doc.metadata.extra["ocr"] = doc.metadata.extra["ocr"].clone();
        }
        pages.push((r.page_number, page_doc));
    }
    Ok((doc, pages))
}

fn set_block_index(b: &mut Block, i: usize) {
    match b {
        Block::Heading { block_index, .. }
        | Block::Paragraph { block_index, .. }
        | Block::List { block_index, .. }
        | Block::Table { block_index, .. }
        | Block::Image { block_index, .. }
        | Block::Link { block_index, .. }
        | Block::Note { block_index, .. } => *block_index = i,
    }
}

#[cfg(test)]
mod tests {
    use super::cleanup_pdf_text;

    #[test]
    fn clamp_dpi_backs_off_by_sqrt() {
        // A4 (612x792 pts): @300 dpi = 8.4 MP → factor 0.218 → 65 → floor 72
        assert_eq!(super::clamp_dpi_to_budget(612.0, 792.0, 300), 72);
        // @200 dpi = 3.74 MP → factor 0.327 → 65 → floor 72
        assert_eq!(super::clamp_dpi_to_budget(612.0, 792.0, 200), 72);
        // 200x200 pts @300 dpi = 0.69 MP → over 0.4 budget → factor 0.759 → 227
        assert_eq!(super::clamp_dpi_to_budget(200.0, 200.0, 300), 227);
        // tiny page under budget: unchanged
        assert_eq!(super::clamp_dpi_to_budget(100.0, 100.0, 300), 300);
    }

    #[test]
    fn render_pages_png_rasterizes_via_pdftoppm() {
        // minimal one-page PDF; poppler repairs the placeholder xref
        let pdf = b"%PDF-1.4\n1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj\n3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >> endobj\ntrailer << /Root 1 0 R /Size 4 >>\n%%EOF".to_vec();
        let pages = super::render_pages_png(&pdf, 100).expect("renders");
        assert!(!pages.is_empty());
        for page in &pages {
            assert!(page.starts_with(b"\x89PNG"), "expected PNG magic");
        }
    }

    #[test]
    fn pdf_cleanup_dehyphenates_and_splits_glued_year_citations() {
        let s = "hydro-\n gen 203070 202312 20250131";
        let out = cleanup_pdf_text(s);
        assert!(out.contains("hydrogen"));
        assert!(out.contains("2030 70"));
        assert!(out.contains("2023-12"));
        assert!(out.contains("2025-01-31"));
    }
}
