use anyhow::{Context, Result, anyhow};
use base64::Engine;
use encoding_rs::Encoding;

use crate::document_ast::{Block, Cell, LinkKind, ListItem, SourceSpan};

use super::{ParsedImage, ParsedOfficeDocument, sha256_hex};

fn sniff_image_mime(bytes: &[u8]) -> Option<&'static str> {
    if bytes.len() >= 8 && bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        return Some("image/png");
    }
    if bytes.len() >= 3 && bytes[0] == 0xff && bytes[1] == 0xd8 && bytes[2] == 0xff {
        return Some("image/jpeg");
    }
    if bytes.len() >= 6 && (bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a")) {
        return Some("image/gif");
    }
    if bytes.len() >= 12 && bytes.starts_with(b"RIFF") && &bytes[8..12] == b"WEBP" {
        return Some("image/webp");
    }
    // Common WMF headers:
    // - Placeable WMF: D7 CD C6 9A
    // - Standard WMF: 01 00 09 00
    if bytes.len() >= 4
        && (bytes.starts_with(b"\xD7\xCD\xC6\x9A") || bytes.starts_with(b"\x01\x00\x09\x00"))
    {
        return Some("image/wmf");
    }
    None
}

fn decode_data_image_uri(uri: &str) -> Result<Option<(Vec<u8>, String)>> {
    let u = uri.trim();
    if !u.to_ascii_lowercase().starts_with("data:image/") {
        return Ok(None);
    }
    let Some((meta, payload)) = u.split_once(',') else {
        return Ok(None);
    };
    let meta_l = meta.to_ascii_lowercase();
    if !meta_l.contains(";base64") {
        return Ok(None);
    }

    let mime = meta
        .trim_start_matches("data:")
        .split(';')
        .next()
        .unwrap_or("application/octet-stream")
        .trim()
        .to_string();

    let payload = payload.trim();
    if payload.is_empty() {
        return Ok(None);
    }
    let bytes = base64::engine::general_purpose::STANDARD
        .decode(payload)
        .context("decode data:image base64")?;
    Ok(Some((bytes, mime)))
}

fn extract_quoted_or_last_token(inst: &str) -> Option<String> {
    if let Some(first) = inst.find('"') {
        if let Some(second) = inst[first + 1..].find('"') {
            return Some(inst[first + 1..first + 1 + second].to_string());
        }
    }
    inst.split_whitespace()
        .rev()
        .find(|t| {
            t.contains("://")
                || t.starts_with("mailto:")
                || t.to_ascii_lowercase().starts_with("data:image/")
        })
        .map(|s| s.trim_matches('"').to_string())
}

fn push_para_link(out: &mut Vec<(String, Option<String>)>, url: String, txt: Option<String>) {
    if let Some((prev_url, prev_txt)) = out.last_mut() {
        if prev_url == &url {
            match (prev_txt.as_mut(), txt) {
                (Some(a), Some(b)) => a.push_str(&b),
                (None, Some(b)) => *prev_txt = Some(b),
                _ => {}
            }
            return;
        }
    }
    out.push((url, txt));
}

fn encoding_from_ansicpg(cp: i32) -> &'static Encoding {
    match cp {
        65001 => encoding_rs::UTF_8,

        1250 => encoding_rs::WINDOWS_1250,
        1251 => encoding_rs::WINDOWS_1251,
        1252 => encoding_rs::WINDOWS_1252,
        1253 => encoding_rs::WINDOWS_1253,
        1254 => encoding_rs::WINDOWS_1254,
        1255 => encoding_rs::WINDOWS_1255,
        1256 => encoding_rs::WINDOWS_1256,
        1257 => encoding_rs::WINDOWS_1257,
        1258 => encoding_rs::WINDOWS_1258,

        932 => encoding_rs::SHIFT_JIS,
        936 => encoding_rs::GBK,
        949 => encoding_rs::EUC_KR,
        950 => encoding_rs::BIG5,

        _ => encoding_rs::WINDOWS_1252,
    }
}

fn is_compressed_rtf(data: &[u8]) -> bool {
    if data.len() < 16 {
        return false;
    }
    let sig = &data[8..12];
    sig == b"LZFu" || sig == b"MELA"
}

fn decompress_rtf(data: &[u8]) -> Result<Vec<u8>> {
    if data.len() < 16 {
        return Err(anyhow!("compressed RTF header must be at least 16 bytes"));
    }

    let compressed_size = u32::from_le_bytes(data[0..4].try_into().unwrap_or([0u8; 4]));
    let raw_size = u32::from_le_bytes(data[4..8].try_into().unwrap_or([0u8; 4]));
    let compression_type: [u8; 4] = data[8..12].try_into().unwrap_or(*b"????");
    let crc32 = u32::from_le_bytes(data[12..16].try_into().unwrap_or([0u8; 4]));

    // Data after the 16-byte header.
    let payload = &data[16..];

    // Guard: if compressed_size exists, ensure it's coherent.
    if compressed_size as usize > data.len() {
        // Some writers set this incorrectly; don't hard-fail.
    }

    if &compression_type == b"MELA" {
        if crc32 != 0 {
            return Err(anyhow!("RTF MELA payload expects CRC32=0"));
        }
        let want = raw_size as usize;
        if payload.len() < want {
            return Err(anyhow!(
                "RTF MELA payload truncated (want {want} bytes, got {})",
                payload.len()
            ));
        }
        return Ok(payload[..want].to_vec());
    }

    if &compression_type != b"LZFu" {
        return Err(anyhow!(
            "unknown compressed RTF signature: {:?}",
            compression_type
        ));
    }

    let calculated = crc32fast::hash(payload);
    if calculated != crc32 {
        return Err(anyhow!(
            "compressed RTF CRC32 mismatch: expected {:#x}, got {:#x}",
            crc32,
            calculated
        ));
    }

    // Per MS-OXRTFCP: initial dictionary.
    const INIT_DICT: &[u8] = b"{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}\
{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArial\
Times New RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\
\\b\\i\\u\\tab\\tx";
    const INIT_DICT_SIZE: usize = 207;
    const MAX_DICT_SIZE: usize = 4096;

    let mut dict = vec![0u8; MAX_DICT_SIZE];
    dict[..INIT_DICT_SIZE].copy_from_slice(INIT_DICT);
    dict[INIT_DICT_SIZE..].fill(b' ');

    let mut write_offset = INIT_DICT_SIZE;
    let mut out: Vec<u8> = Vec::with_capacity(raw_size as usize);
    let mut pos = 0usize;

    while pos < payload.len() {
        let control = payload[pos];
        pos += 1;

        for bit in 0..8 {
            if pos >= payload.len() {
                break;
            }

            if (control & (1 << bit)) != 0 {
                if pos + 1 >= payload.len() {
                    break;
                }
                let token = u16::from_be_bytes([payload[pos], payload[pos + 1]]);
                pos += 2;

                let offset = ((token >> 4) & 0x0fff) as usize;
                let length = (token & 0x0f) as usize;

                // End marker.
                if write_offset == offset {
                    return Ok(out);
                }

                let actual_length = length + 2;
                for step in 0..actual_length {
                    let read_offset = (offset + step) % MAX_DICT_SIZE;
                    let b = dict[read_offset];
                    out.push(b);
                    dict[write_offset] = b;
                    write_offset = (write_offset + 1) % MAX_DICT_SIZE;
                }
            } else {
                let b = payload[pos];
                pos += 1;
                out.push(b);
                dict[write_offset] = b;
                write_offset = (write_offset + 1) % MAX_DICT_SIZE;
            }
        }
    }

    Ok(out)
}

fn maybe_decompress(bytes: &[u8]) -> Result<Vec<u8>> {
    if is_compressed_rtf(bytes) {
        return decompress_rtf(bytes);
    }
    Ok(bytes.to_vec())
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Destination {
    Document,
    Skip,
    Pict,
    FieldInst,
    FieldRslt,
}

#[derive(Clone, Debug, Default)]
struct FieldCtx {
    inst: String,
    rslt: String,
    saw_raster_image: bool,
}

#[derive(Clone, Debug)]
struct PictCtx {
    mime_type: String,
    hex: String,
    raw: Vec<u8>,
}

#[derive(Clone, Debug)]
struct GroupCtx {
    destination: Destination,
    encoding: &'static Encoding,
    unicode_skip: i32,
    unicode_skip_pending: i32,
    star_pending: bool,
    in_nonshppict: bool,
    pntext_mode: bool,
    pntext_buf: String,
    field_started_here: bool,
    pict_started_here: bool,
}

impl Default for GroupCtx {
    fn default() -> Self {
        Self {
            destination: Destination::Document,
            encoding: encoding_rs::WINDOWS_1252,
            unicode_skip: 1,
            unicode_skip_pending: 0,
            star_pending: false,
            in_nonshppict: false,
            pntext_mode: false,
            pntext_buf: String::new(),
            field_started_here: false,
            pict_started_here: false,
        }
    }
}

fn is_hex_digit(b: u8) -> bool {
    (b'0'..=b'9').contains(&b) || (b'a'..=b'f').contains(&b) || (b'A'..=b'F').contains(&b)
}

fn hex_val(b: u8) -> Option<u8> {
    match b {
        b'0'..=b'9' => Some(b - b'0'),
        b'a'..=b'f' => Some(b - b'a' + 10),
        b'A'..=b'F' => Some(b - b'A' + 10),
        _ => None,
    }
}

fn decode_hex_string(s: &str) -> Vec<u8> {
    let bytes = s.as_bytes();
    let mut out = Vec::with_capacity(bytes.len() / 2);
    let mut i = 0usize;
    while i + 1 < bytes.len() {
        let hi = hex_val(bytes[i]);
        let lo = hex_val(bytes[i + 1]);
        if let (Some(hi), Some(lo)) = (hi, lo) {
            out.push((hi << 4) | lo);
            i += 2;
        } else {
            i += 1;
        }
    }
    out
}

fn push_byte_as_text(out: &mut String, enc: &'static Encoding, b: u8) {
    if b < 0x80 {
        out.push(b as char);
        return;
    }
    let buf = [b];
    let (cow, _, _) = enc.decode(&buf);
    out.push_str(&cow);
}

fn normalize_ws(s: &str) -> String {
    let mut out = String::new();
    let mut prev_space = false;
    for ch in s.chars() {
        let is_space = ch.is_whitespace() && ch != '\n';
        if is_space {
            if !prev_space {
                out.push(' ');
                prev_space = true;
            }
            continue;
        }
        prev_space = false;
        out.push(ch);
    }
    out.trim().to_string()
}

fn ordered_from_marker(marker: &str) -> bool {
    let m = marker.trim();
    if m.is_empty() {
        return false;
    }
    let has_alnum = m.chars().any(|c| c.is_ascii_alphanumeric());
    if !has_alnum {
        return false;
    }
    m.ends_with('.') || m.ends_with(')')
}

fn flush_pending_list(
    blocks: &mut Vec<Block>,
    next_block_index: &mut usize,
    pending_list: &mut Option<(bool, Vec<ListItem>)>,
) {
    if let Some((ordered, items)) = pending_list.take() {
        if !items.is_empty() {
            blocks.push(Block::List {
                block_index: *next_block_index,
                ordered,
                items,
                source: SourceSpan::default(),
            });
            *next_block_index += 1;
        }
    }
}

fn flush_table_if_any(
    blocks: &mut Vec<Block>,
    next_block_index: &mut usize,
    in_table: &mut bool,
    table_rows: &mut Vec<Vec<Cell>>,
    row_cells: &mut Vec<Cell>,
) {
    if !*in_table {
        return;
    }
    if !row_cells.is_empty() {
        table_rows.push(std::mem::take(row_cells));
    }
    if !table_rows.is_empty() {
        blocks.push(Block::Table {
            block_index: *next_block_index,
            rows: std::mem::take(table_rows),
            source: SourceSpan::default(),
        });
        *next_block_index += 1;
    }
    *in_table = false;
}

fn flush_paragraph(
    blocks: &mut Vec<Block>,
    next_block_index: &mut usize,
    cur_text: &mut String,
    cur_list_marker: &mut Option<String>,
    pending_list: &mut Option<(bool, Vec<ListItem>)>,
    cur_para_links: &mut Vec<(String, Option<String>)>,
    in_table: bool,
) {
    if in_table {
        cur_text.push('\n');
        return;
    }

    let text = normalize_ws(cur_text);
    cur_text.clear();

    let list_marker = cur_list_marker.take().unwrap_or_default();
    if !list_marker.trim().is_empty() {
        let ordered = ordered_from_marker(&list_marker);
        let item = ListItem {
            level: 0,
            text,
            source: SourceSpan::default(),
        };
        if let Some((ord, items)) = pending_list.as_mut() {
            if *ord == ordered {
                items.push(item);
            } else {
                flush_pending_list(blocks, next_block_index, pending_list);
                *pending_list = Some((ordered, vec![item]));
            }
        } else {
            *pending_list = Some((ordered, vec![item]));
        }

        // Links inside list items aren't represented in the AST yet.
        cur_para_links.clear();
        return;
    }

    if !text.is_empty() {
        flush_pending_list(blocks, next_block_index, pending_list);
        blocks.push(Block::Paragraph {
            block_index: *next_block_index,
            text,
            source: SourceSpan::default(),
        });
        *next_block_index += 1;

        for (url, txt) in cur_para_links.drain(..) {
            blocks.push(Block::Link {
                block_index: *next_block_index,
                url,
                text: txt,
                kind: LinkKind::Unknown,
                source: SourceSpan::default(),
            });
            *next_block_index += 1;
        }
    } else {
        flush_pending_list(blocks, next_block_index, pending_list);
        cur_para_links.clear();
    }
}

pub fn parse_rtf_full(bytes: &[u8]) -> Result<ParsedOfficeDocument> {
    let bytes = maybe_decompress(bytes).context("decompress RTF")?;
    if bytes.len() < 5 || !bytes.starts_with(b"{\\rtf") {
        return Err(anyhow!("not an RTF document"));
    }

    let mut blocks: Vec<Block> = Vec::new();
    let mut images: Vec<ParsedImage> = Vec::new();
    let mut next_block_index: usize = 0;

    let mut group_stack: Vec<GroupCtx> = vec![GroupCtx::default()];
    let mut field_stack: Vec<FieldCtx> = Vec::new();
    let mut pict_stack: Vec<PictCtx> = Vec::new();

    let mut cur_text = String::new();
    let mut cur_para_links: Vec<(String, Option<String>)> = Vec::new();
    let mut cur_list_marker: Option<String> = None;

    let mut in_table = false;
    let mut table_rows: Vec<Vec<Cell>> = Vec::new();
    let mut row_cells: Vec<Cell> = Vec::new();

    let mut pending_list: Option<(bool, Vec<ListItem>)> = None;

    let mut pos = 0usize;
    while pos < bytes.len() {
        let b = bytes[pos];
        pos += 1;

        match b {
            b'{' => {
                group_stack.push(group_stack.last().cloned().unwrap_or_default());
                if let Some(top) = group_stack.last_mut() {
                    top.field_started_here = false;
                    top.pict_started_here = false;
                    top.star_pending = false;
                    top.pntext_mode = false;
                    top.pntext_buf.clear();
                }
            }
            b'}' => {
                let ended = group_stack.pop().unwrap_or_default();

                if ended.pict_started_here {
                    if let Some(pict) = pict_stack.pop() {
                        let data = if !pict.raw.is_empty() {
                            pict.raw
                        } else {
                            decode_hex_string(&pict.hex)
                        };
                        if !data.is_empty() {
                            let mime = if pict.mime_type == "application/octet-stream" {
                                sniff_image_mime(&data)
                                    .map(|s| s.to_string())
                                    .unwrap_or_else(|| pict.mime_type.clone())
                            } else {
                                pict.mime_type.clone()
                            };

                            // Word often emits the same image twice: a raster (png/jpg) under
                            // {\*\shppict ...} and a WMF fallback under {\nonshppict ...}. Prefer
                            // raster when both are present within the same field.
                            let skip_image = ended.in_nonshppict
                                && mime == "image/wmf"
                                && field_stack.last().is_some_and(|f| f.saw_raster_image);

                            if !skip_image {
                                let is_raster = matches!(
                                    mime.as_str(),
                                    "image/png"
                                        | "image/jpeg"
                                        | "image/gif"
                                        | "image/webp"
                                        | "image/bmp"
                                );
                                if is_raster {
                                    if let Some(f) = field_stack.last_mut() {
                                        f.saw_raster_image = true;
                                    }
                                }
                                let hash = sha256_hex(&data);
                                images.push(ParsedImage {
                                    id: format!("sha256:{hash}"),
                                    bytes: data,
                                    mime_type: mime.clone(),
                                    filename: None,
                                });
                                blocks.push(Block::Image {
                                    block_index: next_block_index,
                                    id: format!("sha256:{hash}"),
                                    filename: None,
                                    content_type: Some(mime),
                                    alt: None,
                                    source: SourceSpan::default(),
                                });
                                next_block_index += 1;
                            }
                        }
                    }
                }

                if ended.field_started_here {
                    if let Some(field) = field_stack.pop() {
                        let inst = field.inst;
                        let rslt = normalize_ws(&field.rslt);
                        let inst_upper = inst.to_ascii_uppercase();
                        if inst_upper.contains("HYPERLINK") || inst_upper.contains("INCLUDEPICTURE")
                        {
                            let Some(url) = extract_quoted_or_last_token(&inst)
                                .map(|u| u.trim().to_string())
                                .filter(|u| !u.is_empty())
                            else {
                                continue;
                            };

                            // If someone embedded a data:image URI in a field, decode it to an image
                            // instead of emitting giant extracted_text.
                            if let Some((bytes, mime)) = decode_data_image_uri(&url)? {
                                if !bytes.is_empty() {
                                    let mime = sniff_image_mime(&bytes)
                                        .map(|s| s.to_string())
                                        .unwrap_or(mime);
                                    let hash = sha256_hex(&bytes);
                                    images.push(ParsedImage {
                                        id: format!("sha256:{hash}"),
                                        bytes,
                                        mime_type: mime.clone(),
                                        filename: None,
                                    });
                                    blocks.push(Block::Image {
                                        block_index: next_block_index,
                                        id: format!("sha256:{hash}"),
                                        filename: None,
                                        content_type: Some(mime),
                                        alt: None,
                                        source: SourceSpan::default(),
                                    });
                                    next_block_index += 1;
                                }
                                continue;
                            }

                            let txt = if inst_upper.contains("INCLUDEPICTURE") {
                                None
                            } else if rslt.is_empty() {
                                None
                            } else {
                                Some(rslt)
                            };

                            push_para_link(&mut cur_para_links, url, txt);
                        }
                    }
                }

                if group_stack.is_empty() {
                    group_stack.push(GroupCtx::default());
                }
            }
            b'\n' | b'\r' => {}
            b'\\' => {
                if pos >= bytes.len() {
                    break;
                }
                let c = bytes[pos];
                pos += 1;

                match c {
                    b'*' => {
                        let top = group_stack.last_mut().unwrap();
                        top.star_pending = true;
                    }
                    b'{' => {
                        let dest = group_stack
                            .last()
                            .map(|g| g.destination)
                            .unwrap_or(Destination::Skip);
                        if dest == Destination::Document || dest == Destination::FieldRslt {
                            cur_text.push('{');
                        }
                    }
                    b'}' => {
                        let dest = group_stack
                            .last()
                            .map(|g| g.destination)
                            .unwrap_or(Destination::Skip);
                        if dest == Destination::Document || dest == Destination::FieldRslt {
                            cur_text.push('}');
                        }
                    }
                    b'\\' => {
                        let dest = group_stack
                            .last()
                            .map(|g| g.destination)
                            .unwrap_or(Destination::Skip);
                        if dest == Destination::Document || dest == Destination::FieldRslt {
                            cur_text.push('\\');
                        }
                    }
                    b'~' => {
                        let dest = group_stack
                            .last()
                            .map(|g| g.destination)
                            .unwrap_or(Destination::Skip);
                        if dest == Destination::Document || dest == Destination::FieldRslt {
                            cur_text.push(' ');
                        }
                    }
                    b'-' | b'_' => {
                        let dest = group_stack
                            .last()
                            .map(|g| g.destination)
                            .unwrap_or(Destination::Skip);
                        if dest == Destination::Document || dest == Destination::FieldRslt {
                            cur_text.push('-');
                        }
                    }
                    b'\n' => {}
                    b'\'' => {
                        // Hex escape: \'hh
                        let h1 = bytes.get(pos).copied().unwrap_or(b'0');
                        let h2 = bytes.get(pos + 1).copied().unwrap_or(b'0');
                        pos = (pos + 2).min(bytes.len());
                        if let (Some(hi), Some(lo)) = (hex_val(h1), hex_val(h2)) {
                            let val = (hi << 4) | lo;
                            let top = group_stack.last_mut().unwrap();
                            if top.unicode_skip_pending > 0 {
                                top.unicode_skip_pending -= 1;
                            } else {
                                match top.destination {
                                    Destination::Document | Destination::FieldRslt => {
                                        if top.pntext_mode {
                                            push_byte_as_text(
                                                &mut top.pntext_buf,
                                                top.encoding,
                                                val,
                                            );
                                        } else {
                                            push_byte_as_text(&mut cur_text, top.encoding, val);
                                        }
                                        if top.destination == Destination::FieldRslt {
                                            if let Some(f) = field_stack.last_mut() {
                                                push_byte_as_text(&mut f.rslt, top.encoding, val);
                                            }
                                        }
                                    }
                                    Destination::FieldInst => {
                                        if let Some(f) = field_stack.last_mut() {
                                            push_byte_as_text(&mut f.inst, top.encoding, val);
                                        }
                                    }
                                    Destination::Pict => {
                                        if let Some(p) = pict_stack.last_mut() {
                                            p.raw.push(val);
                                        }
                                    }
                                    Destination::Skip => {}
                                }
                            }
                        }
                    }
                    _ => {
                        // Parse control word: letters
                        let mut word = Vec::new();
                        word.push(c);
                        while pos < bytes.len() {
                            let ch = bytes[pos];
                            if (b'a'..=b'z').contains(&ch) || (b'A'..=b'Z').contains(&ch) {
                                word.push(ch);
                                pos += 1;
                            } else {
                                break;
                            }
                        }

                        // Optional numeric parameter.
                        let mut neg = false;
                        if pos < bytes.len() && bytes[pos] == b'-' {
                            neg = true;
                            pos += 1;
                        }
                        let mut num: Option<i32> = None;
                        let mut n: i32 = 0;
                        let mut saw = false;
                        while pos < bytes.len() {
                            let ch = bytes[pos];
                            if (b'0'..=b'9').contains(&ch) {
                                saw = true;
                                n = n.saturating_mul(10).saturating_add((ch - b'0') as i32);
                                pos += 1;
                            } else {
                                break;
                            }
                        }
                        if saw {
                            num = Some(if neg { -n } else { n });
                        }

                        // Control word delimiter space is swallowed.
                        if pos < bytes.len() && bytes[pos] == b' ' {
                            pos += 1;
                        }

                        let kw = String::from_utf8_lossy(&word).to_ascii_lowercase();
                        let top = group_stack.last_mut().unwrap();

                        if top.star_pending {
                            // \* introduces an optional destination: if we don't recognize it, skip.
                            if kw != "fldinst" {
                                top.destination = Destination::Skip;
                            }
                            top.star_pending = false;
                        }

                        match kw.as_str() {
                            "fonttbl" | "colortbl" | "stylesheet" | "info" => {
                                top.destination = Destination::Skip;
                            }
                            "ansicpg" => {
                                if let Some(v) = num {
                                    top.encoding = encoding_from_ansicpg(v);
                                }
                            }
                            "uc" => {
                                if let Some(v) = num {
                                    top.unicode_skip = v.max(0);
                                }
                            }
                            "u" => {
                                if let Some(v) = num {
                                    let mut code = v;
                                    if code < 0 {
                                        code += 65536;
                                    }
                                    if let Some(ch) = char::from_u32(code as u32) {
                                        match top.destination {
                                            Destination::Document | Destination::FieldRslt => {
                                                if top.pntext_mode {
                                                    top.pntext_buf.push(ch);
                                                } else {
                                                    cur_text.push(ch);
                                                }
                                                if top.destination == Destination::FieldRslt {
                                                    if let Some(f) = field_stack.last_mut() {
                                                        f.rslt.push(ch);
                                                    }
                                                }
                                            }
                                            Destination::FieldInst => {
                                                if let Some(f) = field_stack.last_mut() {
                                                    f.inst.push(ch);
                                                }
                                            }
                                            _ => {}
                                        }
                                    }
                                    top.unicode_skip_pending = top.unicode_skip;
                                }
                            }
                            "par" => {
                                flush_paragraph(
                                    &mut blocks,
                                    &mut next_block_index,
                                    &mut cur_text,
                                    &mut cur_list_marker,
                                    &mut pending_list,
                                    &mut cur_para_links,
                                    in_table,
                                );
                            }
                            "line" => {
                                if in_table {
                                    cur_text.push('\n');
                                } else {
                                    cur_text.push(' ');
                                }
                            }
                            "tab" => {
                                if top.pntext_mode {
                                    top.pntext_mode = false;
                                    cur_list_marker = Some(normalize_ws(&top.pntext_buf));
                                    top.pntext_buf.clear();
                                } else {
                                    cur_text.push('\t');
                                }
                            }
                            "pntext" | "listtext" => {
                                top.pntext_mode = true;
                                top.pntext_buf.clear();
                            }
                            "trowd" => {
                                flush_pending_list(
                                    &mut blocks,
                                    &mut next_block_index,
                                    &mut pending_list,
                                );
                                in_table = true;
                                if !row_cells.is_empty() {
                                    table_rows.push(std::mem::take(&mut row_cells));
                                }
                                cur_text.clear();
                            }
                            "cell" => {
                                if in_table {
                                    let txt = normalize_ws(&cur_text);
                                    cur_text.clear();
                                    row_cells.push(Cell {
                                        text: txt,
                                        colspan: 1,
                                        rowspan: 1,
                                    });
                                }
                            }
                            "row" => {
                                if in_table {
                                    if !cur_text.trim().is_empty() {
                                        let txt = normalize_ws(&cur_text);
                                        cur_text.clear();
                                        row_cells.push(Cell {
                                            text: txt,
                                            colspan: 1,
                                            rowspan: 1,
                                        });
                                    }
                                    if !row_cells.is_empty() {
                                        table_rows.push(std::mem::take(&mut row_cells));
                                    }
                                }
                            }
                            "pard" => {
                                if in_table {
                                    flush_table_if_any(
                                        &mut blocks,
                                        &mut next_block_index,
                                        &mut in_table,
                                        &mut table_rows,
                                        &mut row_cells,
                                    );
                                }
                            }
                            "field" => {
                                field_stack.push(FieldCtx::default());
                                top.field_started_here = true;
                            }
                            "fldinst" => {
                                top.destination = Destination::FieldInst;
                            }
                            "fldrslt" => {
                                top.destination = Destination::FieldRslt;
                            }
                            "nonshppict" => {
                                top.in_nonshppict = true;
                            }
                            "pict" => {
                                pict_stack.push(PictCtx {
                                    mime_type: "application/octet-stream".to_string(),
                                    hex: String::new(),
                                    raw: Vec::new(),
                                });
                                top.destination = Destination::Pict;
                                top.pict_started_here = true;
                            }
                            "pngblip" => {
                                if let Some(p) = pict_stack.last_mut() {
                                    p.mime_type = "image/png".to_string();
                                }
                            }
                            "jpegblip" | "jpgblip" => {
                                if let Some(p) = pict_stack.last_mut() {
                                    p.mime_type = "image/jpeg".to_string();
                                }
                            }
                            "gifblip" => {
                                if let Some(p) = pict_stack.last_mut() {
                                    p.mime_type = "image/gif".to_string();
                                }
                            }
                            "wmetafile8" | "wmfblip" => {
                                if let Some(p) = pict_stack.last_mut() {
                                    p.mime_type = "image/wmf".to_string();
                                }
                            }
                            "emfblip" => {
                                if let Some(p) = pict_stack.last_mut() {
                                    p.mime_type = "image/emf".to_string();
                                }
                            }
                            "dibitmap0" | "dibitmap" => {
                                if let Some(p) = pict_stack.last_mut() {
                                    p.mime_type = "image/bmp".to_string();
                                }
                            }
                            "bin" => {
                                if top.destination == Destination::Pict {
                                    let n = num.unwrap_or(0).max(0) as usize;
                                    let end = (pos + n).min(bytes.len());
                                    if let Some(p) = pict_stack.last_mut() {
                                        p.raw.extend_from_slice(&bytes[pos..end]);
                                    }
                                    pos = end;
                                }
                            }
                            _ => {}
                        }
                    }
                }
            }
            _ => {
                let top = group_stack.last_mut().unwrap();
                if top.unicode_skip_pending > 0 {
                    top.unicode_skip_pending -= 1;
                    continue;
                }

                match top.destination {
                    Destination::Skip => {}
                    Destination::Pict => {
                        if is_hex_digit(b) {
                            if let Some(p) = pict_stack.last_mut() {
                                p.hex.push(b as char);
                            }
                        }
                    }
                    Destination::FieldInst => {
                        if let Some(f) = field_stack.last_mut() {
                            push_byte_as_text(&mut f.inst, top.encoding, b);
                        }
                    }
                    Destination::FieldRslt | Destination::Document => {
                        if top.pntext_mode {
                            push_byte_as_text(&mut top.pntext_buf, top.encoding, b);
                        } else {
                            push_byte_as_text(&mut cur_text, top.encoding, b);
                        }
                        if top.destination == Destination::FieldRslt {
                            if let Some(f) = field_stack.last_mut() {
                                push_byte_as_text(&mut f.rslt, top.encoding, b);
                            }
                        }
                    }
                }
            }
        }
    }

    if in_table && !cur_text.trim().is_empty() {
        let txt = normalize_ws(&cur_text);
        cur_text.clear();
        row_cells.push(Cell {
            text: txt,
            colspan: 1,
            rowspan: 1,
        });
    }

    flush_table_if_any(
        &mut blocks,
        &mut next_block_index,
        &mut in_table,
        &mut table_rows,
        &mut row_cells,
    );
    flush_pending_list(&mut blocks, &mut next_block_index, &mut pending_list);

    let rem = normalize_ws(&cur_text);
    if !rem.is_empty() {
        blocks.push(Block::Paragraph {
            block_index: next_block_index,
            text: rem,
            source: SourceSpan::default(),
        });
        next_block_index += 1;
    }
    cur_text.clear();

    for (url, txt) in cur_para_links.drain(..) {
        blocks.push(Block::Link {
            block_index: next_block_index,
            url,
            text: txt,
            kind: LinkKind::Unknown,
            source: SourceSpan::default(),
        });
        next_block_index += 1;
    }

    Ok(ParsedOfficeDocument {
        blocks,
        images,
        metadata_json: serde_json::json!({"format": "rtf"}),
    })
}

pub fn parse(bytes: &[u8]) -> crate::Result<crate::Document> {
    let parsed =
        parse_rtf_full(bytes).map_err(|e| crate::Error::from_parse(crate::Format::Rtf, e))?;
    Ok(super::finalize(crate::Format::Rtf, parsed))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rtf_data_image_uri_in_field_is_decoded_to_image_block() {
        // 1x1 PNG.
        let b64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMB/6qz0gAAAABJRU5ErkJggg==";
        let rtf = format!(
            "{{\\rtf1\\ansi{{\\field{{\\*\\fldinst HYPERLINK \"data:image/png;base64,{b64}\"}}{{\\fldrslt image}}}}}}"
        );

        let parsed = parse_rtf_full(rtf.as_bytes()).expect("parse rtf");
        assert_eq!(parsed.images.len(), 1);
        assert!(parsed.images[0].mime_type.starts_with("image/"));
        assert!(
            parsed
                .blocks
                .iter()
                .any(|b| matches!(b, Block::Image { .. }))
        );
    }
}
