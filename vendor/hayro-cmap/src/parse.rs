use alloc::boxed::Box;
use alloc::vec::Vec;

use hayro_postscript::{Number, Scanner, Token};

#[cfg(feature = "embed-cmaps")]
use crate::bcmap;
use crate::{
    BfRange, CMap, CMapName, CharacterCollection, CidFamily, CidRange, CodespaceRange,
    MAX_NESTING_DEPTH, Metadata, PartitionedRanges, Range, WritingMode,
};

struct Context<F> {
    buf: Vec<u8>,
    get_cmap: F,
}

pub(crate) pub(crate) fn parse_inner<'a>(
    data: &[u8],
    get_cmap: impl Fn(CMapName<'_>) -> Option<&'a [u8]> + Clone + 'a,
    depth: u32,
) -> Option<CMap> {
    // Prevent stack overflow for malicious cmap files or circular references.
    if depth >= MAX_NESTING_DEPTH {
        return None;
    }

    // Check if it's in our custom embedded cmap format.
    #[cfg(feature = "embed-cmaps")]
    if data.starts_with(b"bcmap") {
        return bcmap::parse(data, get_cmap, depth);
    }

    let mut scanner = Scanner::new(data);
    let mut ctx = Context {
        buf: Vec::new(),
        get_cmap,
    };
    let mut _codespace_ranges = Vec::new();
    let mut ranges = PartitionedRanges::new();
    let mut notdef_ranges = PartitionedRanges::new();
    let mut bf_entries = Vec::new();
    let mut base = None;

    let mut registry = None;
    let mut ordering = None;
    let mut supplement = None;
    let mut cmap_name = None;
    let mut writing_mode = None;
    let mut last = None;

    // PdfCraft patch: track whether the scan was cut short. A CMap with an unparseable tail,
    // including one that ends inside a `begin…` block, keeps the entries read so far (the block
    // parsers push into the shared lists as they go) instead of being discarded wholesale.
    let mut truncated = false;

    // PdfCraft patch: track a `usecmap` reference that could not be resolved or parsed. When the
    // stream yields no mappings of its own, the caller must still see failure, so its encoding
    // fallback paths run exactly as they did before the tolerant-usecmap change.
    let mut usecmap_failed = false;

    while !scanner.at_end() {
        let Ok(obj) = scanner.parse_token() else {
            truncated = true;
            break;
        };
        let prev = last.replace(obj.clone());

        let Token::Name(name) = &obj else { continue };

        if name.is_literal() {
            match name.as_str() {
                // Strictly speaking, Registry and Ordering should be strings,
                // but some PDF generators emit names instead, so be lenient
                // and try both.
                Some("Registry") => {
                    registry = parse_string_or_name(&mut scanner);
                }
                Some("Ordering") => {
                    ordering = parse_string_or_name(&mut scanner);
                }
                Some("Supplement") => {
                    supplement = parse_integer(&mut scanner).and_then(|n| i32::try_from(n).ok());
                }
                Some("CMapName") => {
                    cmap_name = scanner.parse_name().ok().map(|n| n.as_bytes().to_vec());
                }
                Some("WMode") => {
                    writing_mode = parse_writing_mode(&mut scanner);
                }
                _ => {}
            }
        } else {
            match name.as_str() {
                Some("begincodespacerange") => {
                    parse_codespace_range(&mut scanner, &mut _codespace_ranges, &mut ctx);
                }
                Some("begincidrange") => {
                    parse_range(&mut scanner, &mut ranges, &mut ctx, "endcidrange");
                }
                Some("begincidchar") => {
                    parse_char(&mut scanner, &mut ranges, &mut ctx, "endcidchar");
                }
                Some("beginnotdefrange") => {
                    parse_range(&mut scanner, &mut notdef_ranges, &mut ctx, "endnotdefrange");
                }
                Some("beginnotdefchar") => {
                    parse_char(&mut scanner, &mut notdef_ranges, &mut ctx, "endnotdefchar");
                }
                Some("beginbfchar") => {
                    parse_bf_char(&mut scanner, &mut bf_entries, &mut ctx);
                }
                Some("beginbfrange") => {
                    parse_bf_range(&mut scanner, &mut bf_entries, &mut ctx);
                }
                Some("usecmap") => {
                    let Some(Token::Name(operand)) = prev else {
                        truncated = true;
                        break;
                    };

                    if !operand.is_literal() {
                        truncated = true;
                        break;
                    }

                    match (ctx.get_cmap)(CMapName::from_bytes(operand.as_bytes()))
                        .and_then(|nested_data| {
                            parse_inner(nested_data, ctx.get_cmap.clone(), depth + 1)
                        }) {
                        Some(nested) => base = Some(Box::new(nested)),
                        None => usecmap_failed = true,
                    }
                }
                _ => {}
            }
        }
    }

    // Since we will use binary search for finding the correct entry, sort now.
    ranges.sort();
    notdef_ranges.sort();
    bf_entries.sort_by_key(|a| a.range.start);

    // See PDFJS-3323, which has an invalid CIDSystemInfo entry.
    // We ignore it if it's invalid.
    let character_collection = if let (Some(registry), Some(ordering), Some(supplement)) =
        (registry, ordering, supplement)
    {
        Some(CharacterCollection {
            family: CidFamily::from_registry_ordering(&registry, &ordering),
            supplement,
        })
    } else {
        None
    };

    let metadata = Metadata {
        character_collection,
        name: cmap_name,
        writing_mode,
    };

    // PdfCraft patch: a CMap that was cut short or whose `usecmap` base could not be resolved,
    // and that yielded no mappings of its own, still reports failure, so callers fall back to
    // the font's own encoding exactly as before.
    if (truncated || usecmap_failed)
        && ranges.is_empty()
        && notdef_ranges.is_empty()
        && bf_entries.is_empty()
        && base.is_none()
    {
        return None;
    }

    Some(CMap {
        metadata,
        _codespace_ranges,
        cid_ranges: ranges,
        notdef_ranges,
        bf_entries,
        base,
    })
}

fn parse_writing_mode(scanner: &mut Scanner<'_>) -> Option<WritingMode> {
    match parse_integer(scanner)? {
        0 => Some(WritingMode::Horizontal),
        1 => Some(WritingMode::Vertical),
        _ => None,
    }
}

fn parse_integer(scanner: &mut Scanner<'_>) -> Option<i64> {
    match scanner.parse_number().ok()? {
        Number::Integer(n) => Some(n),
        Number::Real(_) => None,
    }
}

fn parse_string_or_name(scanner: &mut Scanner<'_>) -> Option<Vec<u8>> {
    match scanner.parse_token().ok()? {
        Token::String(s) => s.decode().ok(),
        Token::Name(n) => Some(n.as_bytes().to_vec()),
        _ => None,
    }
}

fn parse_codespace_range<F>(
    scanner: &mut Scanner<'_>,
    ranges: &mut Vec<CodespaceRange>,
    ctx: &mut Context<F>,
) {
    loop {
        let Ok(obj) = scanner.parse_token() else { return };

        if name_matches(&obj, "endcodespacerange") {
            return;
        }

        // PdfCraft patch: a malformed codespace entry is skipped, not fatal.
        let Some(low) = extract_u32_code(&obj, &mut ctx.buf) else { continue };
        let Ok(n_bytes) = u8::try_from(ctx.buf.len()) else { continue };
        let Some(high) = read_u32_code(scanner, &mut ctx.buf) else { continue };

        if ctx.buf.len() != usize::from(n_bytes) {
            continue;
        }

        ranges.push(CodespaceRange {
            number_bytes: n_bytes,
            low,
            high,
        });
    }
}

fn parse_range<F>(
    scanner: &mut Scanner<'_>,
    ranges: &mut PartitionedRanges,
    ctx: &mut Context<F>,
    end_marker: &str,
) {
    loop {
        let Ok(obj) = scanner.parse_token() else { return };

        if name_matches(&obj, end_marker) {
            return;
        }

        // PdfCraft patch: a malformed range entry is skipped, not fatal.
        let Some(start) = extract_u32_code(&obj, &mut ctx.buf) else { continue };
        let byte_len = ctx.buf.len();
        let Some(end) = read_u32_code(scanner, &mut ctx.buf) else { continue };
        let Some(cid_start) =
            u32::try_from(parse_integer(scanner).unwrap_or(i64::MIN)).ok()
        else {
            continue;
        };

        ranges.push(
            byte_len,
            CidRange {
                range: Range { start, end },
                cid_start,
            },
        );
    }
}

fn parse_char<F>(
    scanner: &mut Scanner<'_>,
    ranges: &mut PartitionedRanges,
    ctx: &mut Context<F>,
    end_marker: &str,
) {
    loop {
        let Ok(obj) = scanner.parse_token() else { return };

        if name_matches(&obj, end_marker) {
            return;
        }

        // PdfCraft patch: a malformed char entry is skipped, not fatal.
        let Some(code) = extract_u32_code(&obj, &mut ctx.buf) else { continue };
        let byte_len = ctx.buf.len();
        let Some(cid_start) =
            u32::try_from(parse_integer(scanner).unwrap_or(i64::MIN)).ok()
        else {
            continue;
        };

        ranges.push(
            byte_len,
            CidRange {
                range: Range {
                    start: code,
                    end: code,
                },
                cid_start,
            },
        );
    }
}

fn parse_bf_char<F>(
    scanner: &mut Scanner<'_>,
    entries: &mut Vec<BfRange>,
    ctx: &mut Context<F>,
) {
    loop {
        let Ok(obj) = scanner.parse_token() else { return };

        if name_matches(&obj, "endbfchar") {
            return;
        }

        // PdfCraft patch: a malformed bfchar entry (e.g. a non-hex destination) is skipped, not
        // fatal - the entries after it still apply.
        let Some(code) = extract_u32_code(&obj, &mut ctx.buf) else {
            // The malformed code is still followed by its destination object; skip it so later
            // entries stay aligned.
            let _ = scanner.parse_token();
            continue;
        };
        let Ok(dst) = scanner.parse_string() else { continue };
        if dst.decode_into(&mut ctx.buf).is_err() {
            continue;
        }
        let Some(dst_base) = decode_be(&ctx.buf) else { continue };

        entries.push(BfRange {
            range: Range {
                start: code,
                end: code,
            },
            dst_base,
        });
    }
}

fn parse_bf_range<F>(
    scanner: &mut Scanner<'_>,
    entries: &mut Vec<BfRange>,
    ctx: &mut Context<F>,
) {
    loop {
        let Ok(obj) = scanner.parse_token() else { return };

        if name_matches(&obj, "endbfrange") {
            return;
        }

        // PdfCraft patch: a malformed bfrange entry is skipped, not fatal; the tokens of the bad
        // entry are consumed so later entries stay aligned.
        let Some(start) = extract_u32_code(&obj, &mut ctx.buf) else {
            let _ = scanner.parse_token();
            let _ = scanner.parse_token();
            continue;
        };
        let Some(end) = read_u32_code(scanner, &mut ctx.buf) else {
            let _ = scanner.parse_token();
            continue;
        };

        let Ok(next) = scanner.parse_token() else { return };

        match &next {
            Token::String(s) => {
                if s.decode_into(&mut ctx.buf).is_err() {
                    continue;
                }
                if let Some(dst_base) = decode_be(&ctx.buf) {
                    entries.push(BfRange {
                        range: Range { start, end },
                        dst_base,
                    });
                }
            }
            _ if name_matches(&next, "[") => {
                for code in start..=end {
                    let Ok(s) = scanner.parse_string() else { return };
                    if s.decode_into(&mut ctx.buf).is_err() {
                        continue;
                    }
                    if let Some(dst_base) = decode_be(&ctx.buf) {
                        entries.push(BfRange {
                            range: Range {
                                start: code,
                                end: code,
                            },
                            dst_base,
                        });
                    }
                }

                match scanner.parse_token() {
                    Ok(tok) if name_matches(&tok, "]") => {}
                    _ => return,
                }
            }
            _ => continue,
        }
    }
}

/// Convert the buffer into native-endian u16, so that we can use `String::from_utf16`.
fn decode_be(bytes: &[u8]) -> Option<Vec<u16>> {
    if bytes.is_empty() {
        return None;
    }

    let mut out = Vec::with_capacity(bytes.len().div_ceil(2));
    let mut i = 0;

    // My understanding is that the bf strings should always be UTF16-BE encoded,
    // but it seems like some PDFs only have a single byte there? I guess we just
    // pad them? Wasn't able to find anything specific in the specifications.
    if !bytes.len().is_multiple_of(2) {
        out.push(u16::from(bytes[0]));
        i = 1;
    }

    while i < bytes.len() {
        out.push(u16::from_be_bytes([bytes[i], bytes[i + 1]]));
        i += 2;
    }

    Some(out)
}

#[inline]
fn read_u32_code(scanner: &mut Scanner<'_>, buf: &mut Vec<u8>) -> Option<u32> {
    let s = scanner.parse_string().ok()?;
    s.decode_into(buf).ok()?;
    bytes_to_u32(buf)
}

#[inline]
fn extract_u32_code(obj: &Token<'_>, buf: &mut Vec<u8>) -> Option<u32> {
    let Token::String(s) = obj else { return None };
    s.decode_into(buf).ok()?;
    bytes_to_u32(buf)
}

#[inline]
fn bytes_to_u32(bytes: &[u8]) -> Option<u32> {
    if bytes.is_empty() || bytes.len() > 4 {
        return None;
    }

    let mut val = 0_u32;
    for &b in bytes {
        val = (val << 8) | b as u32;
    }

    Some(val)
}

fn name_matches(obj: &Token<'_>, expected: &str) -> bool {
    matches!(obj, Token::Name(name) if !name.is_literal() && name.as_str() == Some(expected))
}
