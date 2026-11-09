## 3. FAMILY PATCHES (full modified bodies; apply by function-name anchor inside the copied tree; every block carries a `PdfCraft patch:` mark)

### F5 — sync poison recovery — `vendor/hayro-syntax/src/sync.rs`
Replace the two `#[cfg(feature = "std")]` impl bodies (the `#[cfg(not(feature = "std"))]` impls stay upstream):
```rust
#[cfg(feature = "std")]
impl<T> MutexExt<T> for Mutex<T> {
    fn get(&self) -> MutexGuard<'_, T> {
        // PdfCraft patch: (#307) cached values remain owned after an unwinding caller; do not cascade a
        // worker panic into unrelated readers. Initialization has its own failure state.
        self.lock().unwrap_or_else(|poisoned| poisoned.into_inner())
    }
}
```
```rust
#[cfg(feature = "std")]
impl<T> RwLockExt<T> for RwLock<T> {
    // PdfCraft patch: (#307) poisoned locks are recovered rather than unwrapped.
    fn get(&self) -> RwLockReadGuard<'_, T> {
        self.read().unwrap_or_else(|poisoned| poisoned.into_inner())
    }

    fn put(&self) -> RwLockWriteGuard<'_, T> {
        self.write()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
    }
}
```

### F4 — Prev-as-indirect-ref — `vendor/hayro-syntax/src/xref.rs`
(a) Insert directly **after** `fn populate_xref_impl` (anchor: the line before `/// Maximum number of allowed xref \`Prev\` pointers before we abort.`):
```rust
/// The offset of the previous xref section (`Prev`).
// PdfCraft patch: (#2366) some writers give `Prev` as a reference `N G R`. When this section's
// own entries (`own`) place that object and it holds an offset (iText's `indirectPrev.pdf`), that
// is the offset; otherwise its object number is (tiff2pdf, in JHOVE's PDF-HUL-157 files), as
// other readers take it.
fn prev_offset(data: &[u8], trailer: &Dict<'_>, own: &XrefMap) -> Option<usize> {
    if let Some(prev) = trailer.get::<i32>(PREV) {
        return usize::try_from(prev).ok();
    }

    let r = trailer.get_ref(PREV)?;
    let id = ObjectIdentifier::new(r.obj_number, r.gen_number);
    let held = match own.get(&id) {
        Some(EntryType::Normal(at)) => {
            let mut reader = Reader::new(data);
            reader.jump(*at);
            reader
                .read_with_context::<IndirectObject<i32>>(&ReaderContext::dummy())
                .filter(|o| o.id() == &id)
                .map(|o| o.get())
                .filter(|p| *p >= 0)
        }
        _ => None,
    };

    usize::try_from(held.unwrap_or(r.obj_number)).ok()
}
```
(b) Replace `fn populate_from_xref_table` with:
```rust
/// Populate the xref table, and return the trailer dict.
fn populate_from_xref_table<'a>(
    data: &'a [u8],
    reader: &mut Reader<'a>,
    insert_map: &mut XrefMap,
    visited: &mut BTreeSet<usize>,
) -> Option<&'a [u8]> {
    let trailer = {
        let mut reader = reader.clone();
        read_xref_table_trailer(&mut reader, &ReaderContext::dummy())?
    };

    reader.skip_white_spaces();
    reader.forward_tag(b"xref")?;
    reader.skip_white_spaces();

    let mut max_obj = 0;
    // PdfCraft patch: (#2366) this section's entries are read first, for `prev_offset`.
    let mut own = XrefMap::default();

    while let Some(header) = reader.read_without_context::<SubsectionHeader>() {
        reader.skip_white_spaces();

        let start = header.start;
        let end = start.checked_add(header.num_entries)?;

        for obj_number in start..end {
            max_obj = max(max_obj, obj_number);
            let bytes = reader.read_bytes(XREF_ENTRY_LEN)?;
            let entry = XRefEntry::read(bytes)?;

            // Specification says we should ignore any object number > SIZE, but probably
            // not important?
            if entry.used {
                own.insert(
                    ObjectIdentifier::new(obj_number, entry.gen_number),
                    EntryType::Normal(entry.offset),
                );
            }
        }
    }

    if let Some(prev) = prev_offset(data, &trailer, &own) {
        // First insert the entries from any previous xref tables.
        populate_xref_impl_inner(data, prev, insert_map, visited)?;
    }

    // In hybrid files, entries in `XRefStm` should have higher priority, therefore we insert them
    // after looking at `PREV`.
    if let Some(xref_stm) = trailer.get::<i32>(XREF_STM) {
        populate_xref_impl_inner(data, xref_stm as usize, insert_map, visited)?;
    }

    insert_map.extend(own);

    Some(trailer.data())
}
```
(c) In `fn populate_from_xref_stream`: delete the early `if let Some(prev) = stream.dict().get::<i32>(PREV) { … }` block (before `let size = …`), and replace from `let xref_data = …` through `Some(stream.dict().data())` with:
```rust
    let xref_data = stream.decoded_data().ok()?;
    // PdfCraft patch: (#2366) this section's entries are read first, for `prev_offset`.
    let mut own = XrefMap::default();
    let mut xref_reader = Reader::new(xref_data.as_ref());

    if let Some(arr) = stream.dict().get::<Array<'_>>(INDEX) {
        let iter = arr.iter::<(u32, u32)>();

        for (start, num_elements) in iter {
            xref_stream_subsection(
                &mut xref_reader,
                start,
                num_elements,
                f1_len,
                f2_len,
                f3_len,
                &mut own,
            )?;
        }
    } else {
        xref_stream_subsection(&mut xref_reader, 0, size, f1_len, f2_len, f3_len, &mut own)?;
    }

    if let Some(prev) = prev_offset(data, stream.dict(), &own) {
        // First insert the entries from any previous xref tables.
        let _ = populate_xref_impl_inner(data, prev, insert_map, visited)?;
    }

    insert_map.extend(own);

    Some(stream.dict().data())
```
(keep upstream's `let size`/`let [f1_len,…]` lines; `xref_stream_subsection` callers keep their current argument shapes; `let mut max_obj = 0;` in the table fn stays for the `max` use above).

### F6 — decode limits — `vendor/hayro-syntax/src/filter/`
**run_length.rs** — replace `decode` and its test:
```rust
/// PdfCraft patch: `limit` is the most bytes produced (see `filter::decode_limit`); a run of 128
/// copies costs two input bytes, so a few megabytes could ask for gigabytes.
pub(crate) fn decode(data: &[u8], limit: usize) -> Option<Vec<u8>> {
    let mut reader = Reader::new(data);
    let mut decoded = vec![];

    loop {
        let left = limit.saturating_sub(decoded.len());
        if left == 0 {
            break;
        }
        let length = reader.read_byte()?;

        match length {
            128 => break,
            0..=127 => {
                // PDFBOX-3990, just abort early if stream is invalid.
                let Some(bytes) = reader.read_bytes(length as usize + 1) else {
                    break;
                };

                decoded.extend(bytes.iter().take(left));
            }
            _ => {
                let length = (257 - length as usize).min(left);
                let byte = reader.read_byte()?;
                decoded.extend(core::iter::repeat_n(byte, length));
            }
        }
    }

    Some(decoded)
}

#[cfg(test)]
mod tests {
    use crate::filter::run_length::decode;

    #[test]
    fn run_length() {
        let input = vec![4, 10, 11, 12, 13, 14, 253, 3, 128];
        assert_eq!(
            decode(&input, usize::MAX).unwrap(),
            vec![10, 11, 12, 13, 14, 3, 3, 3, 3]
        );
    }
}
```
**mod.rs** — replace `Filter::apply` with the limit-threaded version and add the limit helpers after the impl:
```rust
    pub(crate) fn apply(
        &self,
        data: &[u8],
        params: &Dict<'_>,
        #[cfg_attr(not(feature = "images"), allow(unused))] image_params: &ImageDecodeParams,
        limit: usize,
    ) -> Result<FilterResult<'static>, DecodeFailure> {
        match self {
            Self::AsciiHexDecode => ascii_hex::decode(data)
                .map(FilterResult::from_data)
                .ok_or(DecodeFailure::StreamDecode),
            Self::Ascii85Decode => ascii_85::decode(data)
                .map(FilterResult::from_data)
                .ok_or(DecodeFailure::StreamDecode),
            Self::RunLengthDecode => run_length::decode(data, limit)
                .map(FilterResult::from_data)
                .ok_or(DecodeFailure::StreamDecode),
            Self::LzwDecode => lzw_flate::lzw::decode(data, params, limit)
                .map(FilterResult::from_data)
                .ok_or(DecodeFailure::StreamDecode),
            Self::FlateDecode => lzw_flate::flate::decode(data, params, limit)
                .map(FilterResult::from_data)
                .ok_or(DecodeFailure::StreamDecode),
            #[cfg(feature = "images")]
            Self::DctDecode => {
                dct::decode(data, params, image_params).ok_or(DecodeFailure::ImageDecode)
            }
            #[cfg(feature = "images")]
            Self::CcittFaxDecode => {
                ccitt::decode(data, params, image_params).ok_or(DecodeFailure::ImageDecode)
            }
            #[cfg(feature = "images")]
            Self::Jbig2Decode => {
                jbig2::decode(data, params, image_params).ok_or(DecodeFailure::ImageDecode)
            }
            #[cfg(feature = "images")]
            Self::JpxDecode => jpx::decode(data, image_params).ok_or(DecodeFailure::ImageDecode),
            #[cfg(not(feature = "images"))]
            Self::DctDecode | Self::CcittFaxDecode | Self::Jbig2Decode | Self::JpxDecode => {
                Err(DecodeFailure::ImageDecode)
            }
            _ => Err(DecodeFailure::StreamDecode),
        }
    }
```
```rust
/// PdfCraft patch: the most bytes Flate, LZW or RunLength may produce for a stream that isn't an
/// image of known size: content streams, fonts, forms, ICC profiles. These filters expand without
/// limit (Flate about 1000:1), so a 2 MB page that inflated to 2 GB took the renderer past 3 GB in
/// under a second. They never write past the limit. Real streams are far smaller.
pub const MAX_DECODED_STREAM: usize = 256 << 20;

/// PdfCraft patch: the most bytes they may produce for an image of known size, however large it
/// says it is.
const MAX_DECODED_IMAGE: u64 = 1 << 30;

/// PdfCraft patch: the output limit for one filter of a stream decoded with `image_params`;
/// `last` says whether it is the last filter. The last filter of an image whose size, bits per
/// component and components are known produces its pixels, plus at most one PNG predictor tag
/// byte per pixel byte. Earlier filters produce compressed data, which can be larger than the
/// pixels, and everything else gets [`MAX_DECODED_STREAM`]. Decoding stops at the limit and
/// keeps what it has.
pub(crate) fn decode_limit(image_params: &ImageDecodeParams, last: bool) -> usize {
    let (true, Some(bpc), Some(components)) = (last, image_params.bpc, image_params.num_components)
    else {
        return MAX_DECODED_STREAM;
    };
    if image_params.width == 0 || image_params.height == 0 {
        return MAX_DECODED_STREAM;
    }
    let bits = u64::from(image_params.width)
        * u64::from(components)
        * u64::from(bpc);
    let pixels = bits.div_ceil(8).saturating_mul(u64::from(image_params.height));
    let image = pixels.saturating_mul(2).min(MAX_DECODED_IMAGE);
    usize::try_from(image).unwrap_or(MAX_DECODED_STREAM)
}
```
**lzw_flate.rs** — replace `flate::decode` and both `inflate` fns (limit stops decoding, keeps what it has), and `lzw::decode`/`decode_impl`:
```rust
pub(crate) mod flate {
    use crate::filter::predictor;
    use crate::object::Dict;
    use alloc::vec::Vec;

    /// PdfCraft patch: `limit` is the most bytes inflated (see `filter::decode_limit`).
    pub(crate) fn decode(data: &[u8], params: &Dict<'_>, limit: usize) -> Option<Vec<u8>> {
        let data = zlib_body(data).unwrap_or(data);
        // Broken encoders can refer to data before the start, which should read as zeros.
        let decoded = inflate(data, false, limit)
            .or_else(|_| inflate(data, true, limit))
            .unwrap_or_else(|partial| partial);
        predictor::apply(decoded, params)
    }

    const WINDOW_SIZE: usize = 32 * 1024;

    fn zlib_body(data: &[u8]) -> Option<&[u8]> {
        let [cmf, flg, body @ ..] = data else {
            return None;
        };

        ((cmf & 0x0f) == 0x08
            && (u16::from(*cmf) << 8 | u16::from(*flg)).is_multiple_of(31)
            && (flg & 0x20) == 0)
            .then_some(body)
    }

    #[cfg(feature = "unsafe")]
    fn inflate(mut data: &[u8], zero_window: bool, limit: usize) -> Result<Vec<u8>, Vec<u8>> {
        use zlib_rs::{Inflate, InflateFlush, Status};

        static ZEROS: [u8; WINDOW_SIZE] = [0; WINDOW_SIZE];

        let mut inflate = Inflate::new(false, 15);
        if zero_window {
            inflate.set_dictionary(&ZEROS).ok();
        }
        let mut out = Vec::new();
        let mut written = 0;

        loop {
            // PdfCraft patch: never grow the output buffer past `limit`.
            let room = limit.saturating_sub(written);
            if room == 0 {
                out.truncate(limit);
                return Ok(out);
            }
            out.resize((out.len() * 2).max(WINDOW_SIZE).min(written + room), 0);

            let (total_in, total_out) = (inflate.total_in(), inflate.total_out());
            let result = inflate.decompress(data, &mut out[written..], InflateFlush::Finish);
            let consumed = (inflate.total_in() - total_in) as usize;
            let produced = (inflate.total_out() - total_out) as usize;
            written += produced;
            data = data.get(consumed..).unwrap_or_default();

            match result {
                Ok(Status::Ok | Status::BufError) if consumed != 0 || produced != 0 => {}
                Ok(_) => {
                    out.truncate(written.min(limit));
                    return Ok(out);
                }
                Err(_) => {
                    out.truncate(written.min(limit));
                    return Err(out);
                }
            }
        }
    }

    #[cfg(not(feature = "unsafe"))]
    fn inflate(mut data: &[u8], zero_window: bool, limit: usize) -> Result<Vec<u8>, Vec<u8>> {
        use alloc::boxed::Box;
        use alloc::vec;
        use miniz_oxide::inflate::TINFLStatus;
        use miniz_oxide::inflate::core::inflate_flags::TINFL_FLAG_USING_NON_WRAPPING_OUTPUT_BUF;
        use miniz_oxide::inflate::core::{DecompressorOxide, decompress};

        let start = if zero_window { WINDOW_SIZE } else { 0 };
        let mut decompressor = Box::<DecompressorOxide>::default();
        let mut out = vec![0; start];
        let mut written = start;

        loop {
            // PdfCraft patch: never grow the output buffer past `limit` (plus the window prefix
            // that is drained at the end, so `written - start` is what counts).
            if written.saturating_sub(start) >= limit {
                out.truncate(start + limit);
                out.drain(..start);
                return Ok(out);
            }
            let room = limit - written.saturating_sub(start);
            let grown = (out.len() * 2).max(WINDOW_SIZE);
            out.resize(grown.min(out.len() + room.max(WINDOW_SIZE)), 0);

            let (status, consumed, produced) = decompress(
                &mut decompressor,
                data,
                &mut out,
                written,
                TINFL_FLAG_USING_NON_WRAPPING_OUTPUT_BUF,
            );
            written += produced;
            data = data.get(consumed..).unwrap_or_default();

            if status != TINFLStatus::HasMoreOutput {
                out.truncate(written.min(start + limit));
                out.drain(..start);

                return match status {
                    TINFLStatus::Done | TINFLStatus::FailedCannotMakeProgress => Ok(out),
                    _ => Err(out),
                };
            }
        }
    }
}
```
```rust
    /// Decode a LZW-encoded stream.
    pub(crate) fn decode(data: &[u8], params: &Dict<'_>, limit: usize) -> Option<Vec<u8>> {
        let early_change = params
            .get::<u8>(EARLY_CHANGE)
            .map(|e| e != 0)
            .unwrap_or(true);
        let decoded = decode_impl(data, early_change, limit)?;

        predictor::apply(decoded, params)
    }
```
```rust
    fn decode_impl(data: &[u8], early_change: bool, limit: usize) -> Option<Vec<u8>> {
        let mut table = Table::new(early_change);
        let mut bit_size = table.code_length();
        let mut reader = BitReader::new(data);
        let mut decoded = vec![];
        let mut prev = None;

        loop {
            let next = match reader.read(bit_size) {
                Some(code) => code as usize,
                None => {
                    return Some(decoded);
                }
            };

            match next {
                CLEAR_TABLE => {
                    table.clear();
                    prev = None;
                    bit_size = table.code_length();
                }
                EOD => return Some(decoded),
                new => {
                    if new > table.size() {
                        return None;
                    }

                    // PdfCraft patch: never past `limit`.
                    let left = limit.saturating_sub(decoded.len());
                    if left == 0 {
                        return Some(decoded);
                    }

                    if new < table.size() {
                        let entry = table.get(new)?;
                        let first_byte = entry[0];
                        decoded.extend_from_slice(entry.get(..entry.len().min(left))?);

                        if let Some(prev_code) = prev {
                            table.register(prev_code, first_byte);
                        }
                    } else if new == table.size() && prev.is_some() {
                        let prev_code = prev.unwrap();
                        let prev_entry = table.get(prev_code)?;
                        let first_byte = prev_entry[0];

                        let new_entry = table.register(prev_code, first_byte)?;
                        decoded.extend_from_slice(new_entry.get(..new_entry.len().min(left))?);
                    } else {
                        return None;
                    }

                    bit_size = table.code_length();
                    prev = Some(new);
                }
            }
        }
    }
```
**object/stream.rs** — replace `decoded_image` and add `decoded_within`; `decoded`/`decoded_data` keep their shapes (`decoded` = `self.decoded_image(&ImageDecodeParams::default()).map(|r| r.data)`; `decoded_data` unchanged except it now benefits from `decoded`'s internal limit):
```rust
    /// Return the decoded data of the stream, and return image metadata
    /// if available.
    pub fn decoded_image(
        &self,
        image_params: &ImageDecodeParams,
    ) -> Result<FilterResult<'a>, DecodeFailure> {
        self.decoded_within(image_params, usize::MAX)
    }

    /// PdfCraft patch: the decoded data of the stream with an explicit output budget. Flate, LZW
    /// and RunLength stop at `max` bytes (and at [`crate::filter::MAX_DECODED_STREAM`] in any
    /// case). A stream without filters is returned as it is; ASCIIHex and ASCII85 barely expand,
    /// and the image codecs stop at their own pixel limits.
    pub fn decoded_within(
        &self,
        image_params: &ImageDecodeParams,
        max: usize,
    ) -> Result<FilterResult<'a>, DecodeFailure> {
        if matches!(
            self.dict.get::<Object<'_>>(F),
            Some(Object::String(_) | Object::Dict(_))
        ) {
            return Err(DecodeFailure::StreamDecode);
        }

        let data = self.raw_data();
        let filters_and_params = self.filters_and_params();

        let mut current: Option<FilterResult<'a>> = None;

        let last = filters_and_params.filters.len().saturating_sub(1);
        for (i, (filter, params)) in filters_and_params
            .filters
            .iter()
            .zip(filters_and_params.params.iter())
            .enumerate()
        {
            if i == 0 && *filter == Filter::Crypt {
                continue;
            }

            // PdfCraft patch: each filter step is bounded; the last one may produce the image's
            // own pixels (see `decode_limit`), everything else at most `MAX_DECODED_STREAM`,
            // further capped by the caller's budget.
            let limit = crate::filter::decode_limit(image_params, i == last).min(max);

            let new = filter.apply(
                current.as_ref().map(|c| c.data.as_ref()).unwrap_or(&data),
                params,
                image_params,
                limit,
            )?;
            current = Some(new);
        }

        Ok(current.unwrap_or(FilterResult {
            data,
            image_data: None,
        }))
    }
```
**page.rs** — replace `page_stream` (F1 total-bytes cap; per-stream cap comes from F6):
```rust
    /// Return the decoded content stream of the page.
    pub fn page_stream(&self) -> Option<&[u8]> {
        // PdfCraft patch: the most bytes one page's decoded /Contents may hold. Every stream is
        // already capped by `MAX_DECODED_STREAM`; the total is capped here so a page with many
        // large streams cannot hold an unbounded concatenation.
        const MAX_PAGE_STREAM_BYTES: usize = 512 << 20;

        let convert_single = |s: Stream<'_>| {
            let data = s.decoded_data().ok()?;
            Some(data.into_owned())
        };

        self.page_streams
            .get_or_init(|| {
                if let Some(stream) = self.inner.get::<Stream<'_>>(CONTENTS) {
                    convert_single(stream)
                } else if let Some(array) = self.inner.get::<Array<'_>>(CONTENTS) {
                    let streams = array
                        .raw_iter()
                        .filter_map(|item| item.resolve(&self.ctx)?.into_stream())
                        .filter_map(convert_single);

                    let mut collected = vec![];

                    for stream in streams {
                        // PdfCraft patch: stop once the page's content budget is spent.
                        if collected.len() + stream.len() > MAX_PAGE_STREAM_BYTES {
                            break;
                        }
                        collected.extend(stream);
                        // Streams must have at least one whitespace in-between.
                        collected.push(b' ');
                    }

                    Some(collected)
                } else {
                    None
                }
            })
            .as_ref()
            .map(|d| d.as_slice())
    }
```

### F1 — interpret-side content budget — `vendor/hayro-interpret/src/context.rs`
Insert after `pub(crate) const MAX_NESTED_INTERPRETATION_DEPTH: u32 = 50;`:
```rust
std::thread_local! {
    static CONTENT_LEFT: core::cell::Cell<usize> = const { core::cell::Cell::new(MAX_PAGE_CONTENT) };
    // PdfCraft patch: whether this page's content budget ran out.
    static CONTENT_TRUNCATED: core::cell::Cell<bool> = const { core::cell::Cell::new(false) };
}

/// PdfCraft patch: decoded content-stream bytes one page may hold at once: its contents, and
/// the forms, tiling patterns and Type 3 glyph procedures being painted. Each stream already
/// stops at `MAX_DECODED_STREAM`, but forms nest fifty deep and every level keeps its content
/// while the next one runs. A level gives its bytes back when it finishes, so content painted
/// again and again doesn't use the budget up. Ordinary pages hold a few megabytes.
pub(crate) const MAX_PAGE_CONTENT: usize = 512 << 20;

/// PdfCraft patch: a content stream decoded within the page's [`MAX_PAGE_CONTENT`]. The bytes
/// it decoded count against the budget until it is dropped; a stream without filters is
/// borrowed from the file and costs nothing.
pub(crate) struct DecodedContent<'a> {
    data: Cow<'a, [u8]>,
    held: usize,
}

impl AsRef<[u8]> for DecodedContent<'_> {
    fn as_ref(&self) -> &[u8] {
        &self.data
    }
}

impl Drop for DecodedContent<'_> {
    fn drop(&mut self) {
        let held = self.held;
        CONTENT_LEFT.with(|c| c.set(c.get().saturating_add(held).min(MAX_PAGE_CONTENT)));
    }
}

/// PdfCraft patch: decode a content stream within what is left of this page's
/// [`MAX_PAGE_CONTENT`]. `None` once the budget is spent (the content is skipped).
pub(crate) fn decode_content<'a>(stream: &Stream<'a>) -> Option<DecodedContent<'a>> {
    let left = CONTENT_LEFT.with(core::cell::Cell::get);
    if left == 0 {
        CONTENT_TRUNCATED.with(|truncated| truncated.set(true));
        return None;
    }
    let data = stream.decoded_within(left).ok()?;
    let held = match &data {
        Cow::Owned(v) => v.len(),
        Cow::Borrowed(_) => 0,
    };
    charge_content(held);
    Some(DecodedContent { data, held })
}

/// PdfCraft patch: count `len` bytes held for the rest of the page (its own contents).
pub(crate) fn charge_content(len: usize) {
    CONTENT_LEFT.with(|c| c.set(c.get().saturating_sub(len)));
}

/// PdfCraft patch: whether content was skipped because the page's budget ran out.
pub(crate) fn content_was_truncated() -> bool {
    CONTENT_TRUNCATED.with(core::cell::Cell::get)
}
```
Add to the file's `use` block: `use std::borrow::Cow;` and `use hayro_syntax::object::Stream;`. Replace `Context::new` body to add the fresh-budget reset (mark added, rest of the fn unchanged):
```rust
    /// Create a new context.
    pub fn new(
        initial_transform: Affine,
        bbox: Rect,
        cache: &InterpreterCache<'a>,
        xref: &'a XRef,
        settings: InterpreterSettings,
    ) -> Self {
        let state = State::new(initial_transform);
        // PdfCraft patch: a page (or other top-level content) starts with a fresh content budget.
        CONTENT_LEFT.with(|c| c.set(MAX_PAGE_CONTENT));
        CONTENT_TRUNCATED.with(|truncated| truncated.set(false));

        Self::new_with(initial_transform, bbox, cache, xref, settings, state, 0)
    }
```
Add to `impl<'a> Context<'a>` (after `nesting_depth()`):
```rust
    /// PdfCraft patch: a copy of the current graphics state (for re-interpreting a form's
    /// contents on their own, see `cmyk_group`).
    pub(crate) fn snapshot_state(&self) -> State<'a> {
        self.stack.snapshot()
    }
```
(`State` is already imported in context.rs.)

### F1 sites — replace the three amplification decodes with budgeted ones
`vendor/hayro-interpret/src/x_object/form.rs` — in `FormXObject`: field `pub(crate) decoded: Cow<'a, [u8]>,` becomes `pub(crate) decoded: crate::context::DecodedContent<'a>,` and in `FormXObject::new`:
```rust
        // PdfCraft patch: see `context::MAX_PAGE_CONTENT`.
        let decoded = crate::context::decode_content(stream)?;
```
(the `use std::borrow::Cow;` import may be removed if now unused; `decoded.as_ref()` call sites keep working via `AsRef<[u8]>`).

`vendor/hayro-interpret/src/font/type3.rs` — in `render_glyph`:
```rust
    let program = glyph.glyph.program()?;
    // PdfCraft patch: see `context::MAX_PAGE_CONTENT`.
    let decoded = crate::context::decode_content(&program)?;
    let iter = TypedIter::new(decoded.as_ref());
```
`vendor/hayro-interpret/src/pattern.rs` — in `TilingPattern::interpret`:
```rust
        // PdfCraft patch: see `context::MAX_PAGE_CONTENT`.
        let decoded = crate::context::decode_content(&self.stream)?;
        let iter = TypedIter::new(decoded.as_ref());
```

### F2 (interpret side) — decode gate — `vendor/hayro-interpret/src/x_object/decode/mod.rs`
In `fn decode_context`, after the `let (width, height) = decoded…unwrap_or((obj.width(), obj.height()));` block, insert:
```rust
    // PdfCraft patch: images and masks over `MAX_IMAGE_PIXELS` are not decoded at all (the
    // renderer refuses to draw them either); hostile dimensions cannot allocate here.
    if !image_size_ok(width, height) {
        (obj.warning_sink)(InterpreterWarning::ImageDecodeFailure);
        return None;
    }
```
and add at the end of the file:
```rust
/// PdfCraft patch: whether an image or mask of `width` × `height` may be decoded (at most
/// 2^28 pixels, the renderer's `MAX_IMAGE_PIXELS`).
pub(crate) fn image_size_ok(width: u32, height: u32) -> bool {
    width > 0 && height > 0 && u64::from(width) * u64::from(height) <= 1 << 28
}
```

### F3 — lenient cmap — `vendor/hayro-cmap/src/parse.rs`
Replace `parse_inner` (partial acceptance + tolerant usecmap; `MAX_NESTING_DEPTH` kept):
```rust
pub(crate) fn parse_inner<'a>(
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
```
Replace the five per-entry parsers (skip-and-continue per entry):
```rust
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
```
(`parse_writing_mode`, `parse_integer`, `parse_string_or_name`, `decode_be`, `read_u32_code`, `extract_u32_code`, `name_matches` stay upstream; `extract_u32_code`/`read_u32_code` keep `Option` returns. Note the `?`-on-`parse_integer` became `.unwrap_or(i64::MIN)` + `try_from` skip so a bad CID skips only the entry, not the block.)

### F9 — OCG /VE — `vendor/hayro-interpret/src/ocg.rs`
Extend the use line: add `VE` to the keys import and `MaybeRef, Object` to the object import:
```rust
use hayro_syntax::object::dict::keys::{BASE_STATE, D, OCGS, OCMD, OCPROPERTIES, OFF, ON, P, TYPE, VE};
use hayro_syntax::object::{Array, Dict, MaybeRef, Name, Object, ObjectIdentifier};
use std::collections::HashSet;
```
Struct + `dummy` + `from_catalog` first lines:
```rust
pub(crate) struct OcgState {
    inactive_ocgs: HashSet<ObjectIdentifier>,
    /// PdfCraft patch: every group the document lists in `/OCProperties /OCGs`, so a
    /// visibility expression naming anything else is known to be malformed.
    known_ocgs: HashSet<ObjectIdentifier>,
    visibility_stack: Vec<bool>,
}
```
```rust
    fn dummy() -> Self {
        Self {
            inactive_ocgs: HashSet::default(),
            known_ocgs: HashSet::default(),
            visibility_stack: vec![],
        }
    }
```
In `from_catalog`, replace `let mut inactive = HashSet::new();` with:
```rust
        let mut inactive = HashSet::new();
        // PdfCraft patch: the groups the document declares.
        let known: HashSet<ObjectIdentifier> = oc_properties
            .get::<Array<'_>>(OCGS)
            .map(|ocgs| {
                ocgs
                    .raw_iter()
                    .filter_map(|item| item.as_obj_ref())
                    .map(Into::into)
                    .collect()
            })
            .unwrap_or_default();
```
and its `Self { … }` tail with:
```rust
        Self {
            inactive_ocgs: inactive,
            known_ocgs: known,
            visibility_stack: Vec::new(),
        }
```
`begin_ocmd` — prepend the /VE precedence block:
```rust
    pub(crate) fn begin_ocmd(&mut self, ocmd: &Dict<'_>) {
        // PdfCraft patch: a visibility expression takes precedence over /OCGs and /P
        // (ISO 32000-2 §8.11.2.2); a malformed one is ignored and they decide instead.
        if let Some(visible) = ocmd
            .get::<Array<'_>>(VE)
            .and_then(|ve| self.visibility_expression(&ve, 0))
        {
            let visible = self.is_visible() && visible;
            self.visibility_stack.push(visible);
            return;
        }

        let policy = ocmd
// … rest of the upstream body unchanged …
```
Add the expression evaluator after `begin_ocmd` and the depth const before `enum BaseState`:
```rust
    /// PdfCraft patch: evaluate a visibility expression (`/VE`, PDF 1.6): `[/And e…]`,
    /// `[/Or e…]` or `[/Not e]`, each `e` an optional content group (by reference) or a nested
    /// expression. `None` for a malformed expression: an unknown operator, a `/Not` without
    /// exactly one operand, an `/And` or `/Or` without any, an operand that is neither, or
    /// nesting deeper than [`MAX_VISIBILITY_EXPRESSION_DEPTH`].
    fn visibility_expression(&self, expr: &Array<'_>, depth: usize) -> Option<bool> {
        if depth > MAX_VISIBILITY_EXPRESSION_DEPTH {
            return None;
        }
        let mut items = expr.raw_iter();
        let operator = match items.next()? {
            MaybeRef::NotRef(Object::Name(name)) => name,
            _ => return None,
        };
        let (mut count, mut all, mut any) = (0usize, true, false);
        for item in items {
            let value = match item {
                MaybeRef::Ref(r) => {
                    let id: ObjectIdentifier = r.into();
                    if !self.known_ocgs.contains(&id) {
                        return None;
                    }
                    !self.inactive_ocgs.contains(&id)
                }
                MaybeRef::NotRef(Object::Array(nested)) => {
                    self.visibility_expression(&nested, depth + 1)?
                }
                _ => return None,
            };
            count = count.saturating_add(1);
            all &= value;
            any |= value;
        }
        match (operator.as_ref(), count) {
            (b"And", 1..) => Some(all),
            (b"Or", 1..) => Some(any),
            (b"Not", 1) => Some(!all),
            _ => None,
        }
    }
```
```rust
/// PdfCraft patch: how deep visibility expressions may nest; deeper ones are malformed.
const MAX_VISIBILITY_EXPRESSION_DEPTH: usize = 32;
```

### F8 — annotation flags — `vendor/hayro-interpret/src/interpret/mod.rs`
Consts:
```rust
const ANNOTATION_INVISIBLE: u32 = 1 << 0;
const ANNOTATION_HIDDEN: u32 = 1 << 1;
const ANNOTATION_NO_ROTATE: u32 = 1 << 4;
const ANNOTATION_NO_VIEW: u32 = 1 << 5;
```
Warning variant (insert into `InterpreterWarning`):
```rust
    /// PdfCraft patch: a page content stream was skipped after exhausting its safety budget.
    ContentTruncated,
```
Replace `interpret_page`:
```rust
/// interpret the contents of the page and render them into the device.
pub fn interpret_page<'a>(
    page: &Page<'a>,
    context: &mut Context<'a>,
    device: &mut impl Device<'a>,
) {
    let resources = page.resources();
    let render_content = context.settings.render_content;

    if render_content.contains(RenderContent::PAGE) {
        interpret(page.typed_operations(), resources, context, device);
        // PdfCraft patch: report content the budget skipped (see `context::MAX_PAGE_CONTENT`).
        if crate::context::content_was_truncated() {
            (context.settings.warning_sink)(InterpreterWarning::ContentTruncated);
        }
    }

    if render_content.contains(RenderContent::ANNOTATIONS)
        && let Some(annot_arr) = page.raw().get::<Array<'_>>(ANNOTS)
    {
        for annot in annot_arr.iter::<Dict<'_>>() {
            interpret_annotation(&annot, resources, page, context, device);
        }
    }
}
```
Replace `interpret_annotation` and add the three helpers:
```rust
fn interpret_annotation<'a>(
    annot: &Dict<'a>,
    resources: &Resources<'a>,
    page: &Page<'a>,
    context: &mut Context<'a>,
    device: &mut impl Device<'a>,
) {
    let flags = annot.get::<u32>(F).unwrap_or(0);

    if flags & (ANNOTATION_HIDDEN | ANNOTATION_NO_VIEW) != 0 {
        return;
    }

    // PdfCraft patch: Invisible (bit 1) hides an annotation of a type that isn't one of the
    // standard ones (ISO 32000-2 Table 171), and a popup has no appearance of its own
    // (§12.5.6.14): the viewer shows it as a window for its parent, not on the page.
    let subtype = annot.get::<Name<'_>>(SUBTYPE);
    if subtype.as_deref() == Some(b"Popup")
        || (flags & ANNOTATION_INVISIBLE != 0
            && !subtype.as_deref().is_some_and(is_standard_annotation_type))
    {
        return;
    }

    if let Some(apx) = appearance_stream(annot).and_then(|o| FormXObject::new(&o)) {
        let Some(rect) = annot.get::<Rect>(RECT) else {
            return;
        };

        let annot_rect = rect.to_kurbo();
        // 12.5.5. Appearance streams
        // "The algorithm outlined in this subclause shall be used
        // to map from the coordinate system of the appearance XObject."

        // 1) The appearance’s bounding box (specified by its BBox entry) shall be
        // transformed, using Matrix, to produce a quadrilateral with arbitrary
        // orientation. The transformed appearance box is the smallest upright
        // rectangle that encompasses this quadrilateral.
        let transformed_rect = (apx.matrix
            * kurbo::Rect::new(
                apx.bbox[0] as f64,
                apx.bbox[1] as f64,
                apx.bbox[2] as f64,
                apx.bbox[3] as f64,
            )
            .to_path(0.1))
        .bounding_box();

        // 2) A matrix A shall be computed that scales and translates the transformed
        // appearance box to align with the edges of the annotation’s rectangle
        // (specified by the Rect entry).
        let sx = annot_rect.width() / transformed_rect.width();
        let sy = annot_rect.height() / transformed_rect.height();
        let affine = Affine::new([
            sx,
            0.0,
            0.0,
            sy,
            annot_rect.x0 - sx * transformed_rect.x0,
            annot_rect.y0 - sy * transformed_rect.y0,
        ]);
        // PdfCraft patch: NoRotate (bit 5) keeps the appearance upright as the page is
        // displayed (ISO 32000-2 §12.5.3).
        let affine = if flags & ANNOTATION_NO_ROTATE != 0 && no_rotate_applies(annot) {
            no_rotate_turn(page.rotation(), &annot_rect) * affine
        } else {
            affine
        };

        // 3) Matrix shall be concatenated with A to form a matrix AA that maps from
        // the appearance’s coordinate system to the annotation’s rectangle in
        // default user space.
        context.save_state();
        context.pre_concat_affine(affine);
        context.push_root_transform();

        apx.draw(resources, context, device);
        context.pop_root_transform();
        context.restore_state(device);
    } else if annot.get::<Dict<'_>>(AP).is_some() {
        (context.settings.warning_sink)(InterpreterWarning::UnresolvedAnnotationAppearance);
    }
}

/// PdfCraft patch: whether `subtype` is one of the standard annotation types of
/// ISO 32000-2 Table 171.
fn is_standard_annotation_type(subtype: &[u8]) -> bool {
    matches!(
        subtype,
        b"Text"
            | b"Link"
            | b"FreeText"
            | b"Line"
            | b"Square"
            | b"Circle"
            | b"Polygon"
            | b"PolyLine"
            | b"Highlight"
            | b"Underline"
            | b"Squiggly"
            | b"StrikeOut"
            | b"Caret"
            | b"Stamp"
            | b"Ink"
            | b"Popup"
            | b"FileAttachment"
            | b"Sound"
            | b"Movie"
            | b"Screen"
            | b"Widget"
            | b"PrinterMark"
            | b"TrapNet"
            | b"Watermark"
            | b"3D"
            | b"Redact"
            | b"Projection"
            | b"RichMedia"
    )
}

/// PdfCraft patch: whether the NoRotate flag turns this annotation's appearance.
fn no_rotate_applies(annot: &Dict<'_>) -> bool {
    !matches!(
        annot.get::<Name<'_>>(SUBTYPE).as_deref(),
        Some(b"Widget" | b"Link")
    )
}

/// PdfCraft patch: the user-space turn that keeps a NoRotate annotation upright on a page with
/// `/Rotate`. The page is shown turned clockwise, so the appearance (already fitted to `/Rect`)
/// is turned as far counterclockwise about the rectangle's upper-left corner, which stays where
/// it is on the page (ISO 32000-2 §12.5.3).
fn no_rotate_turn(rotation: hayro_syntax::page::Rotation, rect: &kurbo::Rect) -> Affine {
    let (cos, sin) = match rotation {
        hayro_syntax::page::Rotation::None => return Affine::IDENTITY,
        hayro_syntax::page::Rotation::Horizontal => (0.0, 1.0),
        hayro_syntax::page::Rotation::Flipped => (-1.0, 0.0),
        hayro_syntax::page::Rotation::FlippedHorizontal => (0.0, -1.0),
    };
    let (x, y) = (rect.x0.min(rect.x1), rect.y0.max(rect.y1));
    Affine::new([
        cos,
        sin,
        -sin,
        cos,
        x - (cos * x - sin * y),
        y - (sin * x + cos * y),
    ])
}
```
(`SUBTYPE` must be added to the keys import line: `{ ANNOTS, AP, AS, F, N, OC, RECT, SUBTYPE }`.)

### F7 (interpret side) — CMYK group capture + knockout hooks
**New file `vendor/hayro-interpret/src/cmyk_group.rs`** (add `mod cmyk_group;` + `pub use cmyk_group::CmykGroup;` to `src/lib.rs`):
```rust
//! PdfCraft patch: an isolated transparency group that blends in `DeviceCMYK`.
//!
//! ISO 32000-2 §11.3.4: a group's elements blend in the group's colour space, and the
//! components of a subtractive space are complemented before the blend function and
//! complemented back after it. The renderer composites in RGB, which gives other colours
//! wherever a blend mode meets CMYK artwork.
//!
//! An *isolated* group starts from a transparent backdrop of its own (§11.4.7), so its
//! interior can be computed apart from the page and handed back as finished pixels. A device
//! that can blend in CMYK takes such a group through [`Device::draw_isolated_cmyk_group`] and
//! interprets it with [`CmykGroup::interpret`] as often as it needs; every other device, and a
//! device that declines, gets the group's contents as before.

use crate::context::{Context, InterpreterCache};
use crate::device::Device;
use crate::interpret::state::State;
use crate::interpreter_settings::InterpreterSettings;
use crate::x_object::form::FormXObject;
use hayro_syntax::content::TypedIter;
use hayro_syntax::object::dict::keys::{CS, GROUP};
use hayro_syntax::object::{Dict, Name};
use hayro_syntax::page::Resources;
use hayro_syntax::xref::XRef;
use kurbo::{Affine, Rect, Shape};
use std::ops::Deref;

/// An isolated, non-knockout transparency group whose blending colour space is `DeviceCMYK`,
/// captured where it is painted (PdfCraft patch).
pub struct CmykGroup<'a, 'b> {
    form: &'b FormXObject<'a>,
    resources: &'b Resources<'a>,
    /// The graphics state the group's contents start with: the state at `Do`, after the
    /// form's `/Matrix`, with the group's own opacity, soft mask and blend mode already taken
    /// by the layer the group is composited through.
    state: State<'a>,
    root_transform: Affine,
    bbox: Rect,
    interpreter_cache: InterpreterCache<'a>,
    settings: InterpreterSettings,
    xref: &'a XRef,
    nesting_depth: u32,
}

impl<'a, 'b> CmykGroup<'a, 'b> {
    pub(crate) fn new(
        form: &'b FormXObject<'a>,
        resources: &'b Resources<'a>,
        context: &Context<'a>,
    ) -> Self {
        Self {
            form,
            resources,
            state: context.snapshot_state(),
            root_transform: context.root_transform(),
            bbox: context.bbox(),
            interpreter_cache: context.interpreter_cache.clone(),
            settings: context.settings.clone(),
            xref: context.xref,
            nesting_depth: context.nesting_depth(),
        }
    }

    /// The device-space box the group can paint: its `/BBox` under the current transform,
    /// within the area being drawn.
    pub fn device_bbox(&self) -> Rect {
        let [x0, y0, x1, y1] = self.form.bbox;
        let user = Rect::new(x0 as f64, y0 as f64, x1 as f64, y1 as f64);
        let device = (self.state.ctm * user.to_path(0.1)).bounding_box();
        device.intersect(self.bbox)
    }

    /// Interpret the group's contents into `device`, with device space moved by `offset`
    /// (so a device covering only [`Self::device_bbox`] can start at its corner).
    pub fn interpret(&self, device: &mut impl Device<'a>, offset: (f64, f64)) {
        let shift = Affine::translate((-offset.0, -offset.1));
        let mut state = self.state.clone();
        state.ctm = shift * state.ctm;
        let mut context = Context::new_with(
            shift * self.root_transform,
            shift.transform_rect_bbox(self.bbox),
            &self.interpreter_cache,
            self.xref,
            self.settings.clone(),
            state,
            self.nesting_depth,
        );
        let iter = TypedIter::new(self.form.decoded.as_ref());
        crate::interpret::interpret(iter, self.resources, &mut context, device);
    }
}

/// Whether `form`'s transparency group blends in `DeviceCMYK`: it says so in its `/CS`.
pub(crate) fn blends_in_device_cmyk(form: &Dict<'_>) -> bool {
    form.get::<Dict<'_>>(GROUP)
        .and_then(|g| g.get::<Name<'_>>(CS))
        .is_some_and(|cs| matches!(cs.deref(), b"DeviceCMYK" | b"CMYK"))
}

/// Whether `form`'s transparency group names a blending colour space of its own (`/CS`).
pub(crate) fn names_a_blending_space(form: &Dict<'_>) -> bool {
    form.get::<Dict<'_>>(GROUP).is_some_and(|g| g.contains_key(CS))
}
```
(Verify module paths at apply time: `InterpreterSettings` re-exports via `crate::InterpreterSettings` and `State` via `crate::interpret::state::State`; `form` module is `crate::x_object::form` — adjust `use` paths to the actual module layout if `interpreter_settings` is not a standalone module. `Context::new_with` and the `snapshot_state` accessor above are the load-bearing plumbing.)

**`vendor/hayro-interpret/src/device.rs`** — add at top `use crate::CmykGroup;` (lib.rs re-export) and insert into `trait Device<'a>` before `begin_marked_content`:
```rust
    /// PdfCraft patch: whether a transparency group drawn now is an element of a knockout
    /// group that the device wants cleared under the group's shape first. If so the
    /// interpreter draws the group's content once between `begin_knockout_erase` and
    /// `end_knockout_erase`, in which every paint clears what is under its shape instead of
    /// painting, and then draws the group as usual, in a layer of its own.
    fn knockout_erase_wanted(&mut self) -> bool {
        false
    }
    /// PdfCraft patch: see `knockout_erase_wanted`.
    fn begin_knockout_erase(&mut self) {}
    /// PdfCraft patch: see `knockout_erase_wanted`.
    fn end_knockout_erase(&mut self) {}
    /// PdfCraft patch: draw the contents of an isolated transparency group that blends in
    /// `DeviceCMYK` (ISO 32000-2 §11.3.4), already inside the layer it is composited
    /// through. Return `true` when the device drew them; on `false` (the default) the
    /// interpreter draws the contents into the device as for any other group.
    fn draw_isolated_cmyk_group(&mut self, _group: &CmykGroup<'a, '_>) -> bool {
        false
    }
    /// PdfCraft patch: an isolated transparency group that names a blending colour space
    /// other than `DeviceCMYK` is about to be drawn. A device drawing an enclosing group in
    /// CMYK can no longer draw that group right, since this one blends in its own space.
    fn group_blends_outside_device_cmyk(&mut self) {}
```
**`vendor/hayro-interpret/src/x_object/form.rs`** — replace `FormXObject::draw` (neutral-layer skip, knockout erase pass, CMYK hooks; F1 decode change above also lands here):
```rust
    pub(crate) fn draw(
        &self,
        resources: &Resources<'a>,
        context: &mut Context<'a>,
        device: &mut impl Device<'a>,
    ) {
        if !context.ocg_state.is_visible() {
            return;
        }

        if !context.begin_nested_interpretation() {
            return;
        }

        let has_oc = xobject_oc(&self.dict, context);
        if !context.ocg_state.is_visible() {
            if has_oc {
                context.ocg_state.end_marked_content();
            }
            context.end_nested_interpretation();
            return;
        }

        let iter = TypedIter::new(self.decoded.as_ref());

        context.path_mut().truncate(0);
        context.save_state();
        context.pre_concat_affine(self.matrix);
        context.push_root_transform();

        // PdfCraft patch: a transparency group that is neither isolated nor knockout composites
        // its elements with the backdrop already on the page (§11.4.7), so when it is entered at
        // full opacity with no soft mask and Normal blending there is nothing for a layer of its
        // own to add, and drawing straight onto the page is the exact result. A layer here was
        // wrong, not just slower: a blend mode inside the group saw the layer's transparent
        // backdrop instead of the content beneath, and the layer then covered the page as an
        // opaque block. An isolated group starts from a blank backdrop by definition, and a
        // knockout group or one entered at partial opacity, under a soft mask or with a blend
        // mode keeps its layer.
        let entered_plainly = {
            let state = context.fill_paint();
            let stroke = context.stroke_paint();
            state.alpha >= 1.0
                && stroke.alpha >= 1.0
                && context.compositing().soft_mask.is_none()
                && context.compositing().blend_mode == crate::BlendMode::Normal
        };
        let own_layer = self.is_transparency_group && (self.isolated || self.knockout || !entered_plainly);

        // PdfCraft patch: a group drawn as an element of a knockout group replaces what that
        // group drew under its shape. The shape is cleared by drawing the group's content once
        // in the device's erase mode; the group then composites over the cleared area in a
        // layer of its own.
        let knockout_element = self.is_transparency_group && device.knockout_erase_wanted();
        if knockout_element {
            context.save_state();
            device.begin_knockout_erase();
            device.push_clip_path(&ClipPath {
                path: context.ctm()
                    * Rect::new(
                        self.bbox[0] as f64,
                        self.bbox[1] as f64,
                        self.bbox[2] as f64,
                        self.bbox[3] as f64,
                    )
                    .to_path(0.1),
                fill: FillRule::NonZero,
            });
            interpret(
                iter.clone(),
                self.resources.as_ref().unwrap_or(resources),
                context,
                device,
            );
            device.pop_clip();
            device.end_knockout_erase();
            context.restore_state(device);
            context.path_mut().truncate(0);
        }

        if own_layer {
            device.push_transparency_group(
                context.fill_paint().alpha,
                std::mem::take(&mut context.compositing_mut().soft_mask),
                std::mem::take(&mut context.compositing_mut().blend_mode),
                self.isolated,
                self.knockout,
            );

            context.fill_paint_mut().alpha = 1.0;
            context.stroke_paint_mut().alpha = 1.0;
        }

        // PdfCraft patch: an isolated group that names another blending space blends in that
        // space (§11.3.4); a device drawing an enclosing group in CMYK is told, so it can
        // decline that group. A non-isolated group's /CS is ignored: it blends in its parent's
        // space (ISO 32000-2 Table 145, CS).
        if self.is_transparency_group
            && self.isolated
            && names_a_blending_space(&self.dict)
            && !blends_in_device_cmyk(&self.dict)
        {
            device.group_blends_outside_device_cmyk();
        }

        // PdfCraft patch: an isolated group blending in DeviceCMYK starts from a backdrop of
        // its own, so a device able to blend in CMYK computes its interior apart and composites
        // the result through the layer pushed above (see `cmyk_group`). Knockout groups are
        // left out: their elements composite with the group's initial backdrop, not with each
        // other.
        let blended_in_cmyk = own_layer
            && self.isolated
            && !self.knockout
            && blends_in_device_cmyk(&self.dict)
            && device.draw_isolated_cmyk_group(&CmykGroup::new(
                self,
                self.resources.as_ref().unwrap_or(resources),
                context,
            ));

        if !blended_in_cmyk {
            device.push_clip_path(&ClipPath {
                path: context.ctm()
                    * Rect::new(
                        self.bbox[0] as f64,
                        self.bbox[1] as f64,
                        self.bbox[2] as f64,
                        self.bbox[3] as f64,
                    )
                    .to_path(0.1),
                fill: FillRule::NonZero,
            });

            interpret(
                iter,
                self.resources.as_ref().unwrap_or(resources),
                context,
                device,
            );

            device.pop_clip();
        }

        if own_layer {
            device.pop_transparency_group();
        }

        context.pop_root_transform();
        context.restore_state(device);

        if has_oc {
            context.ocg_state.end_marked_content();
        }

        context.end_nested_interpretation();
    }
```
Imports to add in form.rs: `use crate::cmyk_group::{CmykGroup, blends_in_device_cmyk, names_a_blending_space};` and `use crate::interpret::interpret;` (the file already imports `interpret` and `ClipPath`/`FillRule`). NOTE: `iter.clone()` requires `TypedIter: Clone` — upstream type3.rs does `iter.clone()`, so it is.

### F7 (renderer side) + F10 + F2 + glue-cap — `vendor/hayro/src/`
**`src/lib.rs`** changes:
1. Module decl: add `mod cmyk;` next to the other mods.
2. `Renderer` struct — add fields:
```rust
pub(crate) struct Renderer<'a> {
    global: &'a GlobalState,
    pub(crate) ctx: &'a mut RenderContext,
    pub(crate) inside_pattern: bool,
    pub(crate) soft_mask_cache: FxHashMap<u128, Mask>,
    pub(crate) in_type3_glyph: bool,
    // PdfCraft patch: knockout machinery (see `set_knockout`) and the ink pass of a group
    // blending in DeviceCMYK (see `cmyk`).
    pub(crate) knockout: Vec<bool>,
    pub(crate) erasing: bool,
    pub(crate) erase_groups: usize,
    pub(crate) ink: cmyk::InkState,
}
```
3. `Renderer::new` — init the new fields:
```rust
    pub(crate) fn new(ctx: &'r mut RenderContext, global: &'r GlobalState) -> Self {
        Self {
            global,
            ctx,
            inside_pattern: false,
            soft_mask_cache: FxHashMap::default(),
            in_type3_glyph: false,
            knockout: Vec::new(),
            erasing: false,
            erase_groups: 0,
            ink: cmyk::InkState::default(),
        }
    }
```
4. After `misses_target`, add helpers:
```rust
    // PdfCraft patch: whether the innermost transparency group is a knockout group.
    pub(crate) fn in_knockout(&self) -> bool {
        self.knockout.last().copied().unwrap_or(false)
    }
```
and at the end of the file (before `convert_fill_rule`):
```rust
/// PdfCraft patch: inside a knockout group every element replaces what it covers: elements are
/// drawn with `Compose::Copy`, which (weighted by coverage) replaces what earlier elements drew
/// where the element's shape is, so each element composites with the group's initial backdrop
/// (transparent: the non-isolated approximation every group layer here already makes) instead of
/// with the elements before it. Against a transparent backdrop the element's own blend mode has
/// nothing to blend with, so the mix is Normal. This covers paths, strokes, glyph outlines,
/// images and shadings drawn directly; elements drawn into a layer of their own are preceded by
/// a knockout erase pass (see `Device::begin_knockout_erase`).
pub(crate) fn knockout_blend_mode() -> peniko::BlendMode {
    peniko::BlendMode::new(Mix::Normal, Compose::Copy)
}
```
5. `apply_draw_props` / `apply_image_props` — knockout + erase overrides:
```rust
    fn apply_draw_props(&mut self, props: &DrawProps<'_>) {
        self.ctx.set_transform(props.transform);
        // PdfCraft patch: inside a knockout group every element replaces what it covers; in an
        // erase pass it clears what is under its shape instead of painting.
        if self.erasing {
            self.ctx.set_blend_mode(knockout_blend_mode());
        } else if self.in_knockout() {
            self.ctx.set_blend_mode(knockout_blend_mode());
        } else {
            self.ctx.set_blend_mode(convert_blend_mode(props.blend_mode));
        }
        self.apply_soft_mask(props.soft_mask.as_ref());
    }

    fn apply_image_props(&mut self, props: &ImageDrawProps<'_>) {
        self.ctx.set_transform(props.transform);
        // PdfCraft patch: see `apply_draw_props`.
        if self.erasing || self.in_knockout() {
            self.ctx.set_blend_mode(knockout_blend_mode());
        } else {
            self.ctx.set_blend_mode(convert_blend_mode(props.blend_mode));
        }
        self.apply_soft_mask(props.soft_mask.as_ref());
    }
```
6. Device impl — replace `push_transparency_group`:
```rust
    fn push_transparency_group(
        &mut self,
        opacity: f32,
        mask: Option<SoftMask<'a>>,
        blend_mode: BlendMode,
        isolated: bool,
        knockout: bool,
    ) {
        // PdfCraft patch: nothing is layered in a knockout erase pass.
        if self.erasing {
            self.erase_groups += 1;
            return;
        }
        let _ = isolated; // PdfCraft patch: group layers here already start from a blank
                          // backdrop, which is the isolated behaviour.
        Self::push_transparency_group(self, opacity, mask, blend_mode, knockout);
    }

    fn pop_transparency_group(&mut self) {
        // PdfCraft patch: nothing is layered in a knockout erase pass.
        if self.erasing && self.erase_groups > 0 {
            self.erase_groups -= 1;
            return;
        }
        self.knockout.pop();
        Self::pop_transparency_group(self);
    }
```
7. `render()` — pixel-dimension caps (u32 pre-clamp before the u16 casts):
```rust
pub fn render<'a>(
    page: &'a Page<'a>,
    cache: &RenderCache<'a>,
    interpreter_settings: &InterpreterSettings,
    render_settings: &RenderSettings,
    pixmap_settings: &PixmapSettings,
) -> Pixmap {
    let (width, height) = page.render_dimensions();
    // PdfCraft patch: the pixel dimensions are pre-clamped in u32 before the u16 casts, so a
    // hostile page box cannot produce a 65535 × 65535 canvas: the scales are lowered until the
    // output fits both the backend's sides and a 2^28-pixel budget (the renderer's image cap).
    const MAX_RENDER_PIXELS: u64 = 1 << 28;
    let mut x_scale = pixmap_settings.x_scale;
    let mut y_scale = pixmap_settings.y_scale;
    if x_scale.is_finite() && y_scale.is_finite() && x_scale > 0.0 && y_scale > 0.0 {
        let w = (width.max(0.0) as f64 * x_scale as f64) as u64;
        let h = (height.max(0.0) as f64 * y_scale as f64) as u64;
        let fit_side = |s: f32| {
            let max = u64::from(u16::MAX);
            if (height.max(0.0) as f64 * s as f64) as u64 > max {
                max as f32 / height.max(1.0) as f32
            } else {
                s
            }
        };
        x_scale = fit_side(x_scale);
        y_scale = fit_side(y_scale);
        let pixels = (width.max(0.0) as f64 * x_scale as f64) as u64
            * (height.max(0.0) as f64 * y_scale as f64) as u64;
        if pixels > MAX_RENDER_PIXELS {
            let shrink = (MAX_RENDER_PIXELS as f64 / pixels.max(1) as f64).sqrt();
            x_scale *= shrink as f32;
            y_scale *= shrink as f32;
        }
    } else {
        x_scale = 0.0;
        y_scale = 0.0;
    }
    let mut ctx = RenderContext::new(
        (width.max(0.0) * x_scale) as u16,
        (height.max(0.0) * y_scale) as u16,
    );
    let transform = Affine::scale_non_uniform(x_scale as f64, y_scale as f64)
        * page.initial_transform(true).to_kurbo();
    render_into(
        page,
        cache,
        interpreter_settings,
        render_settings,
        &mut ctx,
        transform,
    );
    ctx.flush();

    let mut pixmap = Pixmap::new(ctx.width(), ctx.height());
    ctx.render_with(
        &mut pixmap,
        &mut vello_cpu::Resources::default(),
        vello_cpu::RasterizerSettings {
            target_init: vello_cpu::TargetInit::Clear(pixmap_settings.bg_color),
            ..Default::default()
        },
    );
    pixmap
}
```

**`src/mask.rs`** — replace `push_transparency_group`, `pop_transparency_group`, `apply_soft_mask`:
```rust
    pub(super) fn apply_soft_mask(&mut self, mask: Option<&SoftMask<'_>>) {
        // PdfCraft patch: an erase pass clears, it does not apply soft masks.
        if self.erasing {
            self.ctx.reset_mask();
            return;
        }
        let settings = *self.ctx.render_settings();
        let global = self.global;
        let mask = mask.map(|m| {
            let width = self.ctx.width();
            let height = self.ctx.height();

            self.soft_mask_cache
                .entry(m.cache_key())
                .or_insert_with(|| draw_soft_mask(m, settings, width, height, global))
                .clone()
        });

        if let Some(mask) = mask {
            self.ctx.set_mask(mask);
        } else {
            self.ctx.reset_mask();
        }
    }

    pub(super) fn push_transparency_group(
        &mut self,
        opacity: f32,
        mask: Option<SoftMask<'_>>,
        blend_mode: BlendMode,
        knockout: bool,
    ) {
        let settings = *self.ctx.render_settings();
        let global = self.global;
        self.ctx.push_layer(
            None,
            Some(convert_blend_mode(blend_mode)),
            Some(opacity),
            mask.map(|m| {
                let width = self.ctx.width();
                let height = self.ctx.height();

                self.soft_mask_cache
                    .entry(m.cache_key())
                    .or_insert_with(|| draw_soft_mask(&m, settings, width, height, global))
                    .clone()
            }),
            None,
        );
        // PdfCraft patch: the group's elements knock each other out (#2971).
        self.knockout.push(knockout);
        if knockout {
            self.ctx.set_blend_mode(knockout_blend_mode());
        }
    }

    pub(super) fn pop_transparency_group(&mut self) {
        self.ctx.pop_layer();
        if !self.knockout.is_empty() {
            // PdfCraft patch: restore the knockout blend mode of the enclosing group, if any.
            if self.in_knockout() {
                self.ctx.set_blend_mode(knockout_blend_mode());
            }
        }
    }

    /// PdfCraft patch: clear what the knockout group drew under `shape` (a transparent fill
    /// with `Compose::Copy`, weighted by coverage) before a nested element drawn into a layer
    /// of its own composites there, so the element replaces it.
    pub(crate) fn knockout_erase_shape(&mut self, shape: &BezPath) {
        if !self.in_knockout() || self.erasing {
            return;
        }
        let old_transform = *self.ctx.transform();
        let old_rule = *self.ctx.fill_rule();
        self.ctx.set_transform(Affine::IDENTITY);
        self.ctx.set_fill_rule(Fill::NonZero);
        self.ctx.set_blend_mode(knockout_blend_mode());
        self.ctx
            .set_paint(AlphaColor::<Srgb>::new([0.0, 0.0, 0.0, 0.0]));
        self.ctx.fill_path(shape);
        self.ctx.set_fill_rule(old_rule);
        self.ctx.set_transform(old_transform);
    }
```
Add imports to mask.rs top: `use kurbo::{Affine, BezPath};`, `use vello_cpu::peniko::{Compose, Fill, Mix};`, `use crate::knockout_blend_mode;` (merge with existing use lines).

**`src/paint.rs`** — erase override at the top of `set_paint` (insert as the first statements of the function body; the rest unchanged):
```rust
        // PdfCraft patch: in a knockout erase pass every paint clears what is under its shape
        // instead of painting: the path region is identical, only the paint differs.
        if self.erasing {
            self.ctx.set_paint_transform(Affine::IDENTITY);
            self.ctx
                .set_paint(AlphaColor::<Srgb>::new([0.0, 0.0, 0.0, 0.0]).into());
            return None;
        }
```
(`PaintType: From<AlphaColor<Srgb>>` — upstream uses `.into()` for colors at the top of the same fn, so the conversion exists.)

**`src/path.rs`** — replace `fill_path` and the `DrawMode::Fill` arm of `draw_rect`:
```rust
    pub(super) fn fill_path(&mut self, path: &BezPath, props: DrawProps<'_>, fill_rule: FillRule) {
        // PdfCraft patch: a fill whose control box (which contains the flattened outline)
        // misses the canvas paints nothing: a tile of a large page no longer fills every path.
        if self
            .misses_target(props.transform.transform_rect_bbox(path.control_box()))
        {
            return;
        }
        self.ctx.set_fill_rule(convert_fill_rule(fill_rule));
        self.apply_draw_props(&props);

        let clip_path = self.set_paint(&props.paint, || path.bounding_box(), false);
        if let Some(clip_path) = clip_path.as_ref() {
            self.push_clip_path_inner(clip_path, fill_rule);
        }

        self.ctx.fill_path(path);

        if clip_path.is_some() {
            self.ctx.pop_clip();
        }
    }
```
```rust
            DrawMode::Fill(fill_rule) => {
                // PdfCraft patch: skip rectangles that miss the canvas (see `fill_path`).
                if self.misses_target(props.transform.transform_rect_bbox(*rect)) {
                    return;
                }
                self.ctx.set_fill_rule(convert_fill_rule(*fill_rule));
                self.apply_draw_props(&props);

                let clip_path = self.set_paint(&props.paint, || *rect, false);
                if let Some(clip_path) = clip_path.as_ref() {
                    self.push_clip_path_inner(clip_path, *fill_rule);
                }

                self.ctx.fill_rect(rect);

                if clip_path.is_some() {
                    self.ctx.pop_clip();
                }
            }
```
Add `use crate::RenderMode`? no — nothing new needed; `control_box` comes from `Shape` (already imported in path.rs).

**`src/glyph.rs`** — add the run-bbox guard. Insert helpers + guards:
```rust
    /// PdfCraft patch: the device-space box a glyph run paints (the union of its glyphs'
    /// outline boxes under the run's transforms). A run that misses the canvas is skipped.
    fn run_bbox(&self, glyph_run: &GlyphRun<'_, '_>, transform: Affine) -> Rect {
        glyph_run
            .glyphs()
            .iter()
            .filter_map(|glyph| {
                let Glyph::Outline(outline) = &**glyph else {
                    return None;
                };
                let path = self.cached_outline(outline, outline.identifier().cache_key());
                let combined = transform * glyph.transform();
                Some(combined.transform_rect_bbox(path.bounding_box()))
            })
            .reduce(|a, b| a.union(b))
            .unwrap_or(Rect::ZERO)
    }
```
In `fill_outline_run`, as the first statement:
```rust
        // PdfCraft patch: skip runs that miss the canvas (see `run_bbox`).
        let run_bbox = self.run_bbox(glyph_run, props.transform);
        if self.misses_target(run_bbox) {
            return;
        }
```
In `stroke_outline_run`, as the first statement:
```rust
        // PdfCraft patch: skip runs that miss the canvas; the paint's reach past the outlines
        // is half the (miter-limited) stroke width, grown for flattening slack.
        let extent = f64::from(stroke_props.line_width)
            * stroke_props.miter_limit.max(1.42)
            * 0.5
            + 1.0;
        if self.misses_target(
            self.run_bbox(glyph_run, props.transform).inflate(extent, extent),
        ) {
            return;
        }
```

**`src/image.rs`** — F2 (fallible resampling) + F7 (erase/knockout guards). Add the budget block after `RESAMPLING_FUNCTION`:
```rust
// PdfCraft patch: largest image drawn, in pixels (2^28 ≈ 268 MP, well above real page images).
const MAX_IMAGE_PIXELS: u64 = 1 << 28;

// PdfCraft patch: source axes allowed into resampling planning. Unlike the final pixmap,
// sources need not fit u16: a 65536-pixel strip can safely shrink to a 16-pixel image.
const MAX_RESAMPLING_SOURCE_SIDE: u32 = 1 << 20;
const MAX_RESAMPLING_PLAN_BYTES: usize = 128 << 20;

/// PdfCraft patch: source, intermediate and RGBA byte lengths, independent of backend sides.
fn source_byte_len(width: u32, height: u32, channels: usize, max_pixels: u64) -> Option<usize> {
    if width == 0 || height == 0 || !matches!(channels, 1 | 3 | 4) {
        return None;
    }
    let pixels = u64::from(width).checked_mul(u64::from(height))?;
    if pixels > max_pixels {
        return None;
    }
    let pixels = usize::try_from(pixels).ok()?;
    let rgba_len = pixels.checked_mul(4)?;
    if rgba_len > isize::MAX as usize {
        return None;
    }
    pixels.checked_mul(channels)
}

/// PdfCraft patch: only final pixmaps/padded glyphs need to fit the backend's u16 sides.
fn image_byte_len(width: u32, height: u32, channels: usize, max_pixels: u64) -> Option<usize> {
    if width > u32::from(u16::MAX) || height > u32::from(u16::MAX) {
        return None;
    }
    source_byte_len(width, height, channels, max_pixels)
}

/// PdfCraft patch: conservative resampling planning work for one axis in PicScale 0.7.12. Its
/// four-tap kernels grow on downscaling. Include float/32-aligned i16 weights, temporary
/// float/f64/usize tap arrays and two copies of the bounds. This is a size estimate, not a
/// fallible reserve inside PicScale or a bound on total renderer memory.
fn filter_plan_bytes(source: u32, target: u32) -> Option<usize> {
    if source == target {
        return Some(0);
    }
    let (source, target) = (usize::try_from(source).ok()?, usize::try_from(target).ok()?);
    if target == 0 {
        return None;
    }
    let taps = source.checked_mul(4)?.div_ceil(target).max(4);
    let aligned_taps = taps.div_ceil(32).checked_mul(32)?;
    let float_weights = taps.checked_mul(target)?.checked_mul(size_of::<f32>())?;
    let fixed_weights = aligned_taps
        .checked_mul(target)?
        .checked_mul(size_of::<i16>())?;
    let tap_scratch = taps.checked_mul(size_of::<f32>() + size_of::<f64>() + size_of::<usize>())?;
    let bounds = target.checked_mul(4)?.checked_mul(size_of::<usize>())?;
    float_weights
        .checked_add(fixed_weights)?
        .checked_add(tap_scratch)?
        .checked_add(bounds)
}

/// PdfCraft patch: alpha-aware RGBA plans also need source-sized premultiplication scratch.
/// Bound the combined scratch before planning, including a conservative crossed intermediate
/// (or four rows for a short target) when both convolution axes change.
fn resampling_scratch_byte_len(
    source: (u32, u32),
    target: (u32, u32),
    channels: usize,
    max_pixels: u64,
) -> Option<usize> {
    if source == target {
        return Some(0);
    }
    let alpha = if channels == 4 {
        source_byte_len(source.0, source.1, channels, max_pixels)?
    } else {
        0
    };
    let filter = if source.0 != target.0 && source.1 != target.1 {
        source_byte_len(source.0, target.1.max(4), channels, max_pixels)?
    } else {
        0
    };
    let len = alpha.checked_add(filter)?;
    let limit = max_pixels.saturating_mul(4).min(isize::MAX as u64);
    if u64::try_from(len).ok()? > limit {
        return None;
    }
    Some(len)
}

/// PdfCraft patch: validate output, crossed intermediate and filter-work bounds before planning.
/// Source width × target height also bounds PicScale's smaller single-threaded four-row scratch.
fn resampling_byte_len(
    source: (u32, u32),
    target: (u32, u32),
    channels: usize,
    data_len: usize,
    max_pixels: u64,
) -> Option<usize> {
    if source_byte_len(source.0, source.1, channels, max_pixels)? != data_len {
        return None;
    }
    let target_len = image_byte_len(target.0, target.1, channels, max_pixels)?;
    resampling_scratch_byte_len(source, target, channels, max_pixels)?;
    source_byte_len(source.0, target.1, channels, max_pixels)?;
    if source.0 > MAX_RESAMPLING_SOURCE_SIDE || source.1 > MAX_RESAMPLING_SOURCE_SIDE {
        return None;
    }
    let plan_bytes = filter_plan_bytes(source.0, target.0)?
        .checked_add(filter_plan_bytes(source.1, target.1)?)?;
    if plan_bytes > MAX_RESAMPLING_PLAN_BYTES {
        return None;
    }
    Some(target_len)
}

/// PdfCraft patch: allocation failures in image conversion/resampling keep the renderer alive.
fn image_buffer(len: usize) -> Option<Vec<u8>> {
    let mut out = Vec::new();
    out.try_reserve_exact(len).ok()?;
    Some(out)
}
```
Add `use std::mem::size_of;` to the imports. Replace `resize_image_data`, `resize_image_data_impl`, `downscale_with_alpha`:
```rust
fn resize_image_data(
    scaler: &Scaler,
    data: Vec<u8>,
    src_width: u32,
    src_height: u32,
    new_width: u32,
    new_height: u32,
    pixel_format: ImagePixelFormat,
) -> Option<Vec<u8>> {
    match pixel_format {
        ImagePixelFormat::Luma => resize_image_data_impl::<1>(
            scaler, data, src_width, src_height, new_width, new_height,
            |scaler, source_size, target_size| {
                scaler.plan_planar_resampling(source_size, target_size)
            },
        ),
        ImagePixelFormat::Rgb => resize_image_data_impl::<3>(
            scaler, data, src_width, src_height, new_width, new_height,
            |scaler, source_size, target_size| scaler.plan_rgb_resampling(source_size, target_size),
        ),
        ImagePixelFormat::PremultipliedRgba => resize_image_data_impl::<4>(
            scaler, data, src_width, src_height, new_width, new_height,
            |scaler, source_size, target_size| {
                scaler.plan_rgba_resampling(source_size, target_size, false)
            },
        ),
    }
}

/// PdfCraft patch: image dimensions are validated and buffers reserved before planning; any
/// failure hands the owned input back, so the caller may retain the valid source.
fn resize_image_data_impl<const N: usize>(
    scaler: &Scaler,
    data: Vec<u8>,
    src_width: u32,
    src_height: u32,
    new_width: u32,
    new_height: u32,
    plan: impl FnOnce(&Scaler, ImageSize, ImageSize) -> Result<Arc<Resampling<u8, N>>, PicScaleError>,
) -> Option<Vec<u8>> {
    let resized = (|| {
        let len = resampling_byte_len(
            (src_width, src_height),
            (new_width, new_height),
            N,
            data.len(),
            MAX_IMAGE_PIXELS,
        )?;
        let source_size = ImageSize::new(src_width as usize, src_height as usize);
        let target_size = ImageSize::new(new_width as usize, new_height as usize);
        let src =
            ImageStore::<u8, N>::from_slice(&data, src_width as usize, src_height as usize)?;
        let plan = plan(scaler, source_size, target_size).ok()?;
        let mut out = image_buffer(len)?;
        out.resize(len, 0);
        let mut dst = ImageStoreMut::<u8, N>::from_slice(
            &mut out,
            new_width as usize,
            new_height as usize,
        )?;
        plan.resample(&src, &mut dst).ok()?;
        Some(out)
    })();
    resized
}
```
(`ImageStore::from_slice` returns Option — `?` works in the closure returning Option. `plan.resample` returns `Result<(), PicScaleError>` — `.ok()?`.) Replace `downscale_with_alpha`'s two `resize_image_data(...)` calls and its `premultiply` tail with a fallible version:
```rust
fn downscale_with_alpha<const N: usize>(
    level: Level,
    scaler: &Scaler,
    mut color: Vec<u8>,
    alpha: Vec<u8>,
    src_width: u32,
    src_height: u32,
    new_width: u32,
    new_height: u32,
) -> Option<Vec<u8>> {
    premultiply_with_alpha::<N>(level, &mut color, &alpha);

    let color_format = if N == 1 {
        ImagePixelFormat::Luma
    } else {
        ImagePixelFormat::Rgb
    };
    let color = resize_image_data(
        scaler, color, src_width, src_height, new_width, new_height, color_format,
    )?;
    let alpha = resize_image_data(
        scaler, alpha, src_width, src_height, new_width, new_height, ImagePixelFormat::Luma,
    )?;

    unpremultiply_with_alpha::<N>(level, &mut color, &alpha);
    Some(color)
}
```
(keep the existing `unpremultiply_with_alpha` tail upstream had — read the file tail at apply time; the fn's post-resize body is `unpremultiply_with_alpha::<N>(level, &mut color, &alpha); Some(color)` — match the upstream remainder exactly, only the `?` returns change.)
Then in `draw_image`, wrap the five resize/darken call sites so a `None` keeps the source data and dimensions:
- `Solid + alpha` branch:
```rust
            let alpha_data = if !needs_resize {
                alpha.data
            } else {
                // PdfCraft patch: on failure retain the source alpha.
                match resize_image_data(
                    &self.global.scaler,
                    alpha.data.clone(),
                    img_width,
                    img_height,
                    new_width,
                    new_height,
                    ImagePixelFormat::Luma,
                ) {
                    Some(resized_alpha) => {
                        additional_transform = Affine::scale_non_uniform(
                            img_width as f64 / new_width as f64,
                            img_height as f64 / new_height as f64,
                        );
                        img_width = new_width;
                        img_height = new_height;
                        resized_alpha
                    }
                    None => alpha.data,
                }
            };
```
- `Luma` without alpha: same `match` shape on `luma.data.clone()` (None arm → `luma.data`).
- `Luma` with alpha (`downscale_with_alpha`):
```rust
            } else {
                // PdfCraft patch: on failure expand at the source resolution instead.
                match downscale_with_alpha::<1>(
                    self.global.level,
                    &self.global.scaler,
                    luma.data.clone(),
                    alpha.data.clone(),
                    img_width,
                    img_height,
                    new_width,
                    new_height,
                ) {
                    Some(resized) => {
                        additional_transform = Affine::scale_non_uniform(
                            img_width as f64 / new_width as f64,
                            img_height as f64 / new_height as f64,
                        );
                        img_width = new_width;
                        img_height = new_height;
                        needs_premultiplication = false;
                        resized
                    }
                    None => luma
                        .data
                        .iter()
                        .zip(&alpha.data)
                        .flat_map(|(g, a)| [*g, *g, *g, *a])
                        .collect::<Vec<_>>(),
                }
            }
```
- `Rgb` + needs_resize branch: same match shape (None arm → expand `rgb.data` (+alpha) to RGBA at source dims exactly as the upstream no-resize arm does).
- Final RGBA resize branch: None arm → keep `rgba_data` (already premultiplied — keep `needs_premultiplication = false;` only in the Some arm).
Also add at the top of `draw_image`, after `let image_data = image_data.into();`:
```rust
        // PdfCraft patch: in a knockout erase pass an image clears its placement rectangle
        // instead of painting.
        if self.erasing {
            let cur_transform = *self.ctx.transform();
            let rect = Rect::new(
                0.0,
                0.0,
                f64::from(image_data.width()),
                f64::from(image_data.height()),
            );
            let outline = cur_transform * rect.to_path(0.1);
            self.ctx.set_blend_mode(knockout_blend_mode());
            self.knockout_erase_shape(&outline);
            return;
        }
```
And in `draw_image_with_alpha_mask`, first statement:
```rust
        // PdfCraft patch: this element composites through layers of its own, which cannot be
        // copied; clear what the knockout group drew under the image's placement first.
        if self.in_knockout() && !self.erasing {
            let rect = Rect::new(
                0.0,
                0.0,
                f64::from(image_data.width()),
                f64::from(image_data.height()),
            );
            let outline = image_transform_pre * rect.to_path(0.1);
            self.knockout_erase_shape(&outline);
        }
```
where `image_transform_pre` is the ctx transform before any changes — reorder: capture `let image_transform = *self.ctx.transform();` first (upstream already does), then use that for the erase. Add imports to image.rs: `use crate::{knockout_blend_mode};` and `kurbo::Shape` (already imported). Note the `Rect`/`Shape` are in scope; `to_path` needs `Shape` — already imported (`use kurbo::{Affine, Rect}` + Shape? upstream image.rs imports `kurbo::{Affine, Rect}` — add `Shape`).

**New file `vendor/hayro/src/cmyk.rs`** (InkState + the isolated-CMYK group pass; pixel access via the verified `into_png()` + `image` round-trip):
```rust
//! PdfCraft patch: the ink passes of an isolated transparency group that blends in
//! `DeviceCMYK` (ISO 32000-2 §11.3.4; see `hayro_interpret::CmykGroup`).

use crate::Renderer;
use vello_cpu::color::AlphaColor;
use vello_cpu::peniko::ImageQuality;
use vello_cpu::{Pixmap, RenderContext};
use std::sync::Arc;

/// PdfCraft patch: the largest CMYK group drawn in ink passes, in pixels: the passes render on
/// canvases of their own, and a larger group is drawn as before.
const MAX_CMYK_GROUP_PIXELS: u64 = 1 << 24;

/// PdfCraft patch (27): set while this renderer draws one ink pass of a group blending in
/// `DeviceCMYK`.
#[derive(Clone, Copy, Default)]
pub(crate) struct InkState {
    pass: Option<InkPass>,
    /// Something in the group has no ink tints (an RGB colour, a pattern, an image): the
    /// group is drawn as before instead.
    pub(crate) unsupported: bool,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum InkPass {
    /// Cyan, magenta and yellow as red, green and blue.
    Cmy,
    /// Black as grey.
    K,
}

impl InkState {
    /// The colour a pass paints for `color`; transparent, with the group marked unsupported,
    /// when the colour has no ink tints.
    pub(crate) fn paint(&mut self, color: &hayro_interpret::color::Color) -> AlphaColor {
        let Some((tints, alpha)) = color.ink_tints() else {
            self.unsupported = true;
            return AlphaColor::new([0.0, 0.0, 0.0, 0.0]);
        };
        match self.pass {
            Some(InkPass::Cmy) => AlphaColor::new([1.0 - tints[0], 1.0 - tints[1], 1.0 - tints[2], alpha]),
            Some(InkPass::K) => AlphaColor::new([1.0 - tints[3], 1.0 - tints[3], 1.0 - tints[3], alpha]),
            None => color.to_rgba(),
        }
    }
}

impl Renderer<'_> {
    /// PdfCraft patch (27): one ink pass of `group` on a `width` × `height` canvas whose corner
    /// is device point `offset`; `None` when the group holds something the pass has no tints
    /// for.
    pub(crate) fn ink_pass(
        &self,
        group: &hayro_interpret::CmykGroup<'_, '_>,
        pass: InkPass,
        offset: (f64, f64),
        width: u16,
        height: u16,
    ) -> Option<Pixmap> {
        let mut renderer = Renderer {
            global: self.global,
            ctx: &mut ctx_for(self, width, height),
            inside_pattern: self.inside_pattern,
            soft_mask_cache: Default::default(),
            in_type3_glyph: false,
            knockout: Vec::new(),
            erasing: false,
            erase_groups: 0,
            ink: InkState {
                pass: Some(pass),
                unsupported: false,
            },
        };
        // The pass draws inside a layer of its own, so it composites over the transparent
        // canvas instead of blending with whatever it covers.
        renderer.ctx.push_layer(None, None, None, None, None);
        group.interpret(&mut renderer, offset);
        renderer.ctx.pop_layer();
        if renderer.ink.unsupported {
            return None;
        }
        let mut pix = Pixmap::new(width, height);
        renderer.ctx.flush();
        let mut resources = vello_cpu::Resources::default();
        renderer.ctx.render(&mut pix, &mut resources);
        Some(pix)
    }

    /// PdfCraft patch (27): an isolated group that blends in `DeviceCMYK`. Its contents are
    /// drawn twice on canvases of their own, as cyan-magenta-yellow and as black, and the tints
    /// those passes leave are converted to sRGB the way the device draws solid CMYK colour. The
    /// result goes into the layer the interpreter pushed for the group, which composites it
    /// with the group's opacity, soft mask and blend mode as before.
    pub(crate) fn draw_isolated_cmyk_group(
        &mut self,
        group: &hayro_interpret::CmykGroup<'_, '_>,
    ) -> bool {
        if self.ink.pass.is_some() || self.erasing {
            return false;
        }
        let canvas = vello_cpu::peniko::kurbo::Rect::new(
            0.0,
            0.0,
            f64::from(self.ctx.width()),
            f64::from(self.ctx.height()),
        );
        let area = group.device_bbox().intersect(canvas);
        if !(area.width() > 0.0 && area.height() > 0.0) {
            // Nothing of the group is on the canvas (a NaN box lands here too).
            return true;
        }
        let (x0, y0) = (area.x0.floor(), area.y0.floor());
        // Within the canvas, so the sides fit its u16 ones.
        let width = (area.x1.ceil() - x0) as u16;
        let height = (area.y1.ceil() - y0) as u16;
        if width == 0 || height == 0 {
            return true;
        }
        if u64::from(width) * u64::from(height) > MAX_CMYK_GROUP_PIXELS {
            return false;
        }

        let Some(cmy) = self.ink_pass(group, InkPass::Cmy, (x0, y0), width, height) else {
            return false;
        };
        let Some(k) = self.ink_pass(group, InkPass::K, (x0, y0), width, height) else {
            return false;
        };

        let Some(cmy_pixels) = straight_rgba(&cmy) else {
            return false;
        };
        let Some(k_pixels) = straight_rgba(&k) else {
            return false;
        };

        // A tint is the pass's colour complemented; each pass's own alpha carries its coverage.
        let mut tints = Vec::with_capacity(usize::from(width) * usize::from(height) * 4);
        for (p, q) in cmy_pixels.chunks_exact(4).zip(k_pixels.chunks_exact(4)) {
            tints.extend_from_slice(&[255 - p[0], 255 - p[1], 255 - p[2], 255 - q[0]]);
        }
        let mut rgb = vec![0u8; tints.len() / 4 * 3];
        // PdfCraft patch: converted the same way the device draws solid DeviceCMYK colour.
        if hayro_interpret::color::cmyk_to_srgb8(&tints, &mut rgb).is_none() {
            return false;
        }

        let mut rgba = Vec::with_capacity(rgb.len() / 3 * 4);
        for (rgb, p) in rgb.chunks_exact(3).zip(cmy_pixels.chunks_exact(4)) {
            rgba.extend_from_slice(&[rgb[0], rgb[1], rgb[2], p[3]]);
        }
        let result = Pixmap::from_parts(
            rgba,
            width,
            height,
            vello_cpu::PixelMetadata {
                may_have_transparency: true,
                ..Default::default()
            },
        );

        // The group's own soft mask and blend mode are on its layer; what the contents set
        // must not apply to the finished pixels a second time.
        self.ctx.reset_mask();
        self.ctx.set_blend_mode(vello_cpu::peniko::BlendMode::default());
        self.ctx.set_paint_transform(Affine::IDENTITY);
        self.draw_pixmap(Arc::new(result), ImageQuality::Low, kurbo::Affine::translate((x0, y0)));
        true
    }

    /// PdfCraft patch (27): a nested isolated group that names its own (non-CMYK) blending
    /// space can no longer be drawn inside this pass.
    pub(crate) fn group_blends_outside_device_cmyk(&mut self) {
        if self.ink.pass.is_some() {
            self.ink.unsupported = true;
        }
    }
}

fn ctx_for(_r: &Renderer, width: u16, height: u16) -> RenderContext {
    RenderContext::new(width, height)
}

/// PdfCraft patch: straight (non-premultiplied) RGBA8 pixels of a rendered canvas, through the
/// one pixel path the vendored tree may rely on (`into_png` + the crate's own `image`
/// dependency), because vello_cpu 0.3 exposes no verified direct pixel accessor.
fn straight_rgba(pix: &Pixmap) -> Option<Vec<u8>> {
    let png = pix.clone().into_png().ok()?;
    let img = image::load_from_memory(&png).ok()?.into_rgba8();
    Some(img.into_raw())
}
```
Plus a Device dispatch: `draw_isolated_cmyk_group`/`group_blends_outside_device_cmyk` are inherent `Renderer` methods called from the Device default methods — the trait's default implementations run UNLESS overridden; the Renderer must override them in its `impl<'a,'r> Device<'a> for Renderer<'r>` in lib.rs:
```rust
    fn draw_isolated_cmyk_group(&mut self, group: &hayro_interpret::CmykGroup<'a, '_>) -> bool {
        Self::draw_isolated_cmyk_group(self, group)
    }

    fn group_blends_outside_device_cmyk(&mut self) {
        Self::group_blends_outside_device_cmyk(self)
    }

    fn knockout_erase_wanted(&mut self) -> bool {
        self.in_knockout()
    }

    fn begin_knockout_erase(&mut self) {
        self.erasing = true;
        self.erase_groups = 0;
    }

    fn end_knockout_erase(&mut self) {
        self.erasing = false;
        self.erase_groups = 0;
        if self.in_knockout() {
            self.ctx.set_blend_mode(knockout_blend_mode());
        }
    }
```
And the ink pass must colour paints: `paint.rs` `Paint::Color(c)` arm becomes `let c = self.ink.paint(&c).to_rgba8(); AlphaColor::from_rgba8(c[0], c[1], c[2], c[3]).into()` — fold that into set_paint's Color arm (replacing the upstream `c.to_rgba().to_rgba8()` — behavior identical when `ink.pass.is_none()` because `InkState::paint` returns `color.to_rgba()` then).

One correction on `ctx_for` — borrow shape: `Renderer { ctx: &mut … }` building with a helper that takes `&self` while `self.ctx` is mutably borrowed elsewhere is fine inside `ink_pass(&self, …)` since we construct a fresh RenderContext. Simplify: replace the struct-update with:
```rust
        let mut ctx = RenderContext::new(width, height);
        let mut renderer = Renderer {
            global: self.global,
            ctx: &mut ctx,
            …
        };
```
(drop `ctx_for`).