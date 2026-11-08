// pdf-geometry-spike — TASK 2 prototype (.workflows/spikes/pdf-geometry only; throwaway).
//
// Question: does the existing content interpreter in src/formats/pdf.rs already hold enough
// state at its emit points to emit per-span [x0,y0,x1,y1] boxes with font_name/font_size, and
// does glyph-area / page-area separate a scanned page from a born-digital one WITHOUT pdfium?
//
// This binary REPLICATES the production walk (never imports or modifies it); every carried
// behavior cites its source line. No render-pdfium, no render-hayro, no OCR features.

use std::collections::{BTreeMap, HashMap};

use anyhow::{Context, Result};
use lopdf::{Document, Object};

// ---- geometry primitives (copied; src/formats/pdf.rs:219-272) ----
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
        Self { a: 1.0, b: 0.0, c: 0.0, d: 1.0, e: 0.0, f: 0.0 }
    }
    fn from_components(a: f32, b: f32, c: f32, d: f32, e: f32, f: f32) -> Self {
        Self { a, b, c, d, e, f }
    }
    fn multiply(&self, o: &Self) -> Self {
        // Same order as mcat's pdf_state (pdf.rs:247-256)
        Self {
            a: self.a * o.a + self.c * o.b,
            b: self.b * o.a + self.d * o.b,
            c: self.a * o.c + self.c * o.d,
            d: self.b * o.c + self.d * o.d,
            e: self.a * o.e + self.c * o.f + self.e,
            f: self.b * o.e + self.d * o.f + self.f,
        }
    }
    fn apply_to_origin(&self) -> (f32, f32) {
        (self.e, self.f) // pdf.rs:263-265
    }
    fn scale_hint(&self) -> f32 {
        // pdf.rs:267-272
        let sx = (self.a * self.a + self.b * self.b).sqrt();
        let sy = (self.c * self.c + self.d * self.d).sqrt();
        sx.max(sy).max(1.0)
    }
}

// Interpreter state — fields exactly PdfTextState (pdf.rs:208-216).
#[derive(Clone, Default)]
struct TextState {
    in_text: bool,
    leading: f32,
    font_alias: Option<Vec<u8>>,
    font_size: f32,
    ctm: Matrix2D,
    tm: Matrix2D,
}

impl TextState {
    fn bt(&mut self) {
        self.in_text = true;
        self.tm = Matrix2D::identity();
    } // pdf.rs:276-279
    fn et(&mut self) {
        self.in_text = false;
    } // pdf.rs:280-282
    fn current_position(&self) -> (f32, f32) {
        self.ctm.multiply(&self.tm).apply_to_origin()
    } // pdf.rs:283-288
    fn effective_font_size(&self) -> f32 {
        (self.font_size.abs() * self.tm.scale_hint()).max(1.0)
    } // pdf.rs:290-293
}

// The spike's emission record: TextAtom (pdf.rs:199-205) extended with the box fields the
// contract asks about. This struct IS the extend seam — additive fields at existing emit points.
struct Span {
    seq: u32,
    x0: f32,
    y0: f32,
    x1: f32,
    y1: f32,
    font_name: String,
    font_size: f32,
    text: String,
}

fn obj_to_f32(o: &Object) -> Option<f32> {
    match o {
        Object::Integer(i) => Some(*i as f32),
        Object::Real(f) => Some(*f),
        _ => None,
    }
} // pdf.rs:295-303

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
} // pdf.rs:447-465 verbatim

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
} // pdf.rs:471-485 verbatim

// Simplified from decode_pdf_string (pdf.rs:303-398): UTF-16 BOM + no-BOM heuristic + latin-1
// fallback; the digit-gluing repairs are dropped (irrelevant to boxes/coverage).
fn decode_pdf_string(bytes: &[u8]) -> String {
    if bytes.starts_with(b"\xFE\xFF") || bytes.starts_with(b"\xFF\xFE") {
        let be = bytes[0] == 0xFE;
        let mut u16s = Vec::with_capacity((bytes.len() - 2 + 1) / 2);
        let mut i = 2;
        while i + 1 < bytes.len() {
            let (hi, lo) = if be { (bytes[i], bytes[i + 1]) } else { (bytes[i + 1], bytes[i]) };
            u16s.push(u16::from_be_bytes([hi, lo]));
            i += 2;
        }
        return String::from_utf16_lossy(&u16s);
    }
    if bytes.len() >= 4 && bytes.len() % 2 == 0 {
        let even_nul = bytes.iter().step_by(2).filter(|&&b| b == 0).count();
        let odd_nul = bytes.iter().skip(1).step_by(2).filter(|&&b| b == 0).count();
        let thr = (bytes.len() / 2 * 7) / 10;
        if even_nul >= thr || odd_nul >= thr {
            let be = even_nul >= odd_nul;
            let mut u16s = Vec::with_capacity(bytes.len() / 2);
            let mut i = 0;
            while i + 1 < bytes.len() {
                let (hi, lo) = if be { (bytes[i], bytes[i + 1]) } else { (bytes[i + 1], bytes[i]) };
                u16s.push(u16::from_be_bytes([hi, lo]));
                i += 2;
            }
            return String::from_utf16_lossy(&u16s);
        }
    }
    bytes.iter().map(|&b| b as char).collect() // latin-1-ish Tier-B fallback (cf. pdf.rs:361)
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
    strip_nul(&decode_pdf_string(bytes))
} // pdf.rs:400-437 verbatim

fn decode_text_object(
    encodings: &HashMap<Vec<u8>, lopdf::Encoding>,
    font_alias: Option<&[u8]>,
    o: &Object,
) -> Option<String> {
    match o {
        Object::String(bytes, _) => Some(decode_font_string(encodings, font_alias, bytes)),
        _ => None,
    }
} // pdf.rs:439-445 verbatim

// THE NEW BIT UNDER TEST — pure spike-side arithmetic over state the interpreter already holds
// at its emit points. start = current_position() (:283-288); fs = effective_font_size()
// (:290-293); adv = estimate_text_advance() (:471-485); direction from the combined ctm*tm
// linear part (Matrix2D :219-256). For unrotated content (both fixtures) this is exactly
// [x0, y0, x0+advance, y0+fs].
fn push_span(state: &TextState, atoms: &mut Vec<Span>, t: String) {
    let (x0, y0) = state.current_position();
    let fs = state.effective_font_size();
    let adv = estimate_text_advance(&t, fs);
    let m = state.ctm.multiply(&state.tm);
    let len = (m.a * m.a + m.b * m.b).sqrt();
    let (ux, uy) = if len > 1e-6 { (m.a / len, m.b / len) } else { (1.0, 0.0) };
    let (ex, ey) = (x0 + adv * ux, y0 + adv * uy);
    let (upx, upy) = (-uy, ux);
    let corners = [
        (x0, y0),
        (ex, ey),
        (x0 + fs * upx, y0 + fs * upy),
        (ex + fs * upx, ey + fs * upy),
    ];
    let bx0 = corners.iter().map(|c| c.0).fold(f32::INFINITY, f32::min);
    let by0 = corners.iter().map(|c| c.1).fold(f32::INFINITY, f32::min);
    let bx1 = corners.iter().map(|c| c.0).fold(f32::NEG_INFINITY, f32::max);
    let by1 = corners.iter().map(|c| c.1).fold(f32::NEG_INFINITY, f32::max);
    let font_name = state
        .font_alias
        .as_deref()
        .map(|a| String::from_utf8_lossy(a).to_string())
        .unwrap_or_else(|| "(none)".to_string());
    atoms.push(Span {
        seq: atoms.len() as u32,
        x0: bx0,
        y0: by0,
        x1: bx1,
        y1: by1,
        font_name,
        font_size: fs,
        text: t,
    });
}

fn resolve_dict(doc: &Document, o: &Object) -> Option<lopdf::Dictionary> {
    match o {
        Object::Dictionary(d) => Some(d.clone()),
        Object::Reference(oid) => match doc.get_object(*oid) {
            Ok(Object::Dictionary(d)) => Some(d.clone()),
            _ => None,
        },
        _ => None,
    }
}

fn resolve_stream<'a>(doc: &'a Document, o: &'a Object) -> Option<&'a lopdf::Stream> {
    match o {
        Object::Stream(s) => Some(s),
        Object::Reference(oid) => match doc.get_object(*oid) {
            Ok(Object::Stream(s)) => Some(s),
            _ => None,
        },
        _ => None,
    }
}

fn font_dicts_from(doc: &Document, resources: &lopdf::Dictionary) -> Vec<(Vec<u8>, lopdf::Dictionary)> {
    let mut out = Vec::new();
    let Ok(font_obj) = resources.get(b"Font") else {
        return out;
    };
    let Some(fd) = resolve_dict(doc, font_obj) else {
        return out;
    };
    for (k, v) in fd.iter() {
        if let Some(d) = resolve_dict(doc, v) {
            out.push((k.clone(), d));
        }
    }
    out
}

// Simplified from build_page_resources (pdf.rs:1260-1284): both fixtures keep page-level
// Resources, so the ancestor-chain merge via get_page_resources is skipped deliberately.
fn page_resources(doc: &Document, page_id: lopdf::ObjectId) -> lopdf::Dictionary {
    let mut out = lopdf::Dictionary::new();
    let mut id = page_id;
    for _ in 0..16 {
        let Ok(Object::Dictionary(dict)) = doc.get_object(id) else {
            break;
        };
        if let Ok(res_obj) = dict.get(b"Resources") {
            if let Some(res) = resolve_dict(doc, res_obj) {
                for (k, v) in res.iter() {
                    out.set(k.clone(), v.clone()); // page-level wins (merge order per pdf.rs:1282-1284)
                }
            }
            break;
        }
        match dict.get(b"Parent") {
            Ok(Object::Reference(pid)) => id = *pid,
            _ => break,
        }
    }
    out
}

// MediaBox area from the page dict, walking /Parent for inherited boxes. Page area per spec.
fn page_area(doc: &Document, page_id: lopdf::ObjectId) -> Option<f64> {
    let mut id = page_id;
    for _ in 0..16 {
        let Ok(obj) = doc.get_object(id) else {
            return None;
        };
        let Object::Dictionary(dict) = obj else {
            return None;
        };
        if let Ok(mbox) = dict.get(b"MediaBox") {
            let arr = match mbox {
                Object::Array(a) => a.clone(),
                Object::Reference(r) => match doc.get_object(*r) {
                    Ok(Object::Array(a)) => a.clone(),
                    _ => return None,
                },
                _ => return None,
            };
            let n: Vec<f64> = arr
                .iter()
                .filter_map(|o| match o {
                    Object::Integer(i) => Some(*i as f64),
                    Object::Real(f) => Some(*f as f64),
                    _ => None,
                })
                .collect();
            if n.len() != 4 {
                return None;
            }
            return Some(((n[2] - n[0]) * (n[3] - n[1])).abs());
        }
        match dict.get(b"Parent") {
            Ok(Object::Reference(pid)) => id = *pid,
            _ => return None,
        }
    }
    None
}

// Faithful replication of handle_content_ops (pdf.rs:886-1199); push_span replaces the
// TextAtom push at each emit point. Line cites per arm.
fn walk_ops(
    doc: &Document,
    ops: &[lopdf::content::Operation],
    resources: &lopdf::Dictionary,
    fonts: &HashMap<Vec<u8>, lopdf::Dictionary>,
    encodings: &HashMap<Vec<u8>, lopdf::Encoding>,
    state: &mut TextState,
    state_stack: &mut Vec<TextState>,
    atoms: &mut Vec<Span>,
) -> Result<()> {
    for op in ops {
        let operator = op.operator.as_str();
        let operands = &op.operands;
        match operator {
            "BT" => {
                state.bt();
            } // pdf.rs:906-908
            "q" => {
                state_stack.push(state.clone());
            } // pdf.rs:909-911
            "Q" => {
                if let Some(prev) = state_stack.pop() {
                    *state = prev;
                }
            } // pdf.rs:912-916
            "ET" => {
                state.et();
            } // pdf.rs:917-919
            "Tf" => {
                // pdf.rs:920-932 — font_name source (:923) + font_size source (:925-926)
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
                // pdf.rs:933-947
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
                // pdf.rs:948-963
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
                // pdf.rs:964-973
                if operands.len() >= 2 {
                    if let (Some(tx), Some(ty)) = (obj_to_f32(&operands[0]), obj_to_f32(&operands[1])) {
                        let t = Matrix2D::from_components(1.0, 0.0, 0.0, 1.0, tx, ty);
                        state.tm = state.tm.multiply(&t);
                    }
                }
            }
            "TD" => {
                // pdf.rs:974-984
                if operands.len() >= 2 {
                    if let (Some(tx), Some(ty)) = (obj_to_f32(&operands[0]), obj_to_f32(&operands[1])) {
                        state.leading = (-ty).abs().max(1.0);
                        let t = Matrix2D::from_components(1.0, 0.0, 0.0, 1.0, tx, ty);
                        state.tm = state.tm.multiply(&t);
                    }
                }
            }
            "TL" => {
                // pdf.rs:985-989
                if let Some(tl) = operands.get(0).and_then(obj_to_f32) {
                    state.leading = tl.abs().max(1.0);
                }
            }
            "T*" => {
                // pdf.rs:990-993
                let t = Matrix2D::from_components(1.0, 0.0, 0.0, 1.0, 0.0, -state.leading);
                state.tm = state.tm.multiply(&t);
            }
            "Tj" => {
                // Emit point 1 — pdf.rs:994-1017 (fs at :1006, position at :1007)
                if !state.in_text {
                    continue;
                }
                let alias = state.font_alias.as_deref();
                if let Some(s) = operands.get(0).and_then(|o| decode_text_object(encodings, alias, o)) {
                    let t = normalize_atom_text(&s);
                    if !t.is_empty() {
                        push_span(state, atoms, t);
                    }
                }
            }
            "TJ" => {
                // Emit point 2 — pdf.rs:1018-1047 (alias at :1024)
                if !state.in_text {
                    continue;
                }
                if let Some(Object::Array(arr)) = operands.get(0) {
                    let mut s = String::new();
                    let alias = state.font_alias.as_deref();
                    for o in arr {
                        if let Some(t) = decode_text_object(encodings, alias, o) {
                            s.push_str(&t);
                        }
                    }
                    let t = normalize_atom_text(&s);
                    if !t.is_empty() {
                        push_span(state, atoms, t);
                    }
                }
            }
            "'" => {
                // Emit point 3 — pdf.rs:1048-1075
                let tmove = Matrix2D::from_components(1.0, 0.0, 0.0, 1.0, 0.0, -state.leading);
                state.tm = state.tm.multiply(&tmove);
                if !state.in_text {
                    continue;
                }
                let alias = state.font_alias.as_deref();
                if let Some(s) = operands.get(0).and_then(|o| decode_text_object(encodings, alias, o)) {
                    let t = normalize_atom_text(&s);
                    if !t.is_empty() {
                        push_span(state, atoms, t);
                    }
                }
            }
            "\"" => {
                // Emit point 4 — pdf.rs:1076-1099 (string is operands[2])
                let tmove = Matrix2D::from_components(1.0, 0.0, 0.0, 1.0, 0.0, -state.leading);
                state.tm = state.tm.multiply(&tmove);
                if !state.in_text {
                    continue;
                }
                let alias = state.font_alias.as_deref();
                if let Some(s) = operands.get(2).and_then(|o| decode_text_object(encodings, alias, o)) {
                    let t = normalize_atom_text(&s);
                    if !t.is_empty() {
                        push_span(state, atoms, t);
                    }
                }
            }
            "BDC" => {
                // Emit point 5 — ActualText, pdf.rs:1100-1133 (inline dict only; ref path kept)
                if !state.in_text {
                    continue;
                }
                if operands.len() >= 2 {
                    let dict = match &operands[1] {
                        Object::Dictionary(d) => Some(d.clone()),
                        Object::Reference(oid) => match doc.get_object(*oid) {
                            Ok(Object::Dictionary(d)) => Some(d.clone()),
                            _ => None,
                        },
                        _ => None,
                    };
                    if let Some(d) = dict {
                        if let Ok(Object::String(b, _)) = d.get(b"ActualText") {
                            let t = normalize_atom_text(&decode_pdf_string(b));
                            if !t.is_empty() {
                                push_span(state, atoms, t);
                            }
                        }
                    }
                }
            }
            "Do" => {
                // Form XObject recursion with saved state (pdf.rs:1134-1191); Image = no-op
                // for geometry (pdf.rs:1192-1198).
                let Some(Object::Name(name)) = operands.get(0) else {
                    continue;
                };
                let Some(xobjs) = resources.get(b"XObject").ok().and_then(|o| resolve_dict(doc, o))
                else {
                    continue;
                };
                let stream = match xobjs.get(name.as_slice()) {
                    Ok(obj) => match resolve_stream(doc, obj) {
                        Some(s) => s,
                        None => continue,
                    },
                    Err(_) => continue,
                };
                let subtype_form = stream
                    .dict
                    .get(b"Subtype")
                    .ok()
                    .and_then(|o| match o {
                        Object::Name(n) => std::str::from_utf8(n).ok().map(|s| s.trim() == "Form"),
                        _ => None,
                    })
                    .unwrap_or(false);
                if subtype_form {
                    let mut form_resources = resources.clone();
                    if let Ok(res_obj) = stream.dict.get(b"Resources") {
                        if let Some(res) = resolve_dict(doc, res_obj) {
                            for (k, v) in res.iter() {
                                form_resources.set(k.clone(), v.clone());
                            }
                        }
                    }
                    let mut form_fonts = fonts.clone();
                    for (n, d) in font_dicts_from(doc, &form_resources) {
                        form_fonts.insert(n, d);
                    }
                    let form_encodings: HashMap<Vec<u8>, lopdf::Encoding> = form_fonts
                        .iter()
                        .filter_map(|(n, f)| f.get_font_encoding(doc).ok().map(|e| (n.clone(), e)))
                        .collect();
                    let bytes = stream
                        .decompressed_content()
                        .unwrap_or_else(|_| stream.content.clone());
                    let Ok(content) = lopdf::content::Content::decode(&bytes) else {
                        continue;
                    };
                    let saved = state.clone();
                    let saved_stack = state_stack.clone();
                    walk_ops(
                        doc,
                        &content.operations,
                        &form_resources,
                        &form_fonts,
                        &form_encodings,
                        state,
                        state_stack,
                        atoms,
                    )?;
                    *state = saved;
                    *state_stack = saved_stack;
                }
            }
            _ => {}
        }
    }
    Ok(())
}

fn glyph_area(spans: &[Span]) -> f64 {
    spans
        .iter()
        .map(|s| ((s.x1 - s.x0).abs() as f64) * ((s.y1 - s.y0).abs() as f64))
        .sum()
}

fn one_line(s: &str) -> String {
    let flat: String = s
        .chars()
        .map(|c| if c == '\n' || c == '\r' || c == '\t' { ' ' } else { c })
        .collect();
    if flat.chars().count() <= 48 {
        flat
    } else {
        let t: String = flat.chars().take(48).collect();
        format!("{}...", t)
    }
}

fn run(fixtures: &[String]) -> Result<()> {
    println!("SPIKE pdf-geometry — per-span boxes + ink coverage without pdfium");
    let mut table = Vec::new();
    table.push("fixture\tpage\tspans\tglyph_area_pt2\tpage_area_pt2\tcoverage".to_string());
    let mut dumped = false;
    for path in fixtures {
        let bytes = std::fs::read(path).with_context(|| format!("read {}", path))?;
        let doc = Document::load_mem(&bytes).context("lopdf load_mem")?; // pdf.rs:1659
        let pages: BTreeMap<u32, lopdf::ObjectId> = doc.get_pages(); // pdf.rs:1660
        println!("fixture: {}", path);
        for (page_number, page_id) in pages {
            let resources = page_resources(&doc, page_id);
            let content_bytes = doc
                .get_page_content(page_id)
                .with_context(|| format!("get page content (page {})", page_number))?; // pdf.rs:1677
            let content = lopdf::content::Content::decode(&content_bytes)
                .with_context(|| format!("decode page content (page {})", page_number))?;
            let mut fonts: HashMap<Vec<u8>, lopdf::Dictionary> = HashMap::new();
            if let Ok(page_fonts) = doc.get_page_fonts(page_id) {
                // pdf.rs:1680
                for (name, dict) in page_fonts {
                    fonts.insert(name, dict.clone());
                }
            }
            for (n, d) in font_dicts_from(&doc, &resources) {
                fonts.entry(n).or_insert(d);
            }
            let encodings: HashMap<Vec<u8>, lopdf::Encoding> = fonts
                .iter()
                .filter_map(|(name, font)| {
                    font.get_font_encoding(&doc).ok().map(|enc| (name.clone(), enc))
                })
                .collect(); // pdf.rs:1692-1697
            let mut state = TextState::default(); // pdf.rs:1698
            let mut state_stack: Vec<TextState> = Vec::new();
            let mut spans: Vec<Span> = Vec::new();
            walk_ops(
                &doc,
                &content.operations,
                &resources,
                &fonts,
                &encodings,
                &mut state,
                &mut state_stack,
                &mut spans,
            )?;
            let area = page_area(&doc, page_id);
            let ink = glyph_area(&spans);
            match area {
                Some(a) if a > 0.0 => {
                    println!(
                        "  page {}: spans={} glyph_area={:.2} pt^2 page_area={:.2} pt^2 coverage={:.6} ({:.4}%)",
                        page_number,
                        spans.len(),
                        ink,
                        a,
                        ink / a,
                        100.0 * ink / a
                    );
                    table.push(format!(
                        "{}\t{}\t{}\t{:.2}\t{:.2}\t{:.6}",
                        path,
                        page_number,
                        spans.len(),
                        ink,
                        a,
                        ink / a
                    ));
                }
                _ => {
                    println!(
                        "  page {}: spans={} glyph_area={:.2} pt^2 page_area=n/a coverage=n/a",
                        page_number,
                        spans.len(),
                        ink
                    );
                    table.push(format!(
                        "{}\t{}\t{}\t{:.2}\tn/a\tn/a",
                        path,
                        page_number,
                        spans.len(),
                        ink
                    ));
                }
            }
            if !dumped && !spans.is_empty() {
                dumped = true;
                println!(
                    "BOX DUMP — first 20 spans of {} page {} (fields: x0 y0 x1 y1 font_name font_size \"text\"):",
                    path, page_number
                );
                for s in spans.iter().take(20) {
                    println!(
                        "  [{:>3}] x0={:8.2} y0={:8.2} x1={:8.2} y1={:8.2} font={:<12} size={:6.2} text=\"{}\"",
                        s.seq,
                        s.x0,
                        s.y0,
                        s.x1,
                        s.y1,
                        s.font_name,
                        s.font_size,
                        one_line(&s.text)
                    );
                }
            }
        }
    }
    println!("\nCOVERAGE TABLE (tab-separated)");
    for row in &table {
        println!("{}", row);
    }
    Ok(())
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if args.is_empty() {
        eprintln!("usage: pdf-geometry-spike <fixture.pdf> [more.pdf ...]");
        std::process::exit(2);
    }
    if let Err(e) = run(&args) {
        eprintln!("error: {e:#}");
        std::process::exit(1);
    }
}
