//! TASK 3 spike: deterministic layout_complexity from span boxes (text layer only).
//! Mirrors src/formats/pdf.rs atom extraction (lopdf 0.35) and computes five signals
//! per page: horizontal gutter count, table gutter density, span-overlap ratio,
//! rotation != 0, char density. Run from repo root:
//!   cargo run --release --offline --manifest-path .workflows/spikes/pdf-complexity/Cargo.toml
//! Writes data/scores.json + data/score-table.md next to this crate; prints table.

use std::collections::{BTreeMap, HashMap, HashSet};
use std::fs;
use std::path::PathBuf;
use std::process;
use std::time::{SystemTime, UNIX_EPOCH};

use lopdf::content::Content;
use lopdf::{Document, Object};

const DEFAULT_FIXTURES: [&str; 5] = [
    "tests/fixtures/pdf_scanned.pdf",
    "tests/fixtures/pdf_table_scan.pdf",
    "tests/fixtures/pdf_rot90_scan.pdf",
    "tests/fixtures/pdf_cjk_scan.pdf",
    "tests/fixtures/pdf_mixed.pdf",
];

const W_ROT: f32 = 0.30;
const W_GUTTERS: f32 = 0.25;
const W_TABLE: f32 = 0.20;
const W_OVERLAP: f32 = 0.10;
const W_DENSITY: f32 = 0.15;
const ESCALATE_AT: f32 = 0.30;

#[derive(Clone, Copy, Debug)]
struct Matrix2D { a: f32, b: f32, c: f32, d: f32, e: f32, f: f32 }

impl Matrix2D {
    fn identity() -> Self { Self { a: 1.0, b: 0.0, c: 0.0, d: 1.0, e: 0.0, f: 0.0 } }
    fn new(a: f32, b: f32, c: f32, d: f32, e: f32, f: f32) -> Self { Self { a, b, c, d, e, f } }
    fn translation(tx: f32, ty: f32) -> Self { Self::new(1.0, 0.0, 0.0, 1.0, tx, ty) }
    fn multiply(&self, o: &Matrix2D) -> Matrix2D {
        Matrix2D::new(
            self.a * o.a + self.c * o.b,
            self.b * o.a + self.d * o.b,
            self.a * o.c + self.c * o.d,
            self.b * o.c + self.d * o.d,
            self.a * o.e + self.c * o.f + self.e,
            self.b * o.e + self.d * o.f + self.f,
        )
    }
    fn apply_to_origin(&self) -> (f32, f32) { (self.e, self.f) }
    fn scale_hint(&self) -> f32 {
        let sx = (self.a * self.a + self.b * self.b).sqrt();
        let sy = (self.c * self.c + self.d * self.d).sqrt();
        sx.max(sy).max(1.0)
    }
    fn angle_deg(&self) -> f32 {
        let deg = self.b.atan2(self.a).to_degrees();
        if deg.is_finite() { deg } else { 0.0 }
    }
}

#[derive(Clone)]
struct TextState {
    in_text: bool,
    leading: f32,
    font_alias: Option<Vec<u8>>,
    font_size: f32,
    ctm: Matrix2D,
    tm: Matrix2D,
}

impl TextState {
    fn current_position(&self) -> (f32, f32) { self.ctm.multiply(&self.tm).apply_to_origin() }
    fn effective_font_size(&self) -> f32 { (self.font_size.abs() * self.tm.scale_hint()).max(1.0) }
}

#[derive(Clone, Debug)]
struct Atom { seq: u32, x: f32, y: f32, font_size: f32, text: String, angle_deg: f32 }

#[derive(Clone, Debug)]
struct LineItem { u: f32, adv: f32 }

#[derive(Clone, Debug)]
struct Line {
    font_size: f32,
    u0: f32, u1: f32, v0: f32, v1: f32,
    items: Vec<LineItem>,
}

#[derive(Clone, Debug)]
struct Signals {
    chars: usize,
    lines: usize,
    page_rot_deg: f32,
    rotated: f32,
    gutters: usize,
    gutters_norm: f32,
    aligned_cuts: usize,
    strong_cuts: usize,
    table_density: f32,
    overlap_ratio: f32,
    overlap_norm: f32,
    char_density: f32,
    density_norm: f32,
    score: f32,
    columns: usize,
}

fn obj_to_f32(o: &Object) -> Option<f32> {
    match o {
        Object::Integer(i) => Some(*i as f32),
        Object::Real(f) => Some(*f as f32),
        _ => None,
    }
}

fn two(operands: &[Object]) -> Option<(f32, f32)> {
    if operands.len() < 2 { return None; }
    match (obj_to_f32(&operands[0]), obj_to_f32(&operands[1])) {
        (Some(a), Some(b)) => Some((a, b)),
        _ => None,
    }
}

fn six(operands: &[Object]) -> Option<[f32; 6]> {
    if operands.len() < 6 { return None; }
    let mut out = [0.0f32; 6];
    for (i, slot) in out.iter_mut().enumerate() {
        *slot = obj_to_f32(&operands[i])?;
    }
    Some(out)
}

fn median_f32(mut v: Vec<f32>, default: f32) -> f32 {
    v.retain(|x| x.is_finite() && *x > 0.0);
    if v.is_empty() { return default; }
    v.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let n = v.len();
    if n % 2 == 1 { v[n / 2] } else { (v[n / 2 - 1] + v[n / 2]) / 2.0 }
}

fn clamp01(x: f32) -> f32 { x.clamp(0.0, 1.0) }

fn strip_nul(s: &str) -> String {
    if s.contains('\u{0000}') { s.chars().filter(|c| *c != '\u{0000}').collect() } else { s.to_string() }
}

fn decode_pdf_string(bytes: &[u8]) -> String {
    if bytes.len() >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF {
        let units: Vec<u16> = bytes[2..]
            .chunks_exact(2)
            .map(|c| u16::from_be_bytes([c[0], c[1]]))
            .collect();
        return strip_nul(&String::from_utf16_lossy(&units));
    }
    if let Ok(s) = std::str::from_utf8(bytes) { return s.to_string(); }
    bytes.iter().map(|b| *b as char).collect()
}

fn decode_font_string(
    encodings: &HashMap<Vec<u8>, lopdf::Encoding>,
    alias: Option<&[u8]>,
    bytes: &[u8],
) -> String {
    if let Some(alias) = alias {
        if let Some(enc) = encodings.get(alias) {
            if let Ok(s) = Document::decode_text(enc, bytes) {
                let s_stripped = strip_nul(&s);
                if s_stripped.is_empty() { return strip_nul(&decode_pdf_string(bytes)); }
                let nul_ratio = (s.chars().filter(|c| *c == '\u{0000}').count() as f32)
                    / (s.chars().count().max(1) as f32);
                if nul_ratio >= 0.20 {
                    let alt = strip_nul(&decode_pdf_string(bytes));
                    if alt.chars().count() >= s_stripped.chars().count() { return alt; }
                }
                return s_stripped;
            }
        }
    }
    strip_nul(&decode_pdf_string(bytes))
}

fn normalize_atom_text(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut last_space = false;
    for ch in s.chars() {
        if ch == '\u{0000}' { continue; }
        let is_space = ch.is_whitespace() && ch != '\n';
        if is_space {
            if !last_space { out.push(' '); last_space = true; }
            continue;
        }
        out.push(ch);
        last_space = false;
    }
    out.trim().to_string()
}

fn estimate_text_advance(text: &str, font_size: f32) -> f32 {
    if !font_size.is_finite() || font_size <= 0.0 { return 0.0; }
    let mut units = 0.0f32;
    for ch in text.chars() {
        if ch == ' ' { units += 0.33; }
        else if ch.is_ascii_punctuation() { units += 0.35; }
        else if (ch as u32) >= 0x2E80 { units += 1.0; }
        else { units += 0.55; }
    }
    (units * font_size).max(0.0)
}

fn extract_atoms(doc: &Document, page_id: lopdf::ObjectId) -> Vec<Atom> {
    let mut atoms = Vec::new();
    let Ok(content_bytes) = doc.get_page_content(page_id) else { return atoms; };
    let Ok(content) = Content::decode(&content_bytes) else { return atoms; };

    let mut fonts: HashMap<Vec<u8>, lopdf::Dictionary> = HashMap::new();
    if let Ok(page_fonts) = doc.get_page_fonts(page_id) {
        for (name, dict) in page_fonts { fonts.insert(name, dict.clone()); }
    }
    let encodings: HashMap<Vec<u8>, lopdf::Encoding> = fonts
        .iter()
        .filter_map(|(name, font)| font.get_font_encoding(doc).ok().map(|e| (name.clone(), e)))
        .collect();

    let mut state = TextState { in_text: false, leading: 0.0, font_alias: None, font_size: 0.0, ctm: Matrix2D::identity(), tm: Matrix2D::identity() };
    let mut stack: Vec<TextState> = Vec::new();

    for op in &content.operations {
        let operator = op.operator.as_str();
        let operands = &op.operands;
        match operator {
            "BT" => { state.in_text = true; state.tm = Matrix2D::identity(); }
            "ET" => state.in_text = false,
            "q" => stack.push(state.clone()),
            "Q" => { if let Some(prev) = stack.pop() { state = prev; } }
            "Tf" => {
                if operands.len() >= 2 {
                    if let Object::Name(n) = &operands[0] { state.font_alias = Some(n.clone()); }
                    if let Some(sz) = operands.get(1).and_then(obj_to_f32) {
                        state.font_size = sz.abs();
                        if state.leading <= 0.0 { state.leading = (sz.abs() * 1.2).max(1.0); }
                    }
                }
            }
            "Tm" => {
                if let Some([a, b, c, d, e, f]) = six(operands) {
                    state.tm = Matrix2D::new(a, b, c, d, e, f);
                }
            }
            "cm" => {
                if let Some([a, b, c, d, e, f]) = six(operands) {
                    let m = Matrix2D::new(a, b, c, d, e, f);
                    state.ctm = state.ctm.multiply(&m);
                }
            }
            "Td" | "TD" => {
                if let Some((tx, ty)) = two(operands) {
                    if operator == "TD" { state.leading = (-ty).abs().max(1.0); }
                    state.tm = state.tm.multiply(&Matrix2D::translation(tx, ty));
                }
            }
            "TL" => {
                if let Some(tl) = operands.first().and_then(obj_to_f32) { state.leading = tl.abs().max(1.0); }
            }
            "T*" => {
                let t = Matrix2D::translation(0.0, -state.leading);
                state.tm = state.tm.multiply(&t);
            }
            "Tj" | "'" | "\"" => {
                if !state.in_text { continue; }
                if operator != "Tj" {
                    let t = Matrix2D::translation(0.0, -state.leading);
                    state.tm = state.tm.multiply(&t);
                }
                let idx = if operator == "\"" { 2 } else { 0 };
                if let Some(Object::String(bytes, _)) = operands.get(idx) {
                    let s = normalize_atom_text(&decode_font_string(
                        &encodings, state.font_alias.as_deref(), bytes,
                    ));
                    if !s.is_empty() {
                        let fs = state.effective_font_size();
                        let (x, y) = state.current_position();
                        let angle = state.ctm.multiply(&state.tm).angle_deg();
                        atoms.push(Atom { seq: atoms.len() as u32, x, y, font_size: fs, text: s, angle_deg: angle });
                    }
                }
            }
            "TJ" => {
                if !state.in_text { continue; }
                if let Some(Object::Array(arr)) = operands.get(0) {
                    let mut s = String::new();
                    for part in arr {
                        if let Object::String(bytes, _) = part {
                            s.push_str(&decode_font_string(&encodings, state.font_alias.as_deref(), bytes));
                        }
                    }
                    let s = normalize_atom_text(&s);
                    if !s.is_empty() {
                        let fs = state.effective_font_size();
                        let (x, y) = state.current_position();
                        let angle = state.ctm.multiply(&state.tm).angle_deg();
                        atoms.push(Atom { seq: atoms.len() as u32, x, y, font_size: fs, text: s, angle_deg: angle });
                    }
                }
            }
            _ => {}
        }
    }
    atoms
}

struct PageGeometry { lines: Vec<Line>, page_rot_deg: f32 }

fn build_lines(atoms: &[Atom]) -> PageGeometry {
    let sizes = atoms.iter().map(|a| a.font_size).collect::<Vec<_>>();
    let med = median_f32(sizes, 12.0);
    let eps_y = (0.5 * med).max(0.5);

    let total_chars: usize = atoms.iter().map(|a| a.text.chars().filter(|c| !c.is_whitespace()).count()).sum();
    let mut tilted_chars = 0usize;
    let mut tilted_angles: Vec<f32> = Vec::new();
    for a in atoms {
        if a.angle_deg.abs() > 5.0 {
            tilted_chars += a.text.chars().filter(|c| !c.is_whitespace()).count();
            tilted_angles.push(a.angle_deg);
        }
    }
    let rotated = total_chars > 0 && tilted_chars * 2 >= total_chars;
    let page_rot_deg = if rotated { median_f32(tilted_angles, 0.0) } else { 0.0 };

    let (cos, sin) = (page_rot_deg.to_radians().cos(), page_rot_deg.to_radians().sin());
    let uv = |x: f32, y: f32| (x * cos + y * sin, -x * sin + y * cos);

    let mut sorted: Vec<(f32, f32, f32, u32, &Atom)> = atoms
        .iter()
        .map(|a| {
            let (u, v) = uv(a.x, a.y);
            (v, u, a.font_size, a.seq, a)
        })
        .collect();
    sorted.sort_by(|p, q| {
        q.0.partial_cmp(&p.0).unwrap_or(std::cmp::Ordering::Equal)
            .then(p.1.partial_cmp(&q.1).unwrap_or(std::cmp::Ordering::Equal))
            .then(q.2.partial_cmp(&p.2).unwrap_or(std::cmp::Ordering::Equal))
            .then(p.3.cmp(&q.3))
    });

    let mut lines: Vec<Line> = Vec::new();
    let mut cur_v = f32::NAN;
    for (v, u, _fs, _seq, a) in sorted {
        let adv = estimate_text_advance(&a.text, a.font_size);
        let starts_new = !((v - cur_v).abs() <= eps_y);
        if starts_new || lines.is_empty() {
            lines.push(Line { font_size: a.font_size, u0: u, u1: u + adv, v0: v - a.font_size / 2.0, v1: v + a.font_size / 2.0, items: vec![LineItem { u, adv }] });
            cur_v = v;
        } else {
            let line = lines.last_mut().unwrap();
            line.items.push(LineItem { u, adv });
            line.u0 = line.u0.min(u);
            line.u1 = line.u1.max(u + adv);
            line.v0 = line.v0.min(v - a.font_size / 2.0);
            line.v1 = line.v1.max(v + a.font_size / 2.0);
        }
    }
    for line in &mut lines {
        line.items.sort_by(|p, q| p.u.partial_cmp(&q.u).unwrap_or(std::cmp::Ordering::Equal));
    }
    PageGeometry { lines, page_rot_deg }
}

fn compute_signals(geom: &PageGeometry, atoms: &[Atom]) -> Signals {
    let lines = &geom.lines;
    let n = lines.len();
    let chars: usize = atoms.iter().map(|a| a.text.chars().filter(|c| !c.is_whitespace()).count()).sum();

    let mut sizes = lines.iter().map(|l| l.font_size).collect::<Vec<_>>();
    let med_fs = median_f32(std::mem::take(&mut sizes), 12.0);

    let u_min_all = lines.iter().map(|l| l.u0).fold(f32::INFINITY, f32::min);
    let u_max_all = lines.iter().map(|l| l.u1).fold(f32::NEG_INFINITY, f32::max);
    let text_width = (u_max_all - u_min_all).max(1.0);
    let body: Vec<&Line> = lines.iter().filter(|l| (l.u1 - l.u0) <= 0.7 * text_width).collect();
    let body: Vec<&Line> = if body.len() >= 3 { body } else { lines.iter().collect() };

    let gutter_min = (2.0 * med_fs).max(12.0);
    let mut iv: Vec<(f32, f32)> = body.iter().map(|l| (l.u0, l.u1)).collect();
    iv.sort_by(|p, q| p.0.partial_cmp(&q.0).unwrap_or(std::cmp::Ordering::Equal));
    let mut gutters = 0usize;
    let mut cur_end = iv.first().map(|p| p.1).unwrap_or(0.0);
    for (s, e) in iv.iter().skip(1) {
        if s - cur_end >= gutter_min { gutters += 1; }
        if *e > cur_end { cur_end = *e; }
    }
    let gutters_norm = (gutters as f32 / 2.0).min(1.0);

    let gap_min = |fs: f32| (0.6 * fs).max(2.5);
    let mut cuts: Vec<(f32, usize)> = Vec::new();
    for (li, l) in lines.iter().enumerate() {
        for w in l.items.windows(2) {
            let prev_end = w[0].u + w[0].adv;
            let gap = w[1].u - prev_end;
            if gap >= gap_min(l.font_size) {
                cuts.push(((prev_end + w[1].u) / 2.0, li));
            }
        }
    }
    cuts.sort_by(|p, q| p.0.partial_cmp(&q.0).unwrap_or(std::cmp::Ordering::Equal));
    let tol = (0.5 * med_fs).max(4.0);
    let mut clusters: Vec<Vec<usize>> = Vec::new();
    let mut last_center = f32::NAN;
    for (center, li) in &cuts {
        if clusters.is_empty() || (*center - last_center) > tol {
            clusters.push(vec![*li]);
            last_center = *center;
        } else {
            clusters.last_mut().unwrap().push(*li);
            last_center = *center;
        }
    }
    let distinct = |v: &[usize]| v.iter().collect::<HashSet<_>>().len();
    let aligned = clusters.iter().filter(|c| distinct(c) >= 3.max(((0.3 * n as f32).ceil()) as usize)).count();
    let strong_support = 3.max(((0.4 * n as f32).ceil()) as usize);
    let strong = if n >= 6 { clusters.iter().filter(|c| distinct(c) >= strong_support).count() } else { 0 };
    let table_density = (strong as f32 / 2.0).min(1.0);

    let mut involved: HashSet<usize> = HashSet::new();
    for i in 0..n {
        for j in (i + 1)..n {
            let (a, b) = (&lines[i], &lines[j]);
            let v_ov = a.v1.min(b.v1) - a.v0.max(b.v0);
            let u_ov = a.u1.min(b.u1) - a.u0.max(b.u0);
            let min_h = (a.v1 - a.v0).min(b.v1 - b.v0).max(0.001);
            let min_w = (a.u1 - a.u0).min(b.u1 - b.u0).max(0.001);
            if v_ov >= 0.5 * min_h && u_ov >= 0.3 * min_w {
                involved.insert(i);
                involved.insert(j);
            }
        }
    }
    let overlap_ratio = involved.len() as f32 / n.max(1) as f32;
    let overlap_norm = (overlap_ratio / 0.3).min(1.0);

    let rotated = if geom.page_rot_deg.abs() > 5.0 { 1.0 } else { 0.0 };

    let v_min = lines.iter().map(|l| l.v0).fold(f32::INFINITY, f32::min);
    let v_max = lines.iter().map(|l| l.v1).fold(f32::NEG_INFINITY, f32::max);
    let area = ((u_max_all - u_min_all).max(1.0)) * ((v_max - v_min).max(1.0));
    let char_density = chars as f32 / area;
    let density_norm = (char_density / 0.02).min(1.0);

    let score = clamp01(
        W_ROT * rotated + W_GUTTERS * gutters_norm + W_TABLE * table_density
            + W_OVERLAP * overlap_norm + W_DENSITY * density_norm,
    );

    Signals {
        chars, lines: n,
        page_rot_deg: geom.page_rot_deg,
        rotated, gutters, gutters_norm,
        aligned_cuts: aligned, strong_cuts: strong, table_density,
        overlap_ratio, overlap_norm,
        char_density, density_norm,
        score, columns: gutters + 1,
    }
}

fn json_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out
}

fn verdict(s: &Option<Signals>) -> &'static str {
    match s {
        None => "Escalate (ImageOnly via pdf_text_quality; score undefined)",
        Some(sig) => if sig.score >= ESCALATE_AT { "Escalate" } else { "Keep text layer" },
    }
}

fn table_markdown(rows: &[(String, u32, Option<Signals>)]) -> String {
    let mut m = String::new();
    m.push_str("| fixture | page | text_layer | chars | rot_deg | lines | gutters | aligned_cuts | strong_cuts | table_density | overlap_ratio | char_density | columns | score | verdict |\n");
    m.push_str("|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|\n");
    for (fixture, page, sig) in rows {
        match sig {
            None => m.push_str(&format!("| {} | {} | false | 0 | - | 0 | - | - | - | - | - | - | - | (undefined) | {} |\n", fixture, page, verdict(&None))),
            Some(s) => m.push_str(&format!(
                "| {} | {} | true | {} | {:.1} | {} | {} | {} | {} | {:.2} | {:.3} | {:.4} | {} | {:.3} | {} |\n",
                fixture, page, s.chars, s.page_rot_deg, s.lines, s.gutters,
                s.aligned_cuts, s.strong_cuts, s.table_density, s.overlap_ratio,
                s.char_density, s.columns, s.score, verdict(&Some(s.clone()))
            )),
        }
    }
    m
}

fn main() {
    let fixtures: Vec<String> = std::env::args().skip(1).collect();
    let fixtures = if fixtures.is_empty() {
        DEFAULT_FIXTURES.iter().map(|s| s.to_string()).collect()
    } else {
        fixtures
    };

    let mut rows: Vec<(String, u32, Option<Signals>)> = Vec::new();
    for path in &fixtures {
        let bytes = match fs::read(path) {
            Ok(b) => b,
            Err(e) => { eprintln!("FATAL: cannot read {path}: {e}"); process::exit(1); }
        };
        let doc = match Document::load_mem(&bytes) {
            Ok(d) => d,
            Err(e) => { eprintln!("FATAL: lopdf load failed for {path}: {e}"); process::exit(1); }
        };
        let pages: BTreeMap<u32, lopdf::ObjectId> = doc.get_pages();
        for (page_number, page_id) in pages {
            let atoms = extract_atoms(&doc, page_id);
            let sig = if atoms.is_empty() {
                None
            } else {
                let geom = build_lines(&atoms);
                Some(compute_signals(&geom, &atoms))
            };
            let name = path.rsplit('/').next().unwrap_or(path).to_string();
            rows.push((name, page_number, sig));
        }
    }

    let weights_line = format!(
        "score = clamp01({:.2}*rot + {:.2}*gutters_norm + {:.2}*table_density + {:.2}*overlap_norm + {:.2}*density_norm); escalate at >= {:.2}",
        W_ROT, W_GUTTERS, W_TABLE, W_OVERLAP, W_DENSITY, ESCALATE_AT
    );

    let data_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("data");
    let _ = fs::create_dir_all(&data_dir);
    let mut json = String::from("{\n  \"weights\": \"");
    json.push_str(&json_escape(&weights_line));
    json.push_str("\",\n  \"escalate_at\": ");
    json.push_str(&format!("{ESCALATE_AT}"));
    json.push_str(",\n  \"generated_unix\": ");
    json.push_str(&format!("{}", SystemTime::now().duration_since(UNIX_EPOCH).map(|d| d.as_secs()).unwrap_or(0)));
    json.push_str(",\n  \"pages\": [\n");
    for (i, (fixture, page, sig)) in rows.iter().enumerate() {
        json.push_str("    {\"fixture\": \"");
        json.push_str(&json_escape(fixture));
        json.push_str(&format!("\", \"page\": {page}, \"text_layer\": "));
        match sig {
            None => json.push_str("false, \"score\": null, \"signals\": null}"),
            Some(s) => json.push_str(&format!(
                "true, \"score\": {:.4}, \"signals\": {{\"chars\": {}, \"rot_deg\": {:.1}, \"lines\": {}, \"gutters\": {}, \"gutters_norm\": {:.3}, \"aligned_cuts\": {}, \"strong_cuts\": {}, \"table_density\": {:.3}, \"overlap_ratio\": {:.4}, \"overlap_norm\": {:.3}, \"char_density\": {:.5}, \"density_norm\": {:.3}, \"columns\": {}}}}}",
                s.score, s.chars, s.page_rot_deg, s.lines, s.gutters, s.gutters_norm,
                s.aligned_cuts, s.strong_cuts, s.table_density, s.overlap_ratio,
                s.overlap_norm, s.char_density, s.density_norm, s.columns
            )),
        }
        if i + 1 < rows.len() { json.push(','); }
        json.push('\n');
    }
    json.push_str("  ]\n}\n");
    let _ = fs::write(data_dir.join("scores.json"), &json);

    let md = format!(
        "# pdf-complexity spike - measured score table\n\nGenerated by: cargo run --release --offline --manifest-path .workflows/spikes/pdf-complexity/Cargo.toml\n\n{}\n\n{}\n",
        weights_line,
        table_markdown(&rows)
    );
    let _ = fs::write(data_dir.join("score-table.md"), &md);

    println!("{weights_line}\n");
    print!("{}", table_markdown(&rows));
    println!("\nPaste the table above into .workflows/spikes/pdf-complexity/README.md section \"Score table\" (replacing the UNMEASURED template).");
}
