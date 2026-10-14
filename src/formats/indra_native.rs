//! Native (in-process) Indra OCR — Mode B.
//!
//! A faithful Rust port of the Indra OCR pipeline (`indra-ocr` 0.3.1,
//! Apache-2.0) over ONNX Runtime: line detection (`det.onnx`), CTC line
//! recognition (`rec.onnx` — the graph emits ids/conf/lengths directly),
//! reading order + block classes (`order.onnx`, OrderNet heads), and the
//! block assembly heuristics. No Python, no subprocess.
//!
//! Model files: `det.onnx`, `rec.onnx`, `order.onnx`, `meta.json` from
//! [datalab-to/indra-ocr](https://huggingface.co/datalab-to/indra-ocr)
//! (weights: OpenRAIL-M). Resolved from `INDRA_OCR_MODEL_DIR`, else the
//! standard Hugging Face cache is scanned.
//!
//! V1 omissions vs the Python engine (all degrade gracefully, documented in
//! the module docs of the output): orientation correction, deskew, the
//! CJK/script recognizers, math, CenterMask windows, table cell splitting,
//! and the digital-text-layer path (office-parser has its own).

use std::path::{Path, PathBuf};

use serde::Deserialize;

use crate::document_ast::Block;

pub const DET_LONG_SIDE_DEFAULT: i64 = 1024;
pub const REC_HEIGHT_DEFAULT: i64 = 32;
const INPUT_MULTIPLE: f32 = 32.0;
const OUT_STRIDE: f32 = 2.0;
const CORE_X_SHRINK: f32 = 0.15;
const LINE_HEIGHT: usize = 32;
const PAD_Y: f32 = 0.2;
const PAD_X: f32 = 0.3;
const WIDTH_MULTIPLE: usize = 8;
const REC_MAX_WIDTH: usize = 4096;
const MIN_CONF: f32 = 0.5;
const ADAPTIVE_MIN_H: f32 = 9.0;
const ADAPTIVE_TARGET_H: f32 = 12.0;
const ADAPTIVE_MAX_SIDE: f32 = 2048.0;
/// Cut a table line into cells at a wider gap (label rule in tables).
/// (The text-line `SPLIT_GAP = 3.0` split of pipeline.py assemble is not
/// ported yet — merged-column text lines stay whole.)
const CELL_GAP: f32 = 1.8;

// BLOCK_CLASSES — the OrderNet layout heads' class ids (order_post.py).
pub const BLOCK_CLASSES: [&str; 10] = [
    "Text",
    "Section-Header",
    "Page-Header",
    "Page-Footer",
    "List-Group",
    "Table",
    "Caption",
    "Footnote",
    "Table-Of-Contents",
    "Form",
];
const GRID_LABELS: [&str; 3] = ["Table", "Form", "Table-Of-Contents"];

// --- model dir resolution ------------------------------------------------

fn resolve_model_dir() -> Result<PathBuf, String> {
    if let Some(dir) = std::env::var_os("INDRA_OCR_MODEL_DIR") {
        let p = PathBuf::from(dir);
        if needed_files(&p).is_empty() {
            return Ok(p);
        }
        return Err(format!(
            "INDRA_OCR_MODEL_DIR={} lacks det.onnx/rec.onnx/order.onnx/meta.json",
            p.display()
        ));
    }
    // Scan the standard Hugging Face cache for the indra-ocr snapshot.
    let mut candidates: Vec<PathBuf> = Vec::new();
    if let Some(home) = std::env::var_os("HOME") {
        let hub = Path::new(&home).join(".cache/huggingface/hub/models--datalab-to--indra-ocr/snapshots");
        if let Ok(rd) = std::fs::read_dir(&hub) {
            for e in rd.flatten() {
                let p = e.path();
                if needed_files(&p).is_empty() {
                    candidates.push(p);
                }
            }
        }
    }
    candidates.sort_by_key(|p| {
        std::fs::metadata(p)
            .and_then(|m| m.modified())
            .ok()
            .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
            .map(|d| d.as_secs())
            .unwrap_or(0)
    });
    candidates.pop().ok_or_else(|| {
        "no Indra models found: set INDRA_OCR_MODEL_DIR to a folder with \
         det.onnx/rec.onnx/order.onnx/meta.json (download via \
         `pip install indra-ocr` or huggingface_hub snapshot_download)"
            .to_string()
    })
}

fn needed_files(dir: &Path) -> Vec<&'static str> {
    ["det.onnx", "rec.onnx", "order.onnx", "meta.json"]
        .into_iter()
        .filter(|f| !dir.join(f).exists())
        .collect()
}

// --- meta.json -----------------------------------------------------------

#[derive(Deserialize)]
struct Meta {
    #[serde(default = "default_long_side")]
    det_long_side: i64,
    #[serde(default = "default_rec_height")]
    rec_height: i64,
    /// The vocab is a JSON-encoded string inside meta.json (double-encoded).
    #[serde(default)]
    vocab: Option<String>,
}

fn default_long_side() -> i64 {
    DET_LONG_SIDE_DEFAULT
}
fn default_rec_height() -> i64 {
    REC_HEIGHT_DEFAULT
}

#[derive(Deserialize)]
struct VocabJson {
    chars: Vec<String>,
}

// --- grayscale page ------------------------------------------------------

pub struct GrayPage {
    pub w: usize,
    pub h: usize,
    /// Luma bytes, row-major.
    pub data: Vec<u8>,
}

impl GrayPage {
    pub fn from_png(bytes: &[u8]) -> Result<Self, String> {
        let img = image::load_from_memory(bytes).map_err(|e| format!("png: {e}"))?;
        let luma = img.to_luma8();
        Ok(Self {
            w: luma.width() as usize,
            h: luma.height() as usize,
            data: luma.into_raw(),
        })
    }

    fn resize(&self, nw: usize, nh: usize) -> GrayPage {
        let img =
            image::GrayImage::from_raw(self.w as u32, self.h as u32, self.data.clone())
                .expect("dims match");
        // INTER_AREA ≈ triangle filtering for downscale; INTER_LINEAR ≈ catmull-rom
        // is close enough for upscale. Text-equality is the test, not byte equality.
        let filter = if (nw as f32) < (self.w as f32) {
            image::imageops::FilterType::Triangle
        } else {
            image::imageops::FilterType::CatmullRom
        };
        let out = image::imageops::resize(&img, nw as u32, nh as u32, filter);
        GrayPage {
            w: nw,
            h: nh,
            data: out.into_raw(),
        }
    }
}

// --- detection -----------------------------------------------------------

/// Resize the page so its long side is `long_side`, pad to /32 with white.
/// Returns (ink map [ph*pw], ph, pw, scale).
fn page_to_input(gray: &GrayPage, long_side: i64) -> (Vec<f32>, usize, usize, f32) {
    let scale = long_side as f32 / (gray.w.max(gray.h) as f32);
    let nw = ((gray.w as f32 * scale).round() as usize).max(1);
    let nh = ((gray.h as f32 * scale).round() as usize).max(1);
    let small = if (nw, nh) != (gray.w, gray.h) {
        gray.resize(nw, nh)
    } else {
        GrayPage {
            w: gray.w,
            h: gray.h,
            data: gray.data.clone(),
        }
    };
    let (nw, nh) = (small.w, small.h);
    let ph = (nh as f32 / INPUT_MULTIPLE).ceil() as usize * INPUT_MULTIPLE as usize;
    let pw = (nw as f32 / INPUT_MULTIPLE).ceil() as usize * INPUT_MULTIPLE as usize;
    let mut out = vec![0.0f32; ph * pw]; // ink map: 0 = white
    for r in 0..nh {
        for c in 0..nw {
            out[r * pw + c] = 1.0 - small.data[r * nw + c] as f32 / 255.0;
        }
    }
    (out, ph, pw, scale)
}

/// 4-connected components over a binary mask.
struct Components {
    comps: Vec<Comp>,
}
struct Comp {
    area: i32,
    min_col: i32,
    max_col: i32,
    sum_top: f64,
    sum_bot: f64,
    sum_prob: f64,
}

fn connected_components(
    mask: &[bool],
    w: usize,
    h: usize,
    core: &[f32],
    top: &[f32],
    bot: &[f32],
) -> Components {
    let mut labels = vec![0i32; w * h];
    let mut comps: Vec<Comp> = Vec::new();
    let mut stack: Vec<usize> = Vec::new();
    for start in 0..w * h {
        if !mask[start] || labels[start] != 0 {
            continue;
        }
        let id = comps.len() as i32 + 1;
        let mut comp = Comp {
            area: 0,
            min_col: i32::MAX,
            max_col: 0,
            sum_top: 0.0,
            sum_bot: 0.0,
            sum_prob: 0.0,
        };
        stack.push(start);
        labels[start] = id;
        while let Some(p) = stack.pop() {
            let (r, c) = (p / w, p % w);
            comp.area += 1;
            comp.min_col = comp.min_col.min(c as i32);
            comp.max_col = comp.max_col.max(c as i32 + 1);
            let ys = (r as f32 + 0.5) * OUT_STRIDE;
            comp.sum_top += (ys - top[p]) as f64;
            comp.sum_bot += (ys + bot[p]) as f64;
            comp.sum_prob += core[p] as f64;
            if c > 0 && mask[p - 1] && labels[p - 1] == 0 {
                labels[p - 1] = id;
                stack.push(p - 1);
            }
            if c + 1 < w && mask[p + 1] && labels[p + 1] == 0 {
                labels[p + 1] = id;
                stack.push(p + 1);
            }
            if r > 0 && mask[p - w] && labels[p - w] == 0 {
                labels[p - w] = id;
                stack.push(p - w);
            }
            if r + 1 < h && mask[p + w] && labels[p + w] == 0 {
                labels[p + w] = id;
                stack.push(p + w);
            }
        }
        comps.push(comp);
    }
    Components { comps }
}

/// det_post.decode — out is [3, h, w] (core logit, top dist, bottom dist).
/// Returns boxes in input pixels: (x0, y0, x1, y1, score).
fn det_decode(out: &[f32], h: usize, w: usize, thresh: f32, min_area: i32) -> Vec<(f32, f32, f32, f32, f32)> {
    let n = h * w;
    let core = &out[0..n];
    let top = &out[n..2 * n];
    let bot = &out[2 * n..3 * n];
    let mask: Vec<bool> = core.iter().map(|&l| l > 0.0 && sigmoid(l) > thresh).collect();
    let comps = connected_components(&mask, w, h, core, top, bot);
    let mut boxes = Vec::new();
    for c in &comps.comps {
        if c.area < min_area {
            continue;
        }
        let cnt = c.area as f64;
        let y0 = (c.sum_top / cnt) as f32;
        let y1 = (c.sum_bot / cnt) as f32;
        let score = (c.sum_prob / cnt) as f32;
        let lh = (y1 - y0).max(1.0);
        let ex = (CORE_X_SHRINK * lh).min(0.5 * lh);
        let x0 = c.min_col as f32 * OUT_STRIDE - ex;
        let x1 = c.max_col as f32 * OUT_STRIDE + ex;
        boxes.push((x0, y0, x1, y1, score));
    }
    boxes
}

fn sigmoid(x: f32) -> f32 {
    1.0 / (1.0 + (-x).exp())
}

fn median(vals: &mut [f32]) -> f32 {
    vals.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    if vals.is_empty() {
        0.0
    } else {
        vals[vals.len() / 2]
    }
}

// --- recognition ---------------------------------------------------------

/// rec_preproc.crop_line — expand the box, crop (white outside the page),
/// resize to `height` px keeping aspect. Returns (width, ink map).
fn crop_line(gray: &GrayPage, bx: (f32, f32, f32, f32), height: usize) -> (usize, Vec<f32>) {
    let (x0, y0, x1, y1) = bx;
    let h = (y1 - y0).max(1.0);
    let gx0 = x0 - PAD_X * h;
    let gx1 = x1 + PAD_X * h;
    let gy0 = y0 - PAD_Y * h;
    let gy1 = y1 + PAD_Y * h;
    let ix0 = gx0.floor() as i64;
    let iy0 = gy0.floor() as i64;
    let ix1 = gx1.ceil() as i64;
    let iy1 = gy1.ceil() as i64;
    let cw = (ix1 - ix0).max(1) as usize;
    let ch = (iy1 - iy0).max(1) as usize;
    // crop, white (255) outside the page
    let mut crop = vec![255u8; cw * ch];
    let sx0 = ix0.max(0);
    let sy0 = iy0.max(0);
    let sx1 = (ix1 as i64).min(gray.w as i64);
    let sy1 = (iy1 as i64).min(gray.h as i64);
    if sx1 > sx0 && sy1 > sy0 {
        for r in sy0..sy1 {
            let dst = ((r - iy0) as usize) * cw + ((sx0 - ix0) as usize);
            let src = (r as usize) * gray.w + (sx0 as usize);
            let len = (sx1 - sx0) as usize;
            crop[dst..dst + len].copy_from_slice(&gray.data[src..src + len]);
        }
    }
    let out_w = (((cw as f32) * (height as f32) / (ch as f32)).round() as usize)
        .clamp(WIDTH_MULTIPLE, REC_MAX_WIDTH);
    let img = image::GrayImage::from_raw(cw as u32, ch as u32, crop).expect("dims match");
    let filter = if ch > height {
        image::imageops::FilterType::Triangle
    } else {
        image::imageops::FilterType::CatmullRom
    };
    let resized = image::imageops::resize(&img, out_w as u32, height as u32, filter);
    let ink = resized
        .into_raw()
        .into_iter()
        .map(|v| 1.0 - v as f32 / 255.0)
        .collect();
    (out_w, ink)
}

/// rec_preproc.pad_batch — right-pad ink maps to one width (multiple of 8).
/// Returns ([B,1,H,W] data, widths).
fn pad_batch(crops: &[(usize, Vec<f32>)]) -> (Vec<f32>, Vec<i64>, usize) {
    let b = crops.len();
    let wm = crops
        .iter()
        .map(|(w, _)| *w)
        .max()
        .unwrap_or(WIDTH_MULTIPLE);
    let wm = wm.div_ceil(WIDTH_MULTIPLE) * WIDTH_MULTIPLE;
    let mut data = vec![0.0f32; b * LINE_HEIGHT * wm];
    for (k, (w, ink)) in crops.iter().enumerate() {
        for r in 0..LINE_HEIGHT {
            let dst = k * LINE_HEIGHT * wm + r * wm;
            let src = r * w;
            data[dst..dst + w].copy_from_slice(&ink[src..src + w]);
        }
    }
    let widths = crops.iter().map(|(w, _)| *w as i64).collect();
    (data, widths, wm)
}

/// rec_preproc.width_batches — group line indices into batches that keep
/// batch width bounded. Narrowest first; a batch closes at MAX_BATCH lines,
/// when the next line is wider than MAX_RATIO x the batch's first line
/// (widths below MIN_WIDTH count as MIN_WIDTH), or when lines x width would
/// pass MAX_AREA pixel columns.
fn width_batches(crops: &[(usize, Vec<f32>)]) -> Vec<Vec<usize>> {
    const MAX_BATCH: usize = 48;
    const MIN_WIDTH: usize = 128;
    const MAX_RATIO: f32 = 8.0;
    const MAX_AREA: usize = 512 * 512;
    let mut order: Vec<usize> = (0..crops.len()).collect();
    order.sort_by_key(|&i| crops[i].0);
    let mut batches: Vec<Vec<usize>> = Vec::new();
    for &i in &order {
        let w = crops[i].0.max(MIN_WIDTH);
        match batches.last_mut() {
            Some(b)
                if b.len() < MAX_BATCH
                    && (w as f32) <= crops[b[0]].0.max(MIN_WIDTH) as f32 * MAX_RATIO
                    && b.len() * w <= MAX_AREA =>
            {
                b.push(i)
            }
            _ => batches.push(vec![i]),
        }
    }
    batches
}

// --- reading order (order_post.py) ---------------------------------------

const N_FEATS: usize = 16;

fn line_features(boxes: &[(f32, f32, f32, f32)], wh: (f32, f32), texts: &[String]) -> Vec<f32> {
    let (w, h) = (wh.0.max(1.0), wh.1.max(1.0));
    let n = boxes.len();
    let mut feats = vec![0.0f32; n * N_FEATS];
    if n == 0 {
        return feats;
    }
    let mut heights: Vec<f32> = boxes.iter().map(|b| (b.3 - b.1).max(1e-3)).collect();
    let med_h = median(&mut heights).max(1e-3);
    for (i, b) in boxes.iter().enumerate() {
        let bw = (b.2 - b.0).max(1e-3);
        let bh = (b.3 - b.1).max(1e-3);
        let g = [
            b.0 / w,
            b.1 / h,
            b.2 / w,
            b.3 / h,
            (b.0 + b.2) / (2.0 * w),
            (b.1 + b.3) / (2.0 * h),
            bw / w,
            (bh / med_h).ln(),
            (bw / bh).ln() / 4.0,
            med_h / h * 20.0,
        ];
        feats[i * N_FEATS..i * N_FEATS + 10].copy_from_slice(&g);
        let t = texts.get(i).map(|s| s.trim()).unwrap_or("");
        if t.is_empty() {
            continue;
        }
        let digits = t.chars().filter(|c| c.is_ascii_digit()).count() as f32 / t.len() as f32;
        // t[:3] in Python is CHAR-based; slicing bytes panics on multibyte text
        let first3: String = t.chars().take(3).collect();
        let first3_num = first3
            .trim_end_matches(['.', ')'])
            .chars()
            .all(|c| c.is_ascii_digit())
            && first3.ends_with(|c: char| c.is_ascii_digit() || c == ')' || c == '.');
        let starts_item = "•-–▪◦*·→".contains(t.chars().next().unwrap_or(' '))
            || (t.chars().count() > 3 && first3_num);
        feats[i * N_FEATS + 10] = (t.len() as f32).ln_1p() / 5.0;
        feats[i * N_FEATS + 11] = digits;
        feats[i * N_FEATS + 12] = if starts_item { 1.0 } else { 0.0 };
        feats[i * N_FEATS + 13] = if t.ends_with('-') { 1.0 } else { 0.0 };
        feats[i * N_FEATS + 14] = if t.ends_with(['.', '!', '?', ':', ';']) {
            1.0
        } else {
            0.0
        };
        feats[i * N_FEATS + 15] = if t.chars().next().is_some_and(|c| c.is_uppercase()) {
            1.0
        } else {
            0.0
        };
    }
    feats
}

/// -ln(1 + exp(-a)) — log P(i before j) from the pairwise logit.
fn log_p_before(a: f64) -> f64 {
    if a > 0.0 {
        (-(-a).exp()).ln_1p()
    } else {
        -(-a).exp_m1().ln_1p() - a
    }
}

/// decode_order — row-sum ranking + insertion-move local search.
fn decode_order(a: &[f64], n: usize, sweeps: usize) -> Vec<usize> {
    if n <= 1 {
        return (0..n).collect();
    }
    let mut lp = vec![0.0f64; n * n];
    for i in 0..n {
        for j in 0..n {
            lp[i * n + j] = log_p_before(a[i * n + j]);
        }
    }
    let mut rowsum = vec![0.0f64; n];
    let mut colsum = vec![0.0f64; n];
    for i in 0..n {
        for j in 0..n {
            rowsum[i] += lp[i * n + j];
            colsum[j] += lp[i * n + j];
        }
    }
    let mut idx: Vec<usize> = (0..n).collect();
    idx.sort_by(|&x, &y| {
        (colsum[x] - rowsum[x])
            .partial_cmp(&(colsum[y] - rowsum[y]))
            .unwrap_or(std::cmp::Ordering::Equal)
    });
    for _ in 0..sweeps {
        let mut improved = false;
        for &x in idx.clone().iter() {
            let rest: Vec<usize> = idx.iter().copied().filter(|&i| i != x).collect();
            let m = rest.len();
            // before[j] = lp[rest[j], x]; after[j] = lp[x, rest[j]]
            let before: Vec<f64> = rest.iter().map(|&r| lp[r * n + x]).collect();
            let after: Vec<f64> = rest.iter().map(|&r| lp[x * n + r]).collect();
            // gains[k] = Σ_{j<k} before[j] + Σ_{j≥k} after[j], k in 0..=m
            let mut gains = vec![0.0f64; m + 1];
            for k in 0..m {
                gains[k + 1] = gains[k] + before[k];
            }
            let mut suf = 0.0f64;
            for k in (0..m).rev() {
                // position k sees after[k..]
                suf += after[k];
                gains[k] += suf;
            }
            let cur = idx.iter().position(|&i| i == x).unwrap_or(0);
            let (mut bestk, mut bestg) = (cur, gains[cur.min(m)]);
            for (k, &g) in gains.iter().enumerate().take(m + 1) {
                if g > bestg + 1e-6 {
                    (bestk, bestg) = (k, g);
                }
            }
            if bestk != cur && bestg > gains[cur.min(m)] + 1e-6 {
                idx.remove(cur);
                idx.insert(bestk, x);
                improved = true;
            }
        }
        if !improved {
            break;
        }
    }
    idx
}

/// order_top_down — rows top to bottom, left to right within a row.
fn order_top_down(boxes: &[(f32, f32, f32, f32)]) -> Vec<usize> {
    let mut heights: Vec<f32> = boxes.iter().map(|b| b.3 - b.1).collect();
    let h = median(&mut heights).max(1e-3);
    let mut idx: Vec<usize> = (0..boxes.len()).collect();
    idx.sort_by(|&a, &b| {
        let (ca, cb) = (
            ((boxes[a].1 + boxes[a].3) / 2.0),
            ((boxes[b].1 + boxes[b].3) / 2.0),
        );
        let ka = (ca / (h * 0.5)).round();
        let kb = (cb / (h * 0.5)).round();
        ka.partial_cmp(&kb)
            .unwrap_or(std::cmp::Ordering::Equal)
            .then(boxes[a].0.partial_cmp(&boxes[b].0).unwrap_or(std::cmp::Ordering::Equal))
    });
    idx
}

/// order_xycut — recursive split at the widest whitespace gap on either axis.
fn order_xycut(boxes: &[(f32, f32, f32, f32)]) -> Vec<usize> {
    fn best_cut(boxes: &[(f32, f32, f32, f32)], axis: usize) -> (f32, usize, Vec<usize>) {
        // returns (best_gap, best_k, order_by_lo)
        let (lo, hi): (Vec<f32>, Vec<f32>) = boxes
            .iter()
            .map(|b| if axis == 1 { (b.1, b.3) } else { (b.0, b.2) })
            .unzip();
        let mut o: Vec<usize> = (0..boxes.len()).collect();
        o.sort_by(|&a, &b| lo[a].partial_cmp(&lo[b]).unwrap_or(std::cmp::Ordering::Equal));
        let mut end = hi[o[0]];
        let (mut best_gap, mut best_k) = (0.0f32, usize::MAX);
        for k in 1..o.len() {
            let gap = lo[o[k]] - end;
            if gap > best_gap {
                best_gap = gap;
                best_k = k;
            }
            end = end.max(hi[o[k]]);
        }
        (best_gap, best_k, o)
    }
    fn rec(boxes: &[(f32, f32, f32, f32)], ids: Vec<usize>, out: &mut Vec<usize>) {
        if ids.len() <= 1 {
            out.extend(ids);
            return;
        }
        let sub: Vec<(f32, f32, f32, f32)> = ids.iter().map(|&i| boxes[i]).collect();
        let (gy, ky, oy) = best_cut(&sub, 1);
        let (gx, kx, ox) = best_cut(&sub, 0);
        if ky == usize::MAX && kx == usize::MAX {
            for i in order_top_down(&sub) {
                out.push(ids[i]);
            }
            return;
        }
        let (k, o) = if gy >= gx { (ky, oy) } else { (kx, ox) };
        let left: Vec<usize> = o[..k].iter().map(|&i| ids[i]).collect();
        let right: Vec<usize> = o[k..].iter().map(|&i| ids[i]).collect();
        rec(boxes, left, out);
        rec(boxes, right, out);
    }
    let mut out = Vec::with_capacity(boxes.len());
    rec(boxes, (0..boxes.len()).collect(), &mut out);
    out
}

/// order_post._cluster — average-linkage clustering of pairwise same-block
/// logits; merge while the mean logit between groups is positive.
fn cluster_same(same: &[f32], n: usize) -> Vec<Vec<usize>> {
    if n == 0 {
        return Vec::new();
    }
    let mut m = vec![0.0f64; n * n]; // mean logit between groups; -inf dead
    for i in 0..n {
        for j in 0..n {
            m[i * n + j] = if i == j {
                f64::NEG_INFINITY
            } else {
                same[i * n + j] as f64
            };
        }
    }
    let mut size = vec![1.0f64; n];
    let mut alive = vec![true; n];
    let mut members: Vec<Vec<usize>> = (0..n).map(|i| vec![i]).collect();
    let mut best: Vec<usize> = (0..n)
        .map(|i| {
            (0..n)
                .max_by(|&a, &b| {
                    m[i * n + a]
                        .partial_cmp(&m[i * n + b])
                        .unwrap_or(std::cmp::Ordering::Equal)
                })
                .unwrap_or(i)
        })
        .collect();
    let mut best_v: Vec<f64> = (0..n).map(|i| m[i * n + best[i]]).collect();
    for _ in 0..n.saturating_sub(1) {
        let a = (0..n)
            .filter(|&i| alive[i])
            .max_by(|&x, &y| {
                best_v[x]
                    .partial_cmp(&best_v[y])
                    .unwrap_or(std::cmp::Ordering::Equal)
            });
        let Some(a) = a else { break };
        if !alive[a] || best_v[a] <= 0.0 {
            break;
        }
        let b = best[a];
        let (a, b) = (a.min(b), a.max(b));
        // merge b into a
        for j in 0..n {
            m[a * n + j] += m[b * n + j];
            m[j * n + a] += m[j * n + b];
        }
        m[a * n + a] = 0.0;
        size[a] += size[b];
        alive[b] = false;
        let taken = std::mem::take(&mut members[b]);
        members[a].extend(taken);
        for j in 0..n {
            m[b * n + j] = f64::NEG_INFINITY;
            m[j * n + b] = f64::NEG_INFINITY;
        }
        // recompute row/col a means
        let mut row = vec![f64::NEG_INFINITY; n];
        for j in 0..n {
            if j == a || !alive[j] {
                continue;
            }
            row[j] = m[a * n + j] / (size[a] * size[j]);
            m[j * n + a] = row[j];
        }
        m[a * n + a] = f64::NEG_INFINITY;
        m[a * n..a * n + n].copy_from_slice(&row);
        m[a * n + a] = f64::NEG_INFINITY;
        best_v[b] = f64::NEG_INFINITY;
        // refresh best partners affected by the merge
        for j in 0..n {
            if !alive[j] {
                continue;
            }
            if j == a {
                let bj = (0..n)
                    .filter(|&k| alive[k] && k != a)
                    .max_by(|&x, &y| {
                        m[j * n + x]
                            .partial_cmp(&m[j * n + y])
                            .unwrap_or(std::cmp::Ordering::Equal)
                    });
                best[j] = bj.unwrap_or(j);
                best_v[j] = if bj.is_some() { m[j * n + best[j]] } else { f64::NEG_INFINITY };
            } else if best[j] == a || best[j] == b {
                let bj = (0..n)
                    .filter(|&k| alive[k] && k != j)
                    .max_by(|&x, &y| {
                        m[j * n + x]
                            .partial_cmp(&m[j * n + y])
                            .unwrap_or(std::cmp::Ordering::Equal)
                    });
                best[j] = bj.unwrap_or(j);
                best_v[j] = if bj.is_some() { m[j * n + best[j]] } else { f64::NEG_INFINITY };
            }
        }
    }
    (0..n)
        .filter(|&i| alive[i])
        .map(|i| {
            let mut g = members[i].clone();
            g.sort_unstable();
            g
        })
        .collect()
}

/// decode_blocks_order — blocks first (clustering), then lines inside a block
/// by geometry, blocks by mean pairwise order logit. Returns reading order.
fn decode_blocks_order(
    a: &[f32],
    same: &[f32],
    boxes: &[(f32, f32, f32, f32)],
    cls: Option<&[f32]>,
    n: usize,
) -> Vec<usize> {
    if n <= 1 {
        return (0..n).collect();
    }
    let blocks = cluster_same(same, n);
    let grid: Vec<usize> = GRID_LABELS
        .iter()
        .filter_map(|c| BLOCK_CLASSES.iter().position(|b| b == c))
        .collect();
    let mut inner: Vec<Vec<usize>> = Vec::with_capacity(blocks.len());
    for b in &blocks {
        let bx: Vec<(f32, f32, f32, f32)> = b.iter().map(|&i| boxes[i]).collect();
        let is_grid = cls
            .map(|c| {
                let mut sum = [0.0f32; 10];
                for &i in b {
                    for k in 0..10 {
                        sum[k] += c[i * 10 + k];
                    }
                }
                let arg = (0..10)
                    .max_by(|&x, &y| sum[x].partial_cmp(&sum[y]).unwrap_or(std::cmp::Ordering::Equal))
                    .unwrap_or(0);
                grid.contains(&arg)
            })
            .unwrap_or(false);
        let local = if is_grid {
            order_top_down(&bx)
        } else {
            order_xycut(&bx)
        };
        inner.push(local.iter().map(|&k| b[k]).collect());
    }
    // Mean order logit between blocks, antisymmetrized: G·a·Gᵀ.
    let m = blocks.len();
    let mut ab = vec![0.0f64; m * m];
    for (p, bp) in blocks.iter().enumerate() {
        for (q, bq) in blocks.iter().enumerate() {
            if p == q {
                continue;
            }
            let mut s = 0.0f64;
            for &i in bp {
                for &j in bq {
                    s += a[i * n + j] as f64;
                }
            }
            let denom = (bp.len() * bq.len()) as f64;
            ab[p * m + q] = s / denom;
        }
    }
    let mut antisym = vec![0.0f64; m * m];
    for p in 0..m {
        for q in 0..m {
            antisym[p * m + q] = ab[p * m + q] - ab[q * m + p];
        }
    }
    let block_order = decode_order(&antisym, m, 4);
    let mut out = Vec::with_capacity(n);
    for k in block_order {
        out.extend(&inner[k]);
    }
    out
}

// --- line splitting + table grid (pipeline.py split_line / layout_post) ---

fn is_space(tok: &str) -> bool {
    tok == " "
}

fn decode_toks<'a>(toks: impl IntoIterator<Item = &'a (String, f32)>) -> String {
    let joined: String = toks.into_iter().map(|(s, _)| s.as_str()).collect();
    joined.split_whitespace().collect::<Vec<_>>().join(" ")
}

/// pipeline.py split_line — cut a line where neighbouring tokens are more
/// than `gap` line-heights apart (the detector merged two cells or columns).
/// Returns (sub-box, text) segments; one item when nothing is cut.
fn split_line(
    box_: &(f32, f32, f32, f32),
    toks: &[(String, f32)],
    gap: f32,
) -> Vec<((f32, f32, f32, f32), String)> {
    let (x0, y0, x1, y1) = *box_;
    let h = (y1 - y0).max(1.0);
    let mut segs: Vec<Vec<&(String, f32)>> = vec![Vec::new()];
    let mut prev: Option<f32> = None;
    for tk in toks {
        let x = tk.1;
        if !is_space(&tk.0) {
            let cut = prev.is_some_and(|p| x - p > gap * h);
            if cut {
                segs.push(Vec::new());
            }
            prev = Some(x);
        }
        segs.last_mut().expect("segs never empty").push(tk);
    }
    if segs.len() == 1 {
        return vec![(*box_, decode_toks(toks))];
    }
    let mut out = Vec::new();
    for sg in segs {
        let text = decode_toks(sg.iter().copied());
        let xs: Vec<f32> = sg.iter().filter(|tk| !is_space(&tk.0)).map(|tk| tk.1).collect();
        if !text.is_empty() && !xs.is_empty() {
            let min_x = xs.iter().cloned().fold(f32::MAX, f32::min);
            let max_x = xs.iter().cloned().fold(f32::MIN, f32::max);
            let bx = (
                x0.max(min_x - 0.35 * h),
                y0,
                x1.min(max_x + 0.35 * h),
                y1,
            );
            out.push((bx, text));
        }
    }
    out
}

/// layout_post _rows + column clustering: group cells into rows by vertical
/// centre, then assign columns by clustered x positions. Returns rows × cols
/// of cell texts, or None when nothing parses.
fn build_table_grid(cells: &[(f32, f32, f32, f32, String)]) -> Option<Vec<Vec<String>>> {
    if cells.is_empty() {
        return None;
    }
    // rows: sort by centre y, merge within half the median height
    let mut heights: Vec<f32> = cells.iter().map(|c| c.3 - c.1).collect();
    let h = {
        heights.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
        *heights.get(heights.len() / 2).unwrap_or(&1.0)
    }
    .max(1.0);
    let mut order: Vec<usize> = (0..cells.len()).collect();
    order.sort_by(|&a, &b| {
        let ca = (cells[a].1 + cells[a].3) / 2.0;
        let cb = (cells[b].1 + cells[b].3) / 2.0;
        ca.partial_cmp(&cb).unwrap_or(std::cmp::Ordering::Equal)
    });
    let mut rows: Vec<Vec<usize>> = Vec::new();
    for &i in &order {
        let cy = (cells[i].1 + cells[i].3) / 2.0;
        if let Some(last) = rows.last_mut() {
            let last_cy = last.iter().map(|&j| (cells[j].1 + cells[j].3) / 2.0).sum::<f32>()
                / last.len() as f32;
            if (cy - last_cy).abs() < 0.5 * h {
                last.push(i);
                continue;
            }
        }
        rows.push(vec![i]);
    }
    for r in &mut rows {
        r.sort_by(|&a, &b| {
            cells[a].0
                .partial_cmp(&cells[b].0)
                .unwrap_or(std::cmp::Ordering::Equal)
        });
    }
    // columns: cluster all cell left edges with a gap threshold
    let mut xs: Vec<f32> = cells.iter().map(|c| c.0).collect();
    xs.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let mut col_centers: Vec<f32> = Vec::new();
    for &x in &xs {
        match col_centers.last() {
            Some(&c) if x - c <= 1.2 * h => continue, // same column band
            _ => col_centers.push(x),
        }
    }
    let col_of = |x: f32| -> usize {
        col_centers
            .iter()
            .enumerate()
            .min_by(|(_, a), (_, b)| {
                (**a - x)
                    .abs()
                    .partial_cmp(&((**b - x).abs()))
                    .unwrap_or(std::cmp::Ordering::Equal)
            })
            .map(|(i, _)| i)
            .unwrap_or(0)
    };
    let ncols = col_centers.len();
    let mut grid: Vec<Vec<String>> = rows
        .iter()
        .map(|_| vec![String::new(); ncols])
        .collect();
    for (ri, r) in rows.iter().enumerate() {
        for &i in r {
            let col = col_of(cells[i].0).min(ncols - 1);
            let text = cells[i].4.trim().to_string();
            if grid[ri][col].is_empty() {
                grid[ri][col] = text;
            } else {
                grid[ri][col].push(' ');
                grid[ri][col].push_str(&text);
            }
        }
    }
    // drop empty trailing columns
    let mut keep = ncols;
    while keep > 1 && grid.iter().all(|r| r[keep - 1].is_empty()) {
        keep -= 1;
    }
    let mut out: Vec<Vec<String>> = grid
        .into_iter()
        .map(|mut r| {
            r.truncate(keep);
            r
        })
        .collect();
    // drop fully-empty rows
    out.retain(|r| r.iter().any(|c| !c.is_empty()));
    (!out.is_empty()).then_some(out)
}

// --- engine --------------------------------------------------------------

pub struct IndraNative {
    det: ort::session::Session,
    rec: ort::session::Session,
    order: Option<ort::session::Session>,
    vocab: Vec<String>,
    det_long_side: i64,
    rec_height: usize,
    min_page_confidence: f64,
    last_page_confidence: Option<f64>,
}

pub struct NativeLine {
    pub box_: (f32, f32, f32, f32),
    pub text: String,
    pub conf: f32,
}

/// One recognized line with per-token page-x positions (for cell splitting).
pub struct LineRec {
    pub box_: (f32, f32, f32, f32),
    pub text: String,
    pub conf: f32,
    /// (token text, page x) for every kept CTC token, in reading order.
    pub toks: Vec<(String, f32)>,
}

pub struct NativeBlock {
    pub label: String,
    pub bbox: (f32, f32, f32, f32),
    pub text: String,
    /// Reconstructed cells for Table-class blocks (rows × columns).
    pub table: Option<Vec<Vec<String>>>,
}

pub struct NativePage {
    pub lines: Vec<NativeLine>,
    pub blocks: Vec<NativeBlock>,
}

fn session(path: &Path) -> Result<ort::session::Session, String> {
    let threads: usize = std::env::var("ONNX_THREADS")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(4);
    ort::session::Session::builder()
        .map_err(|e| format!("session builder: {e:?}"))?
        .with_intra_threads(threads)
        .map_err(|e| format!("intra threads: {e:?}"))?
        .with_memory_pattern(false)
        .map_err(|e| format!("memory pattern: {e:?}"))?
        .commit_from_file(path)
        .map_err(|e| format!("{}: {e:?}", path.display()))
}

impl IndraNative {
    pub fn load() -> Result<Self, String> {
        let dir = resolve_model_dir()?;
        let meta: Meta = serde_json::from_str(
            &std::fs::read_to_string(dir.join("meta.json"))
                .map_err(|e| format!("meta.json: {e}"))?,
        )
        .map_err(|e| format!("meta.json parse: {e}"))?;
        let vocab: Vec<String> = serde_json::from_str::<VocabJson>(
            meta.vocab.as_deref().ok_or("meta.json lacks vocab")?,
        )
        .map_err(|e| format!("vocab parse: {e}"))?
        .chars;
        Ok(Self {
            det: session(&dir.join("det.onnx"))?,
            rec: session(&dir.join("rec.onnx"))?,
            order: session(&dir.join("order.onnx")).ok(),
            vocab,
            det_long_side: meta.det_long_side,
            rec_height: meta.rec_height as usize,
            min_page_confidence: super::pdf_ocr::DEFAULT_MIN_PAGE_CONFIDENCE,
            last_page_confidence: None,
        })
    }

    fn detect(&mut self, gray: &GrayPage) -> Result<Vec<(f32, f32, f32, f32, f32)>, String> {
        let mut long_side = self.det_long_side;
        let mut boxes = {
            let (input, ph, pw, _s) = page_to_input(gray, long_side);
            let x = ort::value::Tensor::from_array(([1i64, 1, ph as i64, pw as i64], input))
                .map_err(|e| format!("det input: {e:?}"))?;
            let outputs = self
                .det
                .run(ort::inputs!["x" => x])
                .map_err(|e| format!("det run: {e:?}"))?;
            let (_, out) = outputs["out"]
                .try_extract_tensor::<f32>()
                .map_err(|e| format!("det out: {e:?}"))?;
            let oh = (ph / 2).max(1);
            let ow = (pw / 2).max(1);
            det_decode(&out[..oh * ow * 3], oh, ow, 0.5, 2)
        };
        // adaptive second pass for very small text
        if boxes.len() >= 8 {
            let mut heights: Vec<f32> = boxes.iter().map(|b| b.3 - b.1).collect();
            let med_h = median(&mut heights);
            if med_h < ADAPTIVE_MIN_H {
                let ls = ((long_side as f32 * ADAPTIVE_TARGET_H / med_h.max(1.0))
                    .min(ADAPTIVE_MAX_SIDE) as i64)
                    / 32
                    * 32;
                if ls > long_side {
                    long_side = ls;
                    let (input, ph, pw, s2) = page_to_input(gray, long_side);
                    let x = ort::value::Tensor::from_array(([1i64, 1, ph as i64, pw as i64], input))
                        .map_err(|e| format!("det input: {e:?}"))?;
                    let outputs = self
                        .det
                        .run(ort::inputs!["x" => x])
                        .map_err(|e| format!("det run: {e:?}"))?;
                    let (_, out) = outputs["out"]
                        .try_extract_tensor::<f32>()
                        .map_err(|e| format!("det out: {e:?}"))?;
                    let oh = (ph / 2).max(1);
                    let ow = (pw / 2).max(1);
                    boxes = det_decode(&out[..oh * ow * 3], oh, ow, 0.5, 2);
                    let _ = s2;
                }
            }
        }
        // back to page pixels, clipped
        let s = long_side as f32 / (gray.w.max(gray.h) as f32);
        for b in &mut boxes {
            b.0 /= s;
            b.1 /= s;
            b.2 /= s;
            b.3 /= s;
            b.0 = b.0.clamp(0.0, gray.w as f32);
            b.2 = b.2.clamp(0.0, gray.w as f32);
            b.1 = b.1.clamp(0.0, gray.h as f32);
            b.3 = b.3.clamp(0.0, gray.h as f32);
        }
        Ok(boxes)
    }

    fn recognize(
        &mut self,
        gray: &GrayPage,
        boxes: &[(f32, f32, f32, f32, f32)],
    ) -> Result<Vec<LineRec>, String> {
        let crops: Vec<(usize, Vec<f32>)> = boxes
            .iter()
            .map(|b| crop_line(gray, (b.0, b.1, b.2, b.3), self.rec_height))
            .collect();
        let mut out: Vec<LineRec> = boxes
            .iter()
            .map(|b| LineRec {
                box_: (b.0, b.1, b.2, b.3),
                text: String::new(),
                conf: 0.0,
                toks: Vec::new(),
            })
            .collect();
        if crops.is_empty() {
            return Ok(out);
        }
        // width-aware batching (rec_preproc.width_batches port): never let
        // lines x width blow the arena — Python measured 64 x 4096px = 5 GB.
        for batch in width_batches(&crops) {
            let chunk: Vec<(usize, Vec<f32>)> =
                batch.iter().map(|&i| crops[i].clone()).collect();
            let _ = batch[0];
            let (data, widths, wm) = pad_batch(&chunk);
            let b = chunk.len() as i64;
            let x = ort::value::Tensor::from_array(([b, 1, LINE_HEIGHT as i64, wm as i64], data))
                .map_err(|e| format!("rec input: {e:?}"))?;
            let w = ort::value::Tensor::from_array((vec![b], widths))
                .map_err(|e| format!("rec widths: {e:?}"))?;
            let outputs = self
                .rec
                .run(ort::inputs!["x" => x, "widths" => w])
                .map_err(|e| format!("rec run: {e:?}"))?;
            let (ids_shape, ids) = outputs["ids"]
                .try_extract_tensor::<i64>()
                .map_err(|e| format!("rec ids: {e:?}"))?;
            let (_, confs) = outputs["conf"]
                .try_extract_tensor::<f32>()
                .map_err(|e| format!("rec conf: {e:?}"))?;
            let (_, lengths) = outputs["lengths"]
                .try_extract_tensor::<i64>()
                .map_err(|e| format!("rec lengths: {e:?}"))?;
            let t = ids_shape.iter().nth(1).copied().unwrap_or(0) as usize;
            for k in 0..chunk.len() {
                let orig = batch[k];
                let n = (lengths[k].max(0) as usize).min(t);
                let (bx0, _, bx1, by0, by1) = boxes[orig];
                let bh = (by1 - by0).max(1.0);
                // token page-x: timestep → ink px → crop px (rec_preproc decode)
                let ix0 = (bx0 - PAD_X * bh).floor();
                let cw = (bx1 + PAD_X * bh).ceil() - ix0;
                let f = cw / chunk[k].0 as f32;
                let stride = wm as f32 / t.max(1) as f32;
                let mut toks: Vec<(String, f32)> = Vec::new();
                let mut conf_sum = 0.0f32;
                let mut conf_n = 0usize;
                let mut prev = 0i64;
                for i in 0..n {
                    let id = ids[k * t + i];
                    if id != 0 && id != prev {
                        let tok = self.vocab.get(id as usize - 1).cloned().unwrap_or_default();
                        let px = ix0 as f32 + (i as f32 + 0.5) * stride * f;
                        toks.push((tok, px));
                        conf_sum += confs[k * t + i];
                        conf_n += 1;
                    }
                    prev = id;
                }
                let conf = if conf_n > 0 { conf_sum / conf_n as f32 } else { 0.0 };
                let text: String = toks.iter().map(|(s, _)| s.as_str()).collect();
                out[orig] = LineRec {
                    box_: (bx0, by0, bx1, by1),
                    text: if conf >= MIN_CONF { text } else { String::new() },
                    conf,
                    toks: if conf >= MIN_CONF { toks } else { Vec::new() },
                };
            }
        }
        Ok(out)
    }

    fn reading_order(
        &mut self,
        boxes: &[(f32, f32, f32, f32)],
        texts: &[String],
        wh: (f32, f32),
    ) -> Result<(Vec<usize>, Option<(Vec<f32>, Vec<f32>)>), String> {
        let n = boxes.len();
        if n < 2 {
            return Ok((order_xycut(boxes), None));
        }
        let Some(order_sess) = &mut self.order else {
            return Ok((order_xycut(boxes), None));
        };
        let feats = line_features(boxes, wh, texts);
        let f = ort::value::Tensor::from_array(([1i64, n as i64, N_FEATS as i64], feats))
            .map_err(|e| format!("order feats: {e:?}"))?;
        let mask = ort::value::Tensor::from_array(([1i64, n as i64], vec![true; n]))
            .map_err(|e| format!("order mask: {e:?}"))?;
        let outputs = order_sess
            .run(ort::inputs!["feats" => f, "mask" => mask])
            .map_err(|e| format!("order run: {e:?}"))?;
        let (_a_shape, a) = outputs["a"]
            .try_extract_tensor::<f32>()
            .map_err(|e| format!("order a: {e:?}"))?;
        let (_, cls) = outputs["cls"]
            .try_extract_tensor::<f32>()
            .map_err(|e| format!("order cls: {e:?}"))?;
        let (_, same) = outputs["same"]
            .try_extract_tensor::<f32>()
            .map_err(|e| format!("order same: {e:?}"))?;
        let order = decode_blocks_order(&a, &same, boxes, Some(cls), n);
        Ok((order, Some((cls.to_vec(), same.to_vec()))))
    }

    /// OCR one grayscale page → lines (reading order) + layout blocks.
    pub fn ocr_page(&mut self, gray: &GrayPage) -> Result<NativePage, String> {
        let boxes = self.detect(gray)?;
        let recs = self.recognize(gray, &boxes)?;
        let rects: Vec<(f32, f32, f32, f32)> =
            boxes.iter().map(|b| (b.0, b.1, b.2, b.3)).collect();
        let texts: Vec<String> = recs.iter().map(|r| r.text.clone()).collect();
        let wh = (gray.w as f32, gray.h as f32);
        let (order, layout) = self.reading_order(&rects, &texts, wh)?;
        let lines: Vec<NativeLine> = order
            .iter()
            .map(|&i| NativeLine {
                box_: rects[i],
                text: texts[i].clone(),
                conf: recs[i].conf,
            })
            .collect();
        let mut blocks = Vec::new();
        if let Some((cls, same)) = layout {
            // split_blocks: runs of the reading order with same-logit > 0 between neighbours
            let mut runs: Vec<Vec<usize>> = Vec::new();
            for &i in &order {
                if let Some(last) = runs.last_mut() {
                    let prev = *last.last().expect("runs hold at least one line");
                    if same[prev * rects.len() + i] > 0.0 {
                        last.push(i);
                        continue;
                    }
                }
                runs.push(vec![i]);
            }
            for run in runs {
                let mut sum = [0.0f32; 10];
                let (mut bx0, mut by0, mut bx1, mut by1) =
                    (f32::MAX, f32::MAX, f32::MIN, f32::MIN);
                let mut joined = String::new();
                for &i in &run {
                    for k in 0..10 {
                        sum[k] += cls[i * 10 + k];
                    }
                    bx0 = bx0.min(rects[i].0);
                    by0 = by0.min(rects[i].1);
                    bx1 = bx1.max(rects[i].2);
                    by1 = by1.max(rects[i].3);
                    let t = texts[i].trim();
                    if t.is_empty() {
                        continue;
                    }
                    if joined.len() > 1
                        && joined.ends_with('-')
                        && joined
                            .chars()
                            .nth(joined.len() - 2)
                            .is_some_and(|c| c.is_alphabetic())
                        && t.chars().next().is_some_and(|c| c.is_lowercase())
                    {
                        joined.pop();
                        joined.push_str(t);
                    } else {
                        if !joined.is_empty() {
                            joined.push(' ');
                        }
                        joined.push_str(t);
                    }
                }
                let label = (0..10)
                    .max_by(|&x, &y| {
                        sum[x].partial_cmp(&sum[y]).unwrap_or(std::cmp::Ordering::Equal)
                    })
                    .unwrap_or(0);
                let label = BLOCK_CLASSES[label].to_string();
                // Table-class blocks: split lines into cells and rebuild the grid.
                let table = if label == "Table" {
                    let cells: Vec<(f32, f32, f32, f32, String)> = run
                        .iter()
                        .flat_map(|&i| {
                            split_line(&rects[i], &recs[i].toks, CELL_GAP)
                                .into_iter()
                                .map(|(bx, txt)| (bx.0, bx.1, bx.2, bx.3, txt))
                        })
                        .collect();
                    build_table_grid(&cells)
                } else {
                    None
                };
                blocks.push(NativeBlock {
                    label,
                    bbox: (bx0, by0, bx1, by1),
                    text: joined,
                    table,
                });
            }
        }
        Ok(NativePage { lines, blocks })
    }

    /// Parse a PDF: image-like pages are OCR'd in-process, text-rich pages
    /// keep their text layer (per-page routing for mixed documents).
    pub fn parse_document(
        &mut self,
        bytes: &[u8],
        min_page_confidence: f64,
    ) -> crate::Result<crate::Document> {
        self.min_page_confidence = min_page_confidence;
        super::pdf::parse_with_page_ocr(bytes, self, super::pdf_ocr::DEFAULT_DPI)
    }

    /// Full-control variant (min_chars / force) for the CLI.
    pub fn parse_document_opts(
        &mut self,
        bytes: &[u8],
        opts: &super::pdf_ocr::PdfOcrOptions,
    ) -> crate::Result<crate::Document> {
        self.min_page_confidence = opts.min_page_confidence;
        super::pdf::parse_with_page_ocr_opts(bytes, self, opts.dpi, opts.min_chars, opts.force)
    }
}

impl super::pdf::PageOcrEngine for IndraNative {
    fn ocr_page_png(&mut self, page_png: &[u8]) -> Result<(Vec<Block>, Option<f64>), String> {
        let gray = GrayPage::from_png(page_png)?;
        let page = self.ocr_page(&gray)?;
        let conf = (!page.lines.is_empty()).then(|| {
            page.lines.iter().map(|l| l.conf as f64).sum::<f64>() / page.lines.len() as f64
        });
        let mut blocks = Vec::new();
        for b in &page.blocks {
            match b.label.as_str() {
                // page furniture: dropped, consistent with DOCX policy
                "Page-Header" | "Page-Footer" => {}
                "Title" => blocks.push(Block::Heading {
                    block_index: blocks.len(),
                    level: 1,
                    text: b.text.clone(),
                    marks: Vec::new(),
                    source: Default::default(),
                }),
                "Section-Header" => blocks.push(Block::Heading {
                    block_index: blocks.len(),
                    level: 2,
                    text: b.text.clone(),
                    marks: Vec::new(),
                    source: Default::default(),
                }),
                "Table" if b.table.is_some() => {
                    let rows = b.table.as_ref().expect("checked");
                    let table_rows: Vec<Vec<crate::document_ast::Cell>> = rows
                        .iter()
                        .map(|r| {
                            r.iter()
                                .map(|c| crate::document_ast::Cell {
                                    text: c.clone(),
                                    colspan: 1,
                                    rowspan: 1,
                                })
                                .collect()
                        })
                        .collect();
                    // keep a paragraph twin of the table text for searchability
                    blocks.push(Block::Table {
                        block_index: blocks.len(),
                        rows: table_rows,
                        widths: Vec::new(),
                        source: Default::default(),
                    });
                }
                _ => blocks.push(Block::Paragraph {
                    block_index: blocks.len(),
                    text: b.text.clone(),
                    marks: Vec::new(),
                    source: Default::default(),
                }),
            }
        }
        self.last_page_confidence = conf;
        Ok((blocks, conf))
    }

    fn name(&self) -> &'static str {
        "indra-native"
    }
}
