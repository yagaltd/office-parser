//! Own table-structure recognition for OCR'd pages (feature `ocr-table-structure`).
//!
//! Contract: `.workflows/specs/task-13.spec` — human ruling 2026-11-08: "keeps
//! the Rust core lean if PaddleOCR logic ports". The vendored rusto-core /
//! rusto-nn crates are gone; the whole pipeline is ported in-tree:
//!
//! 1. The Indra engine (feature-gated, `indra_native.rs`) exposes the geometry
//!    it already computes when it emits a `Table` block: the grid bbox surfaces
//!    as an [`OcrLayoutRegion`] (label `"Table"`, raster pixels) and the cell
//!    rects as [`OcrWordBox`]es. RECOGNITION UNTOUCHED — with the feature off
//!    the engine keeps the defaulted [`PageOcrEngine::ocr_page`] and its
//!    payload is byte-identical to pre-task-13 builds.
//! 2. This module enriches the resulting [`OcrPage`]: for every `"Table"`
//!    region, crop → 488×488 preprocess → SLANet-plus ONNX through the rten
//!    runtime (already shipped by `ocr-rten`; rten's default features include
//!    `onnx_format`, so `.onnx` loads natively) → TableLabelDecode port →
//!    TableMatch port → own `<table>` rendering into `table_html`.
//!
//! Ports (behaviour-faithful, f64 where upstream uses Python floats):
//! - `ppstructure/table/predict_structure.py::build_pre_process_list` +
//!   `configs/table/SLANet_plus.yml` + `ppocr/data/imaug/table_ops.py`:
//!   ResizeTableImage (ratio = 488/max(h,w), `int()` truncation) →
//!   NormalizeImage (mean [0.485,0.456,0.406], std [0.229,0.224,0.225] —
//!   PaddleOCR loads BGR, so the constants sit on BGR slots; we decode RGB and
//!   reverse them, identical per-channel arithmetic) → PaddingTableImage
//!   (488×488 zeros in NORMALIZED space — upstream order is resize, normalize,
//!   pad) → ToCHW. shape_list = [h, w, ratio_h, ratio_w, pad_h, pad_w].
//!   cv2 INTER_LINEAR reproduced (align-corners=False mapping; f64 math instead
//!   of cv2's fixed-point u8 rounding — sub-quantisation difference only).
//! - `ppocr/postprocess/table_postprocess.py::TableLabelDecode.decode`:
//!   argmax per step, break at end token (`eos`, index 29) when step > 0, skip
//!   ignored tokens (`sos`/`eos`), `<td>`/`<td` tokens consume one row of the
//!   loc head; `_bbox_decode`: x *= pad_w / ratio_w, y *= pad_h / ratio_h,
//!   then min/max over the four corners (matcher's 8→4 conversion folded in).
//! - `ppstructure/table/matcher.py::TableMatch`: L1 distance
//!   (dis + min(dis2, dis3)), IoU, per-det best pred by (1 - IoU, L1) with
//!   first-index tie-break, multiple dets may share a pred (accumulate);
//!   `get_pred_html` port incl. the multi-det space joining, `<b>` strip
//!   quirks (substring test, prefix/suffix truncation) and `html.escape`.
//! - Model: SLANet-plus ONNX, RapidAI/RapidTable v2.0.0 pin
//!   https://www.modelscope.cn/models/RapidAI/RapidTable/resolve/v2.0.0/slanet-plus.onnx
//!   SHA256 d57a942af6a2f57d6a4a0372573c696a2379bf5857c45e2ac69993f3b334514b.
//!   The 48-token structure vocabulary is EMBEDDED below — extracted from
//!   the model's own ONNX `character` metadata prop (the shipped model is
//!   CH-dict; the EN-dict assumption was wrong), no external file.
//!
//! Transcript safety: enrichment runs AFTER the engine produced
//! blocks + confidence and only fills `table_html` (via struct-update
//! passthrough); an engine that already supplied markup wins. Coordinate
//! decision: Indra cell rects are translated into crop space before matching
//! (RapidTable semantics — det + structure share the crop; upstream PaddleOCR
//! end-to-end passes full-image boxes, which only coincides for tables at the
//! page origin). Upstream `_filter_ocr_result` is not ported: our cells are
//! the engine's Table-block cells, already region-scoped.
//!
//! No-model no-op (TASK 11 contract preserved): one one-shot warning, no
//! fields added — pages without table geometry return untouched before the
//! model is even resolved. Model is loaded per enrichment call (stateless
//! enrich, task-11 precedent; 7.8 MB load is ~ms).

use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

use super::pdf::{OcrLayoutRegion, OcrPage, OcrWordBox};

/// Default SLANet-plus model path, relative to the process working directory.
pub const DEFAULT_TABLE_MODEL: &str = "models/slanet-plus.onnx";

/// Environment override: exact path to the SLANet-plus model file (`.onnx`).
pub const TABLE_MODEL_ENV: &str = "OFFICE_PARSER_TABLE_MODEL";

/// SLANet-plus structure vocabulary, exact file order — extracted from the
/// model's own ONNX `character` metadata prop (48 tokens; +sos/eos = the 50
/// classes of the structure head). Attribute tokens carry their leading
/// space; `<td` + attribute(s) + `>` concatenate into `<td ...>` tags.
const STRUCTURE_VOCAB: [&str; 48] = [
    "<thead>",
    "</thead>",
    "<tbody>",
    "</tbody>",
    "<tr>",
    "</tr>",
    "<td",
    ">",
    "</td>",
    " colspan=\"2\"",
    " colspan=\"3\"",
    " colspan=\"4\"",
    " colspan=\"5\"",
    " colspan=\"6\"",
    " colspan=\"7\"",
    " colspan=\"8\"",
    " colspan=\"9\"",
    " colspan=\"10\"",
    " colspan=\"11\"",
    " colspan=\"12\"",
    " colspan=\"13\"",
    " colspan=\"14\"",
    " colspan=\"15\"",
    " colspan=\"16\"",
    " colspan=\"17\"",
    " colspan=\"18\"",
    " colspan=\"19\"",
    " colspan=\"20\"",
    " rowspan=\"2\"",
    " rowspan=\"3\"",
    " rowspan=\"4\"",
    " rowspan=\"5\"",
    " rowspan=\"6\"",
    " rowspan=\"7\"",
    " rowspan=\"8\"",
    " rowspan=\"9\"",
    " rowspan=\"10\"",
    " rowspan=\"11\"",
    " rowspan=\"12\"",
    " rowspan=\"13\"",
    " rowspan=\"14\"",
    " rowspan=\"15\"",
    " rowspan=\"16\"",
    " rowspan=\"17\"",
    " rowspan=\"18\"",
    " rowspan=\"19\"",
    " rowspan=\"20\"",
    "<td></td>",
];

/// AttnLabelDecode.add_special_char: `[beg] + vocab + [end]`.
const BEG_IDX: usize = 0;
const END_IDX: usize = 49;
/// TableLabelDecode.td_token.
const TD_TOKENS: [&str; 3] = ["<td>", "<td", "<td></td>"];

fn vocab_at(idx: usize) -> Option<&'static str> {
    match idx {
        BEG_IDX => Some("sos"),
        1..=48 => STRUCTURE_VOCAB.get(idx - 1).copied(),
        END_IDX => Some("eos"),
        _ => None,
    }
}

/// Model resolution rule (documented contract): [`TABLE_MODEL_ENV`] wins when
/// it names an existing file, else [`DEFAULT_TABLE_MODEL`]. `None` means the
/// enrichment is a no-op and `table_html` stays absent.
pub fn table_model_path() -> Option<PathBuf> {
    if let Some(dir) = std::env::var_os(TABLE_MODEL_ENV) {
        let p = PathBuf::from(dir);
        return p.is_file().then_some(p);
    }
    let p = PathBuf::from(DEFAULT_TABLE_MODEL);
    p.is_file().then_some(p)
}

/// Enrich one OCR'd page. Called from `parse_with_page_ocr_inner` after
/// `engine.ocr_page(png)`; `page_png` is the same rendered PNG the engine
/// received, so region/cell coords share the [`OcrWordBox`] raster space.
pub(crate) fn enrich_ocr_page(page: OcrPage, page_png: &[u8]) -> OcrPage {
    let model = table_model_path();
    enrich_with_model(page, page_png, model.as_deref())
}

/// Testable core: `model_path = None` models the no-model environment
/// deterministically (no env access).
fn enrich_with_model(page: OcrPage, page_png: &[u8], model_path: Option<&Path>) -> OcrPage {
    // Opt-in seam (task-10): an engine that already supplied markup wins.
    if page.table_html.is_some() {
        return page;
    }
    let Some(regions) = &page.layout_regions else {
        return page; // no geometry on this page: nothing to enrich, no warning
    };
    let regions: Vec<&OcrLayoutRegion> =
        regions.iter().filter(|r| r.label == "Table").collect();
    if regions.is_empty() {
        return page;
    }
    let Some(model_path) = model_path else {
        warn_table_model_missing();
        return page;
    };
    let Ok(img) = image::load_from_memory(page_png) else {
        eprintln!("office-parser: table-structure: page png decode failed");
        return page;
    };
    let rgb = img.to_rgb8();
    let page_rgb = PageRgb {
        w: rgb.width() as usize,
        h: rgb.height() as usize,
        buf: rgb.into_raw(),
    };
    let cells: &[OcrWordBox] = page.word_boxes.as_deref().unwrap_or(&[]);
    let mut htmls: Vec<String> = Vec::new();
    for region in &regions {
        match table_region_html(&page_rgb, region, cells, model_path) {
            Ok(Some(html)) => htmls.push(html),
            Ok(None) => {}
            Err(e) => eprintln!("office-parser: table-structure region failed: {e}"),
        }
    }
    let table_html = (!htmls.is_empty()).then(|| htmls.join("\n"));
    OcrPage {
        table_html,
        ..page // blocks, confidence, layout_regions, word_boxes pass through
    }
}

struct PageRgb {
    w: usize,
    h: usize,
    buf: Vec<u8>,
}

struct Crop {
    /// Crop origin in page pixels (for translating Indra cell rects).
    x0: i64,
    y0: i64,
    w: usize,
    h: usize,
    rgb: Vec<u8>,
}

/// cv2-style integer crop of the region, clamped to the page.
fn crop_region(page: &PageRgb, region: &OcrLayoutRegion) -> Option<Crop> {
    let px0 = (region.x0.floor() as i64).clamp(0, page.w as i64);
    let px1 = (region.x1.ceil() as i64).clamp(0, page.w as i64);
    let py0 = (region.y0.floor() as i64).clamp(0, page.h as i64);
    let py1 = (region.y1.ceil() as i64).clamp(0, page.h as i64);
    if px1 <= px0 || py1 <= py0 {
        return None;
    }
    let w = (px1 - px0) as usize;
    let h = (py1 - py0) as usize;
    let mut rgb = Vec::with_capacity(w * h * 3);
    for y in py0..py1 {
        let start = (y as usize * page.w + px0 as usize) * 3;
        rgb.extend_from_slice(&page.buf[start..start + w * 3]);
    }
    Some(Crop { x0: px0, y0: py0, w, h, rgb })
}

const SLANET_MAX_LEN: usize = 488;
/// NormalizeImage constants (SLANet_plus.yml) on PaddleOCR's BGR slots,
/// reversed for our RGB buffer.
const MEANS_RGB: [f64; 3] = [0.406, 0.456, 0.485];
const STDS_RGB: [f64; 3] = [0.225, 0.224, 0.229];

/// OpenCV INTER_LINEAR bilinear mapping (align-corners=False):
/// src = (dst + 0.5) * scale - 0.5, clamped at the borders.
fn resize_linear(rgb: &[u8], w: usize, h: usize, dw: usize, dh: usize) -> Vec<f64> {
    let mut out = vec![0f64; dw * dh * 3];
    let sx = w as f64 / dw as f64;
    let sy = h as f64 / dh as f64;
    let coords = |dst: usize, scale: f64, src_len: usize| -> (usize, f64) {
        let s = (dst as f64 + 0.5) * scale - 0.5;
        if s <= 0.0 {
            return (0, 0.0);
        }
        let f = s.floor();
        let i = f as usize;
        if i >= src_len - 1 {
            return (src_len - 1, 0.0);
        }
        (i, s - f)
    };
    for dy in 0..dh {
        let (y0, fy) = coords(dy, sy, h);
        let y1 = (y0 + 1).min(h - 1);
        for dx in 0..dw {
            let (x0, fx) = coords(dx, sx, w);
            let x1 = (x0 + 1).min(w - 1);
            for c in 0..3 {
                let p00 = rgb[(y0 * w + x0) * 3 + c] as f64;
                let p01 = rgb[(y0 * w + x1) * 3 + c] as f64;
                let p10 = rgb[(y1 * w + x0) * 3 + c] as f64;
                let p11 = rgb[(y1 * w + x1) * 3 + c] as f64;
                let top = p00 * (1.0 - fx) + p01 * fx;
                let bot = p10 * (1.0 - fx) + p11 * fx;
                out[(dy * dw + dx) * 3 + c] = top * (1.0 - fy) + bot * fy;
            }
        }
    }
    out
}

/// 488×488 preprocess (predict_structure.py order: resize keep-ratio →
/// normalize → zero-pad → CHW). Returns the [1,3,488,488] input and the
/// shape_list [h, w, ratio_h, ratio_w, pad_h, pad_w] that `_bbox_decode`
/// consumes.
fn preprocess_488(crop: &Crop) -> (Vec<f32>, [f64; 6]) {
    let (h, w) = (crop.h, crop.w);
    let ratio = SLANET_MAX_LEN as f64 / (h.max(w)) as f64;
    // upstream int() truncation
    let rh = ((h as f64 * ratio) as usize).clamp(1, SLANET_MAX_LEN);
    let rw = ((w as f64 * ratio) as usize).clamp(1, SLANET_MAX_LEN);
    let mut resized = resize_linear(&crop.rgb, w, h, rw, rh);
    for px in resized.chunks_exact_mut(3) {
        for c in 0..3 {
            px[c] = (px[c] / 255.0 - MEANS_RGB[c]) / STDS_RGB[c];
        }
    }
    // zero-pad in normalized space, paste at (0,0)
    let mut hwc = vec![0f64; SLANET_MAX_LEN * SLANET_MAX_LEN * 3];
    for y in 0..rh {
        let src = &resized[y * rw * 3..(y + 1) * rw * 3];
        let dst = y * SLANET_MAX_LEN * 3;
        hwc[dst..dst + rw * 3].copy_from_slice(src);
    }
    // CHW
    let mut chw = vec![0f32; 3 * SLANET_MAX_LEN * SLANET_MAX_LEN];
    for y in 0..SLANET_MAX_LEN {
        for x in 0..SLANET_MAX_LEN {
            for c in 0..3 {
                chw[(c * SLANET_MAX_LEN + y) * SLANET_MAX_LEN + x] =
                    hwc[(y * SLANET_MAX_LEN + x) * 3 + c] as f32;
            }
        }
    }
    (
        chw,
        [h as f64, w as f64, ratio, ratio, SLANET_MAX_LEN as f64, SLANET_MAX_LEN as f64],
    )
}

/// Run SLANet-plus over one crop. Outputs are assigned by last dimension
/// (8 = loc head, 30 = structure probs) — robust to export order, which
/// upstream fixes by index (`outputs[0]` loc, `outputs[1]` probs).
fn run_slanet(
    model_path: &Path,
    input: Vec<f32>,
) -> Result<(Vec<f32>, Vec<i64>, Vec<f32>, Vec<i64>), String> {
    use rten::{ValueOrView};
    use rten_tensor::prelude::*;
    use rten_tensor::Tensor;
    let model = rten::Model::load_file(model_path)
        .map_err(|e| format!("slanet load ({}): {e}", model_path.display()))?;
    let inputs = model.input_ids().to_vec();
    if inputs.len() != 1 {
        return Err(format!("slanet: expected 1 input, got {}", inputs.len()));
    }
    let out_ids = model.output_ids().to_vec();
    if out_ids.len() != 2 {
        return Err(format!("slanet: expected 2 outputs, got {}", out_ids.len()));
    }
    let shape = vec![1usize, 3, SLANET_MAX_LEN, SLANET_MAX_LEN];
    let value: ValueOrView = Tensor::from_data(&shape, input).into();
    let results = model
        .run(vec![(inputs[0], value)], &out_ids, None)
        .map_err(|e| format!("slanet run: {e:?}"))?;
    let mut loc = None;
    let mut probs = None;
    for v in &results {
        let t: Tensor<f32> = v
            .clone()
            .try_into()
            .map_err(|_| "slanet output: not f32".to_string())?;
        let dims: Vec<i64> = t.shape().iter().map(|&d| d as i64).collect();
        let data = t.into_data().to_vec();
        match dims.last() {
            Some(&8) => loc = Some((data, dims)),
            Some(&50) => probs = Some((data, dims)),
            _ => return Err(format!("slanet: unexpected output shape {dims:?}")),
        }
    }
    let (loc, loc_shape) = loc.ok_or("slanet: loc head (last dim 8) missing")?;
    let (probs, probs_shape) = probs.ok_or("slanet: structure head (last dim 50) missing")?;
    Ok((loc, loc_shape, probs, probs_shape))
}

struct StructureResult {
    tokens: Vec<String>,
    /// One [x0, y0, x1, y1] per `<td>`-token, crop pixel space.
    boxes: Vec<[f64; 4]>,
}

/// TableLabelDecode.decode + _bbox_decode + the matcher's 8→4 corner fold.
fn decode_structure(
    loc: &[f32],
    loc_shape: &[i64],
    probs: &[f32],
    probs_shape: &[i64],
    shape_list: &[f64; 6],
) -> Result<StructureResult, String> {
    let steps = probs_shape.get(1).copied().unwrap_or(0).max(0) as usize;
    let classes = probs_shape.get(2).copied().unwrap_or(0).max(0) as usize;
    let loc_w = loc_shape.get(2).copied().unwrap_or(0).max(0) as usize;
    if classes == 0 || loc_w != 8 {
        return Err(format!("slanet shapes: probs {probs_shape:?}, loc {loc_shape:?}"));
    }
    let [_h, _w, ratio_h, ratio_w, pad_h, pad_w] = *shape_list;
    let mut tokens = Vec::new();
    let mut boxes = Vec::new();
    for step in 0..steps {
        let row = probs
            .get(step * classes..(step + 1) * classes)
            .ok_or("probs truncated")?;
        // numpy argmax: first maximum wins
        let mut best = 0usize;
        let mut best_p = f64::NEG_INFINITY;
        for (i, &p) in row.iter().enumerate() {
            let p = p as f64;
            if p > best_p {
                best_p = p;
                best = i;
            }
        }
        if step > 0 && best == END_IDX {
            break;
        }
        if best == BEG_IDX || best == END_IDX {
            continue; // ignored tokens
        }
        let text = vocab_at(best).ok_or("class out of vocab range")?;
        if TD_TOKENS.contains(&text) {
            let b = loc
                .get(step * loc_w..(step + 1) * loc_w)
                .ok_or("loc truncated")?;
            let mut pts = [0f64; 8];
            for (k, v) in pts.iter_mut().enumerate() {
                // _bbox_decode: even (x) *= pad_w / ratio_w; odd (y) *= pad_h / ratio_h
                let (num, den) = if k % 2 == 0 { (pad_w, ratio_w) } else { (pad_h, ratio_h) };
                *v = if den != 0.0 { b[k] as f64 * num / den } else { b[k] as f64 };
            }
            let xs = [pts[0], pts[2], pts[4], pts[6]];
            let ys = [pts[1], pts[3], pts[5], pts[7]];
            boxes.push([
                xs.iter().cloned().fold(f64::MAX, f64::min),
                ys.iter().cloned().fold(f64::MAX, f64::min),
                xs.iter().cloned().fold(f64::MIN, f64::max),
                ys.iter().cloned().fold(f64::MIN, f64::max),
            ]);
        }
        tokens.push(text.to_string());
    }
    Ok(StructureResult { tokens, boxes })
}

/// matcher.py::distance.
fn l1_distance(a: &[f64; 4], b: &[f64; 4]) -> f64 {
    let (x1, y1, x2, y2) = (a[0], a[1], a[2], a[3]);
    let (x3, y3, x4, y4) = (b[0], b[1], b[2], b[3]);
    let dis = (x3 - x1).abs() + (y3 - y1).abs() + (x4 - x2).abs() + (y4 - y2).abs();
    let dis2 = (x3 - x1).abs() + (y3 - y1).abs();
    let dis3 = (x4 - x2).abs() + (y4 - y2).abs();
    dis + dis2.min(dis3)
}

/// matcher.py::compute_iou (axis-convention-agnostic: both boxes share one).
fn compute_iou(a: &[f64; 4], b: &[f64; 4]) -> f64 {
    let s1 = (a[2] - a[0]) * (a[3] - a[1]);
    let s2 = (b[2] - b[0]) * (b[3] - b[1]);
    let left = a[0].max(b[0]);
    let right = a[2].min(b[2]);
    let top = a[1].max(b[1]);
    let bottom = a[3].min(b[3]);
    if left >= right || top >= bottom {
        return 0.0;
    }
    let intersect = (right - left) * (bottom - top);
    let denom = s1 + s2 - intersect;
    if denom <= 0.0 {
        // degenerate zero-area boxes: Python would produce NaN here; we
        // degrade to "no overlap" so matching stays total.
        return 0.0;
    }
    intersect / denom
}

/// matcher.py::match_result — every det claims its best pred
/// (sorted by (1 - IoU, L1); first-index tie-break like `distances.index`),
/// and multiple dets may share one pred (the multi-line-cell case).
fn match_cells(dt_boxes: &[[f64; 4]], pred_boxes: &[[f64; 4]]) -> BTreeMap<usize, Vec<usize>> {
    let better = |a: (f64, f64), b: (f64, f64)| {
        let c = a.0.partial_cmp(&b.0).unwrap_or(std::cmp::Ordering::Equal);
        let c = c.then(a.1.partial_cmp(&b.1).unwrap_or(std::cmp::Ordering::Equal));
        c == std::cmp::Ordering::Less
    };
    let mut matched: BTreeMap<usize, Vec<usize>> = BTreeMap::new();
    for (i, gt) in dt_boxes.iter().enumerate() {
        let mut best: Option<(usize, (f64, f64))> = None;
        for (j, pred) in pred_boxes.iter().enumerate() {
            let key = (1.0 - compute_iou(gt, pred), l1_distance(gt, pred));
            if best.map_or(true, |(_, bk)| better(key, bk)) {
                best = Some((j, key));
            }
        }
        if let Some((j, _)) = best {
            matched.entry(j).or_default().push(i);
        }
    }
    matched
}

/// Python `html.escape` (quote=True).
fn escape_html(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&#x27;")
}

/// matcher.py::get_pred_html port (the `<td></td>` dict token — reachable
/// in the CH vocabulary — emits the empty cell itself, upstream-faithful).
/// Wrapping: predict_structure.py adds `<html><body><table>` around the model
/// stream; we keep only the `<table>` element — the documented
/// `table_html` shape (`docs/quality.md`) and what consumers embed.
fn render_table_html(
    tokens: &[String],
    matched: &BTreeMap<usize, Vec<usize>>,
    contents: &[String],
) -> String {
    let mut html = String::from("<table>");
    let mut td_index = 0usize;
    for tag in tokens {
        if tag.contains("</td>") {
            if let Some(det_idxs) = matched.get(&td_index) {
                let b_with =
                    det_idxs.len() > 1 && contents[det_idxs[0]].contains("<b>");
                if b_with {
                    html.push_str("<b>");
                }
                for (k, &di) in det_idxs.iter().enumerate() {
                    let mut content = contents[di].clone();
                    if det_idxs.len() > 1 {
                        if content.is_empty() {
                            continue;
                        }
                        if content.starts_with(' ') {
                            content = content[1..].to_string();
                        }
                        // Task-13 review F-2: boundary-safe strips (upstream
                        // slices code points; byte offsets would panic on
                        // multi-byte neighbours).
                        if let Some(stripped) = content.strip_prefix("<b>") {
                            content = stripped.to_string();
                        }
                        if let Some(stripped) = content.strip_suffix("</b>") {
                            content = stripped.to_string();
                        }
                        if content.is_empty() {
                            continue;
                        }
                        if k != det_idxs.len() - 1 && !content.ends_with(' ') {
                            content.push(' ');
                        }
                    }
                    html.push_str(&escape_html(&content));
                }
                if b_with {
                    html.push_str("</b>");
                }
            }
            if tag == "<td></td>" {
                // The CH structure dict carries this token (index 48);
                // upstream emits the empty cell itself — one td per token.
                html.push_str("<td></td>");
            } else {
                html.push_str(tag);
            }
            td_index += 1;
        } else {
            html.push_str(tag);
        }
    }
    html.push_str("</table>");
    html
}

/// One region → one `<table>` element, or `Ok(None)` when the region is
/// degenerate / the model produced no cells.
fn table_region_html(
    page: &PageRgb,
    region: &OcrLayoutRegion,
    cells: &[OcrWordBox],
    model_path: &Path,
) -> Result<Option<String>, String> {
    let Some(crop) = crop_region(page, region) else {
        return Ok(None);
    };
    let (input, shape_list) = preprocess_488(&crop);
    let (loc, loc_shape, probs, probs_shape) = run_slanet(model_path, input)?;
    let res = decode_structure(&loc, &loc_shape, &probs, &probs_shape, &shape_list)?;
    if res.boxes.is_empty() {
        return Ok(None);
    }
    // Indra cell rects → crop pixel space (RapidTable semantics).
    let mut dts: Vec<[f64; 4]> = Vec::new();
    let mut texts: Vec<String> = Vec::new();
    for w in cells {
        let inside = w.x0 < region.x1 && w.x1 > region.x0
            && w.y0 < region.y1 && w.y1 > region.y0;
        if !inside {
            continue;
        }
        let (cx, cy) = (crop.x0 as f32, crop.y0 as f32);
        dts.push([(w.x0 - cx) as f64, (w.y0 - cy) as f64, (w.x1 - cx) as f64, (w.y1 - cy) as f64]);
        texts.push(w.text.clone());
    }
    let matched = match_cells(&dts, &res.boxes);
    Ok(Some(render_table_html(&res.tokens, &matched, &texts)))
}

/// One warning per process — operators must learn why `table_html` is absent,
/// but not once per page.
fn warn_table_model_missing() {
    static WARNED: std::sync::atomic::AtomicBool = std::sync::atomic::AtomicBool::new(false);
    if !WARNED.swap(true, std::sync::atomic::Ordering::Relaxed) {
        eprintln!(
            "office-parser: ocr-table-structure enabled but no SLANet-plus model found \
             (set {TABLE_MODEL_ENV} to a model file, default {DEFAULT_TABLE_MODEL}); \
             table_html stays absent (engine geometry still surfaces)"
        );
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::document_ast::{Block, SourceSpan};

    fn vocab_full() -> Vec<&'static str> {
        (0..50).map(|i| vocab_at(i).expect("50 entries")).collect()
    }

    #[test]
    fn vocab_is_48_tokens_plus_specials() {
        let v = vocab_full();
        assert_eq!(v.len(), 50);
        assert_eq!(v[0], "sos");
        assert_eq!(v[49], "eos");
        assert_eq!(v[1], "<thead>");
        assert_eq!(v[7], "<td");
        assert_eq!(v[8], ">");
        assert_eq!(v[9], "</td>");
        assert_eq!(v[10], " colspan=\"2\"");
        assert_eq!(v[28], " colspan=\"20\"");
        assert_eq!(v[29], " rowspan=\"2\"");
        assert_eq!(v[48], "<td></td>");
    }

    /// One-hot probs builder over the 50-class vocab.
    fn one_hot(steps: &[usize]) -> (Vec<f32>, Vec<i64>) {
        let mut probs = vec![0f32; steps.len() * 50];
        for (s, &c) in steps.iter().enumerate() {
            probs[s * 50 + c] = 1.0;
        }
        (probs, vec![1, steps.len() as i64, 50])
    }

    fn loc_zero(steps: usize) -> (Vec<f32>, Vec<i64>) {
        (vec![0f32; steps * 8], vec![1, steps as i64, 8])
    }

    #[test]
    fn decode_breaks_at_eos_and_skips_specials() {
        // sos(0), thead(1), eos(49), tr(5) — eos at step 2 (>0) breaks.
        let (probs, probs_shape) = one_hot(&[0, 1, 49, 5]);
        let (loc, loc_shape) = loc_zero(4);
        let res = decode_structure(&loc, &loc_shape, &probs, &probs_shape,
            &[100.0, 100.0, 1.0, 1.0, 488.0, 488.0]).expect("decode");
        assert_eq!(res.tokens, ["<thead>"]);
        assert!(res.boxes.is_empty());
        // degenerate: eos at step 0 does not break (upstream idx>0 guard) but
        // is ignored; a beg-only stream decodes empty.
        let (probs, probs_shape) = one_hot(&[49, 5, 49]);
        let (loc, loc_shape) = loc_zero(3);
        let res = decode_structure(&loc, &loc_shape, &probs, &probs_shape,
            &[100.0, 100.0, 1.0, 1.0, 488.0, 488.0]).expect("decode");
        assert_eq!(res.tokens, ["<tr>"]);
    }

    #[test]
    fn decode_td_bbox_uses_pad_over_ratio() {
        // tokens: td(7), eos(49) — break at step 1.
        let (probs, probs_shape) = one_hot(&[7, 49]);
        let mut loc = vec![0f32; 2 * 8];
        // step 0 corners (the td decodes at step 0; eos at step 1 breaks
        // before its loc row would ever be read):
        let pts = [0.25f32, 0.5, 0.5, 0.5, 0.5, 1.0, 0.25, 1.0];
        loc[0..8].copy_from_slice(&pts);
        let loc_shape = vec![1i64, 2, 8];
        let res = decode_structure(&loc, &loc_shape, &probs, &probs_shape,
            &[100.0, 200.0, 0.25, 0.5, 488.0, 488.0]).expect("decode");
        assert_eq!(res.tokens, ["<td"]);
        // x = v * pad_w / ratio_w = v * 488 / 0.25; y = v * 488 / 0.5
        assert_eq!(res.boxes.len(), 1);
        // _bbox_decode: x = v * pad_w / ratio_w = v * 488 / 0.5;
        // y = v * pad_h / ratio_h = v * 488 / 0.25 (shape_list slots:
        // [h, w, ratio_h, ratio_w, pad_h, pad_w]).
        assert_eq!(res.boxes[0], [244.0, 976.0, 488.0, 1952.0]);
    }

    #[test]
    fn iou_matches_upstream_math() {
        let a = [0.0, 0.0, 2.0, 2.0];
        assert_eq!(compute_iou(&a, &a), 1.0);
        let b = [10.0, 10.0, 12.0, 12.0];
        assert_eq!(compute_iou(&a, &b), 0.0);
        let c = [1.0, 1.0, 3.0, 3.0];
        let iou = compute_iou(&a, &c); // 1 / 7
        assert!((iou - 1.0 / 7.0).abs() < 1e-12);
    }

    #[test]
    fn l1_distance_matches_upstream_math() {
        let a = [1.0, 2.0, 3.0, 4.0];
        let b = [5.0, 6.0, 7.0, 8.0];
        // dis = 16, dis2 = 8, dis3 = 8 → 24
        assert_eq!(l1_distance(&a, &b), 24.0);
    }

    #[test]
    fn match_cells_accumulates_multiple_dets_per_pred() {
        let preds = [[0.0, 0.0, 10.0, 10.0], [40.0, 0.0, 50.0, 10.0]];
        let dets = [[0.0, 0.0, 10.0, 10.0], [1.0, 1.0, 9.0, 9.0]];
        let m = match_cells(&dets, &preds);
        assert_eq!(m.get(&0), Some(&vec![0, 1]));
        assert!(!m.contains_key(&1));
    }

    #[test]
    fn render_escapes_and_joins_multi_match_cells() {
        let tokens = ["<tr>".to_string(), "<td>".to_string(), "</td>".to_string(), "</tr>".to_string()];
        let mut m = BTreeMap::new();
        m.insert(0usize, vec![0usize, 1usize]);
        let contents = ["a<b".to_string(), " b".to_string()];
        // content0: k!=last, no trailing space → "a "; escaped → "a&lt;b "…
        // python: "a<b" has no leading space, no <b>; sep ' ' appended → "a<b "
        // then escaped "a&lt;b "; content1 strips leading space → "b".
        let html = render_table_html(&tokens, &m, &contents);
        assert_eq!(html, "<table><tr><td>a&lt;b b</td></tr></table>");
        // empty cell → <td></td>
        let html = render_table_html(&tokens, &BTreeMap::new(), &[]);
        assert_eq!(html, "<table><tr><td></td></tr></table>");
    }

    #[test]
    fn escape_html_matches_python_html_escape() {
        assert_eq!(escape_html("&<>\"'"), "&amp;&lt;&gt;&quot;&#x27;");
    }

    /// 2x1 crop: ratio = 488/2 = 244, so rh = rw = 244 rows/cols are resized
    /// content and rows 244..488 are genuine zero-pad (a 1x1 crop would be
    /// upscaled to fill the whole 488x488 canvas — no padding to assert).
    fn crop_1x1() -> Crop {
        Crop { x0: 0, y0: 0, w: 2, h: 1, rgb: vec![255, 128, 0, 255, 128, 0] }
    }

    #[test]
    fn preprocess_shape_list_ratio_and_truncation() {
        let crop = Crop { x0: 0, y0: 0, w: 488, h: 976, rgb: vec![0; 488 * 976 * 3] };
        let (input, shape) = preprocess_488(&crop);
        assert_eq!(input.len(), 3 * 488 * 488);
        assert_eq!(shape, [976.0, 488.0, 0.5, 0.5, 488.0, 488.0]);
    }

    #[test]
    fn preprocess_normalizes_with_bgr_constants_and_zero_pads() {
        let (input, _shape) = preprocess_488(&crop_1x1());
        // channel 0 (R): (1 - 0.406)/0.225 ; 1 (G): (128/255 - 0.456)/0.224 ;
        // 2 (B): (0 - 0.485)/0.229 — BGR-ordered constants, reversed.
        let r = (1.0 - 0.406) / 0.225;
        let g = (128.0 / 255.0 - 0.456) / 0.224;
        let b = -0.485 / 0.229;
        assert!((input[0] - r as f32).abs() < 1e-6);
        assert!((input[488 * 488] - g as f32).abs() < 1e-6);
        assert!((input[2 * 488 * 488] - b as f32).abs() < 1e-6);
        // padded corner stays zero (normalized space)
        assert_eq!(input[487 * 488 + 487], 0.0);
        assert_eq!(input[488 * 488 + 487 * 488 + 487], 0.0);
    }

    fn para(text: &str) -> Block {
        Block::Paragraph {
            block_index: 0,
            text: text.to_string(),
            marks: Vec::new(),
            source: SourceSpan::default(),
        }
    }

    fn page_with_geometry() -> OcrPage {
        OcrPage {
            blocks: vec![para("hello")],
            confidence: Some(0.9),
            layout_regions: Some(vec![OcrLayoutRegion {
                label: "Table".to_string(),
                x0: 40.0, y0: 60.0, x1: 560.0, y1: 300.0,
            }]),
            word_boxes: Some(vec![OcrWordBox {
                text: "alpha".to_string(),
                x0: 40.0, y0: 62.0, x1: 280.0, y1: 80.0,
            }]),
            table_html: None,
        }
    }

    /// No model ⇒ no-op: transcript + confidence untouched, all rich fields
    /// keep their engine-supplied values. `&[]` PNG proves no decode happens
    /// before the model resolves.
    #[test]
    fn enrich_no_model_leaves_page_untouched() {
        let enriched = enrich_with_model(page_with_geometry(), &[], None);
        assert_eq!(enriched.blocks.len(), 1);
        assert_eq!(enriched.confidence, Some(0.9));
        assert!(enriched.layout_regions.is_some());
        assert!(enriched.word_boxes.is_some());
        assert!(enriched.table_html.is_none());
    }

    /// An engine that already produced markup wins (task-10 opt-in seam).
    #[test]
    fn enrich_engine_supplied_html_wins() {
        let mut page = page_with_geometry();
        page.table_html = Some("<table><tr><td>x</td></tr></table>".to_string());
        let enriched = enrich_with_model(page, &[], None);
        assert_eq!(
            enriched.table_html.as_deref(),
            Some("<table><tr><td>x</td></tr></table>")
        );
    }

    /// Pages without table geometry return untouched — never a warning path,
    /// never a decode.
    #[test]
    fn enrich_no_table_geometry_is_noop() {
        let page = OcrPage {
            blocks: vec![para("plain")],
            confidence: Some(0.9),
            layout_regions: None,
            word_boxes: None,
            table_html: None,
        };
        let enriched = enrich_with_model(page, &[], None);
        assert!(enriched.layout_regions.is_none());
        assert!(enriched.table_html.is_none());
    }

    /// Bogus model path ⇒ region pipeline fails, page returned unchanged.
    #[test]
    fn enrich_bogus_model_leaves_page_untouched() {
        let enriched = enrich_with_model(
            page_with_geometry(),
            &[],
            Some(Path::new("/nonexistent/office-parser-slanet.onnx")),
        );
        assert_eq!(enriched.blocks.len(), 1);
        assert_eq!(enriched.confidence, Some(0.9));
        assert!(enriched.table_html.is_none());
        assert!(enriched.layout_regions.is_some());
    }
}
