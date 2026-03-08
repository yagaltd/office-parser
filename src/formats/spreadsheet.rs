use crate::document_ast::{Block, Cell, SourceSpan};

use crate::spreadsheet::ParseOptions;

pub(crate) type Limits = ParseOptions;

#[derive(Clone, Debug, Default)]
pub(crate) struct CellDisplay {
    pub text: String,
    pub is_error: bool,
    pub formula: Option<String>,
    pub has_cached_value: bool,
}

impl CellDisplay {
    pub fn is_empty(&self) -> bool {
        self.text.trim().is_empty() && !self.is_error && self.formula.is_none()
    }
}

pub(crate) fn normalize_cell_text(mut s: String, max_chars: usize) -> String {
    s = s.replace('\r', " ").replace('\n', " ");
    let t = s.trim();
    if t.chars().count() <= max_chars {
        return t.to_string();
    }
    let mut out = String::new();
    for (i, ch) in t.chars().enumerate() {
        if i >= max_chars {
            break;
        }
        out.push(ch);
    }
    out
}

pub(crate) fn excel_col_name(mut col0: usize) -> String {
    // 0 -> A, 25 -> Z, 26 -> AA
    let mut out = String::new();
    col0 += 1;
    while col0 > 0 {
        let rem = (col0 - 1) % 26;
        out.push((b'A' + rem as u8) as char);
        col0 = (col0 - 1) / 26;
    }
    out.chars().rev().collect()
}

pub(crate) fn excel_col_index_from_name(col: &str) -> Option<usize> {
    let col = col.trim();
    if col.is_empty() {
        return None;
    }
    let mut out: usize = 0;
    for ch in col.chars() {
        let up = ch.to_ascii_uppercase();
        if !(('A'..='Z').contains(&up)) {
            return None;
        }
        out = out * 26 + ((up as u8 - b'A') as usize + 1);
    }
    Some(out.saturating_sub(1))
}

pub(crate) fn a1_ref(row0: usize, col0: usize) -> String {
    format!("{}{}", excel_col_name(col0), row0 + 1)
}

pub(crate) fn split_regions(
    rows: &[Vec<CellDisplay>],
    empty_row_run: usize,
) -> Vec<(usize, usize)> {
    let empty_row_run = empty_row_run.max(1);
    let mut out = Vec::new();

    let mut in_region = false;
    let mut start = 0usize;
    let mut empty_run = 0usize;

    for (i, r) in rows.iter().enumerate() {
        let is_empty = r.iter().all(|c| c.is_empty());
        if !in_region {
            if !is_empty {
                in_region = true;
                start = i;
                empty_run = 0;
            }
            continue;
        }

        if is_empty {
            empty_run += 1;
            if empty_run >= empty_row_run {
                let end = i + 1 - empty_run;
                if end > start {
                    out.push((start, end));
                }
                in_region = false;
            }
        } else {
            empty_run = 0;
        }
    }

    if in_region {
        let end = rows.len().saturating_sub(empty_run);
        if end > start {
            out.push((start, end));
        }
    }

    out
}

pub(crate) fn split_region_columns(
    region_rows: &[Vec<CellDisplay>],
    col_start: usize,
    col_end: usize,
) -> Vec<(usize, usize)> {
    let cols = col_end.saturating_sub(col_start);
    if cols <= 1 {
        return vec![(col_start, col_end)];
    }

    let mut non_empty_by_col = vec![0usize; cols];
    for r in region_rows {
        for (i, c) in r.iter().enumerate().skip(col_start).take(cols) {
            if !c.is_empty() {
                non_empty_by_col[i - col_start] += 1;
            }
        }
    }

    let rows_n = region_rows.len().max(1);
    let mostly_empty_thresh = (rows_n / 50).max(1);

    let mut splits = Vec::new();
    let mut cur_start = 0usize;
    let mut i = 0usize;
    while i < cols {
        if non_empty_by_col[i] <= mostly_empty_thresh {
            let sep_start = i;
            let mut sep_end = i + 1;
            while sep_end < cols && non_empty_by_col[sep_end] <= mostly_empty_thresh {
                sep_end += 1;
            }

            let left_has = (cur_start..sep_start).any(|c| non_empty_by_col[c] > 0);
            let right_has = (sep_end..cols).any(|c| non_empty_by_col[c] > 0);
            if left_has && right_has {
                splits.push((col_start + cur_start, col_start + sep_start));
                cur_start = sep_end;
            }

            i = sep_end;
        } else {
            i += 1;
        }
    }

    splits.push((col_start + cur_start, col_end));

    splits.into_iter().filter(|(s, e)| e > s).collect()
}

fn norm_cmp(s: &str) -> String {
    s.trim().to_ascii_lowercase()
}

fn is_totalish_row(row: &[CellDisplay]) -> bool {
    let mut first_words: Vec<String> = Vec::new();
    for c in row.iter().filter(|c| !c.is_empty()) {
        let t = norm_cmp(&c.text);
        if !t.is_empty() {
            first_words.push(t);
        }
        if first_words.len() >= 2 {
            break;
        }
    }
    let has_total = first_words
        .iter()
        .any(|t| t.contains("total") || t.contains("subtotal"));
    if !has_total {
        return false;
    }
    let (ne, s, n) = row_stringiness(row);
    ne >= 2 && (n > 0 || s <= ne / 2)
}

fn is_repeated_header_row(row: &[CellDisplay], header_row: &[CellDisplay]) -> bool {
    let cols = header_row.len().max(row.len());
    let mut header_non_empty = 0usize;
    let mut matches = 0usize;
    for ci in 0..cols {
        let h = header_row
            .get(ci)
            .map(|c| norm_cmp(&c.text))
            .unwrap_or_default();
        if h.is_empty() {
            continue;
        }
        header_non_empty += 1;
        let v = row.get(ci).map(|c| norm_cmp(&c.text)).unwrap_or_default();
        if v == h {
            matches += 1;
        }
    }
    if header_non_empty == 0 {
        return false;
    }
    let ratio = (matches as f32) / (header_non_empty as f32);
    if ratio < 0.80 {
        return false;
    }
    let (_ne, s_like, n_like) = row_stringiness(row);
    s_like >= n_like
}

/// Return row index ranges (start..end) in `rows` to use as DATA rows (header excluded).
///
/// For `has_header=true`, `rows[0]` is treated as the header row.
pub(crate) fn semantic_row_ranges(
    rows: &[Vec<CellDisplay>],
    has_header: bool,
    max_table_rows_per_segment: usize,
) -> Vec<(usize, usize)> {
    if rows.is_empty() {
        return Vec::new();
    }

    let header_row = if has_header { rows.get(0) } else { None };
    let start0 = if has_header { 1usize } else { 0usize };
    if rows.len() <= start0 {
        return Vec::new();
    }

    let mut out: Vec<(usize, usize)> = Vec::new();
    let mut cur = start0;

    for i in start0..rows.len() {
        if let Some(h) = header_row {
            if i > cur + 3 && is_repeated_header_row(&rows[i], h) {
                if i > cur {
                    out.push((cur, i));
                }
                cur = i + 1; // skip the repeated header row
                continue;
            }
        }

        if is_totalish_row(&rows[i]) {
            let end = i + 1;
            if end > cur {
                out.push((cur, end));
            }
            cur = end;
        }
    }

    if rows.len() > cur {
        out.push((cur, rows.len()));
    }

    if out.is_empty() {
        out.push((start0, rows.len()));
    }

    let cap = max_table_rows_per_segment;
    if cap == 0 {
        return out;
    }

    let mut windowed: Vec<(usize, usize)> = Vec::new();
    for (s, e) in out {
        let mut i = s;
        while i < e {
            let j = (i + cap).min(e);
            windowed.push((i, j));
            i = j;
        }
    }
    windowed
}

#[derive(Clone, Debug)]
pub(crate) struct Segment {
    pub start: usize,
    pub end: usize,
    pub group_key: Option<String>,
}

pub(crate) fn resolve_group_by_col_and_label(
    rows: &[Vec<CellDisplay>],
    has_header: bool,
    spec: &str,
    col_start: usize,
    col_end: usize,
) -> Option<(usize, String)> {
    let spec_t = spec.trim();
    if spec_t.is_empty() {
        return None;
    }

    if spec_t.chars().all(|c| c.is_ascii_alphabetic()) {
        if let Some(ci) = excel_col_index_from_name(spec_t) {
            if ci >= col_start && ci < col_end {
                return Some((ci, excel_col_name(ci)));
            }
        }
    }

    if has_header {
        if let Some(h) = rows.first() {
            let want = norm_cmp(spec_t);
            for ci in col_start..col_end {
                let got = h.get(ci).map(|c| norm_cmp(&c.text)).unwrap_or_default();
                if !got.is_empty() && got == want {
                    let label = h.get(ci).map(|c| c.text.trim()).unwrap_or("key");
                    return Some((ci, label.to_string()));
                }
            }
        }
    }

    None
}

pub(crate) fn semantic_segments(
    rows: &[Vec<CellDisplay>],
    has_header: bool,
    max_table_rows_per_segment: usize,
    group_by_col: Option<usize>,
) -> Vec<Segment> {
    let mut ranges = semantic_row_ranges(rows, has_header, 0);
    if ranges.is_empty() {
        return Vec::new();
    }

    let cap = max_table_rows_per_segment;
    let mut out: Vec<Segment> = Vec::new();

    for (s, e) in ranges.drain(..) {
        if let Some(kc) = group_by_col {
            let mut cur_norm: Option<String> = None;
            let mut cur_display: Option<String> = None;
            let mut cur_start = s;

            for i in s..e {
                let key_raw = rows
                    .get(i)
                    .and_then(|r| r.get(kc))
                    .map(|c| c.text.trim())
                    .unwrap_or("");
                let key_norm = norm_cmp(key_raw);
                if key_norm.is_empty() {
                    continue;
                }
                if let Some(prev) = cur_norm.as_deref() {
                    if prev != key_norm {
                        if i > cur_start {
                            out.push(Segment {
                                start: cur_start,
                                end: i,
                                group_key: cur_display.take(),
                            });
                        }
                        cur_start = i;
                        cur_norm = Some(key_norm);
                        cur_display = Some(key_raw.to_string());
                    }
                } else {
                    cur_norm = Some(key_norm);
                    cur_display = Some(key_raw.to_string());
                }
            }

            if e > cur_start {
                out.push(Segment {
                    start: cur_start,
                    end: e,
                    group_key: cur_display.take(),
                });
            }
        } else {
            out.push(Segment {
                start: s,
                end: e,
                group_key: None,
            });
        }
    }

    if cap == 0 {
        return out;
    }

    let mut windowed: Vec<Segment> = Vec::new();
    for seg in out {
        let mut i = seg.start;
        while i < seg.end {
            let j = (i + cap).min(seg.end);
            windowed.push(Segment {
                start: i,
                end: j,
                group_key: seg.group_key.clone(),
            });
            i = j;
        }
    }
    windowed
}

pub(crate) fn region_col_bounds(region_rows: &[Vec<CellDisplay>]) -> Option<(usize, usize)> {
    let mut min_c: Option<usize> = None;
    let mut max_c: Option<usize> = None;
    for r in region_rows {
        for (ci, c) in r.iter().enumerate() {
            if c.is_empty() {
                continue;
            }
            min_c = Some(min_c.map(|v| v.min(ci)).unwrap_or(ci));
            max_c = Some(max_c.map(|v| v.max(ci)).unwrap_or(ci));
        }
    }
    let (min_c, max_c) = (min_c?, max_c?);
    Some((min_c, max_c + 1))
}

fn row_stringiness(row: &[CellDisplay]) -> (usize, usize, usize) {
    // (non_empty, string_like, numeric_like)
    let mut non_empty = 0usize;
    let mut string_like = 0usize;
    let mut numeric_like = 0usize;
    for c in row {
        let t = c.text.trim();
        if t.is_empty() {
            continue;
        }
        non_empty += 1;
        if t.parse::<f64>().is_ok() {
            numeric_like += 1;
        } else {
            string_like += 1;
        }
    }
    (non_empty, string_like, numeric_like)
}

pub(crate) fn first_row_looks_like_header(rows: &[Vec<CellDisplay>]) -> bool {
    let Some(r0) = rows.first() else {
        return false;
    };
    let (ne0, s0, n0) = row_stringiness(r0);
    if ne0 == 0 {
        return false;
    }
    if (s0 as f32) / (ne0 as f32) < 0.50 {
        return false;
    }
    let Some(r1) = rows.get(1) else {
        return true;
    };
    let (_ne1, _s1, n1) = row_stringiness(r1);
    n1 > 0 || n0 == 0
}

pub(crate) fn build_table_block_from_parts(
    block_index: usize,
    header_row: Option<&[CellDisplay]>,
    data_rows: &[&Vec<CellDisplay>],
    col_start: usize,
    col_end: usize,
    drop_empty_cols: bool,
) -> Block {
    let mut selected_cols: Vec<usize> = (col_start..col_end).collect();
    if drop_empty_cols {
        selected_cols.retain(|&c| {
            header_row
                .and_then(|r| r.get(c))
                .is_some_and(|cc| !cc.is_empty())
                || data_rows
                    .iter()
                    .any(|r| r.get(c).is_some_and(|cc| !cc.is_empty()))
        });
        if selected_cols.is_empty() {
            selected_cols.push(col_start);
        }
    }
    let cols = selected_cols.len().max(1);

    let mut out_rows: Vec<Vec<Cell>> = Vec::new();

    let mut header: Vec<Cell> = Vec::with_capacity(cols);
    if let Some(h) = header_row {
        for &c in &selected_cols {
            header.push(Cell {
                text: h.get(c).map(|cc| cc.text.clone()).unwrap_or_default(),
                colspan: 1,
                rowspan: 1,
            });
        }
    } else {
        for c in 0..cols {
            header.push(Cell {
                text: excel_col_name(c),
                colspan: 1,
                rowspan: 1,
            });
        }
    }
    out_rows.push(header);

    for r in data_rows {
        let mut out_r: Vec<Cell> = Vec::with_capacity(cols);
        for &c in &selected_cols {
            out_r.push(Cell {
                text: r.get(c).map(|cc| cc.text.clone()).unwrap_or_default(),
                colspan: 1,
                rowspan: 1,
            });
        }
        out_rows.push(out_r);
    }

    Block::Table {
        block_index,
        rows: out_rows,
        source: SourceSpan::default(),
    }
}

pub(crate) fn sheet_heading(block_index: usize, name: &str) -> Block {
    Block::Heading {
        block_index,
        level: 2,
        text: format!("Sheet: {}", name.trim()),
        source: SourceSpan::default(),
    }
}

pub(crate) fn table_heading(block_index: usize, table_idx_1: usize) -> Block {
    Block::Heading {
        block_index,
        level: 3,
        text: format!("Table {table_idx_1}"),
        source: SourceSpan::default(),
    }
}

pub(crate) fn named_table_heading(block_index: usize, name: &str) -> Block {
    Block::Heading {
        block_index,
        level: 3,
        text: format!("Table: {}", name.trim()),
        source: SourceSpan::default(),
    }
}

pub(crate) fn warning_paragraph(block_index: usize, msg: &str) -> Block {
    Block::Paragraph {
        block_index,
        text: msg.trim().to_string(),
        source: SourceSpan::default(),
    }
}

pub(crate) fn rows_heading(block_index: usize, start_row_1: usize, end_row_1: usize) -> Block {
    Block::Heading {
        block_index,
        level: 4,
        text: format!("Rows {start_row_1}–{end_row_1}"),
        source: SourceSpan::default(),
    }
}

pub(crate) fn group_rows_heading(
    block_index: usize,
    group_label: &str,
    group_key: &str,
    start_row_1: usize,
    end_row_1: usize,
) -> Block {
    let mut key = group_key.trim();
    if key.len() > 120 {
        key = &key[..120];
    }
    Block::Heading {
        block_index,
        level: 4,
        text: format!(
            "{}: {} (Rows {start_row_1}–{end_row_1})",
            group_label.trim(),
            key
        ),
        source: SourceSpan::default(),
    }
}
