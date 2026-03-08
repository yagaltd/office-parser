use anyhow::{Context, Result};

use crate::formats::{ParsedOfficeDocument, spreadsheet};

pub fn parse(bytes: &[u8]) -> crate::Result<crate::Document> {
    let parsed = parse_ods(bytes, crate::spreadsheet::ParseOptions::default())
        .map_err(|e| crate::Error::from_parse(crate::Format::Ods, e))?;
    Ok(super::finalize(crate::Format::Ods, parsed))
}

pub fn parse_with_options(
    bytes: &[u8],
    opts: crate::spreadsheet::ParseOptions,
) -> crate::Result<crate::Document> {
    let parsed =
        parse_ods(bytes, opts).map_err(|e| crate::Error::from_parse(crate::Format::Ods, e))?;
    Ok(super::finalize(crate::Format::Ods, parsed))
}

fn parse_ods(bytes: &[u8], limits: spreadsheet::Limits) -> Result<ParsedOfficeDocument> {
    use calamine::{Data, Ods, Reader};
    use std::collections::HashMap;
    use std::io::Cursor;

    let mut wb: Ods<_> = Ods::new(Cursor::new(bytes)).context("open ods")?;

    let formula_by_sheet: HashMap<String, HashMap<String, String>> =
        ods_formulas(bytes, &limits).unwrap_or_default();

    let mut blocks = Vec::new();
    let mut warnings: Vec<String> = Vec::new();
    let mut sheets_meta: Vec<serde_json::Value> = Vec::new();
    let mut formulas_meta: Vec<serde_json::Value> = Vec::new();
    let mut block_index: usize = 0;

    let sheet_names: Vec<String> = wb.sheet_names().to_vec();
    if sheet_names.len() > limits.max_sheets {
        warnings.push(format!(
            "sheet count truncated: {} -> {}",
            sheet_names.len(),
            limits.max_sheets
        ));
    }

    for name in sheet_names.into_iter().take(limits.max_sheets) {
        blocks.push(spreadsheet::sheet_heading(block_index, &name));
        block_index += 1;

        let mut sheet_warnings: Vec<String> = Vec::new();
        let mut sheet_truncated = false;

        let formula_by_a1: HashMap<String, String> =
            formula_by_sheet.get(&name).cloned().unwrap_or_default();

        let range = match wb.worksheet_range(&name) {
            Ok(r) => r,
            Err(e) => {
                blocks.push(spreadsheet::warning_paragraph(
                    block_index,
                    &format!("warning: failed to read worksheet: {e}"),
                ));
                block_index += 1;
                continue;
            }
        };

        let (nrows, ncols) = range.get_size();
        let rows_cap = nrows.min(limits.max_rows_per_sheet);
        let cols_cap = ncols.min(limits.max_cols_per_sheet);
        if nrows > rows_cap {
            sheet_truncated = true;
            if limits.tail_rows > 0 && rows_cap >= 2 {
                let tail = limits.tail_rows.min(rows_cap / 2);
                let head = rows_cap.saturating_sub(tail);
                sheet_warnings.push(format!(
                    "sheet '{name}' rows truncated: {} -> {} (head {} + tail {})",
                    nrows, rows_cap, head, tail
                ));
            } else {
                sheet_warnings.push(format!(
                    "sheet '{name}' rows truncated: {} -> {}",
                    nrows, rows_cap
                ));
            }
        }
        if ncols > cols_cap {
            sheet_truncated = true;
            sheet_warnings.push(format!(
                "sheet '{name}' cols truncated: {} -> {}",
                ncols, cols_cap
            ));
        }

        if !sheet_warnings.is_empty() {
            blocks.push(spreadsheet::warning_paragraph(
                block_index,
                &format!("warning: {}", sheet_warnings.join("; ")),
            ));
            block_index += 1;
            warnings.extend(sheet_warnings.iter().cloned());
        }

        let mut grid: Vec<Vec<spreadsheet::CellDisplay>> =
            vec![vec![spreadsheet::CellDisplay::default(); cols_cap]; rows_cap];
        let mut max_r_used: Option<usize> = None;
        let mut max_c_used: Option<usize> = None;

        let (head_rows, tail_rows) = if nrows > rows_cap && limits.tail_rows > 0 && rows_cap >= 2 {
            let tail = limits.tail_rows.min(rows_cap / 2);
            let head = rows_cap.saturating_sub(tail);
            (head, tail)
        } else {
            (rows_cap, 0)
        };

        let mut row_map: Vec<usize> = Vec::with_capacity(rows_cap);
        for r in 0..rows_cap {
            let actual_r = if r < head_rows {
                r
            } else if tail_rows > 0 {
                (nrows - tail_rows) + (r - head_rows)
            } else {
                r
            };
            row_map.push(actual_r);
        }

        for r in 0..rows_cap {
            let actual_r = row_map[r];
            for c in 0..cols_cap {
                let a1 = spreadsheet::a1_ref(actual_r, c);
                let dt: Option<&Data> = range.get((actual_r, c));
                let formula = formula_by_a1.get(&a1).cloned();
                let cd = cell_display(dt, formula, &name, &a1, &mut formulas_meta, &limits);
                if !cd.is_empty() {
                    max_r_used = Some(max_r_used.map(|v| v.max(r)).unwrap_or(r));
                    max_c_used = Some(max_c_used.map(|v| v.max(c)).unwrap_or(c));
                }
                grid[r][c] = cd;
            }
        }

        let Some(max_r_used) = max_r_used else {
            sheets_meta.push(serde_json::json!({"name": name, "used": false}));
            continue;
        };
        let Some(max_c_used) = max_c_used else {
            sheets_meta.push(serde_json::json!({"name": name, "used": false}));
            continue;
        };

        let used_rows = max_r_used + 1;
        let used_cols = max_c_used + 1;
        grid.truncate(used_rows);
        row_map.truncate(used_rows);
        for r in &mut grid {
            r.truncate(used_cols);
        }

        let sheet_warnings_for_meta = sheet_warnings.clone();
        sheets_meta.push(serde_json::json!({
            "name": name,
            "used": true,
            "rows": used_rows,
            "cols": used_cols,
            "rows_total": nrows,
            "cols_total": ncols,
            "truncated": sheet_truncated,
            "emitted_tables": 0,
            "warnings": sheet_warnings_for_meta
        }));

        let mut emitted_tables = 0usize;
        let mut group_by_used = false;

        let regions = spreadsheet::split_regions(&grid, limits.empty_row_run);
        let regions_len = regions.len();
        let mut table_idx = 0usize;
        for (rs, re) in regions {
            let region_rows = &grid[rs..re];
            let region_row_map = &row_map[rs..re];
            let Some((cs, ce)) = spreadsheet::region_col_bounds(region_rows) else {
                continue;
            };

            let col_splits = spreadsheet::split_region_columns(region_rows, cs, ce);
            let col_splits_len = col_splits.len();
            for (cs2, ce2) in col_splits {
                let has_header = spreadsheet::first_row_looks_like_header(region_rows);
                let (group_by_col, group_label) = limits
                    .group_by
                    .as_deref()
                    .and_then(|spec| {
                        spreadsheet::resolve_group_by_col_and_label(
                            region_rows,
                            has_header,
                            spec,
                            cs2,
                            ce2,
                        )
                    })
                    .map(|(c, l)| (Some(c), Some(l)))
                    .unwrap_or((None, None));
                if group_by_col.is_some() {
                    group_by_used = true;
                }

                let segs = spreadsheet::semantic_segments(
                    region_rows,
                    has_header,
                    limits.max_table_rows_per_segment,
                    group_by_col,
                );
                let segs_len = segs.len();

                table_idx += 1;
                if regions_len > 1 || col_splits_len > 1 || segs_len > 1 {
                    blocks.push(spreadsheet::table_heading(block_index, table_idx));
                    block_index += 1;
                }

                for seg in segs {
                    let ds = seg.start;
                    let de = seg.end;
                    if segs_len > 1 {
                        let start1 = region_row_map[ds] + 1;
                        let end1 = region_row_map[de - 1] + 1;
                        if let (Some(label), Some(key)) =
                            (group_label.as_deref(), seg.group_key.as_deref())
                        {
                            blocks.push(spreadsheet::group_rows_heading(
                                block_index,
                                label,
                                key,
                                start1,
                                end1,
                            ));
                        } else {
                            blocks.push(spreadsheet::rows_heading(block_index, start1, end1));
                        }
                        block_index += 1;
                    }

                    let hdr = if has_header {
                        region_rows.first().map(|r| r.as_slice())
                    } else {
                        None
                    };
                    let data: Vec<&Vec<spreadsheet::CellDisplay>> =
                        region_rows[ds..de].iter().collect();
                    blocks.push(spreadsheet::build_table_block_from_parts(
                        block_index,
                        hdr,
                        &data,
                        cs2,
                        ce2,
                        limits.drop_empty_cols,
                    ));
                    block_index += 1;
                    emitted_tables += 1;
                }
            }
        }

        if limits.group_by.is_some() && !group_by_used {
            let msg = format!(
                "group_by '{}' did not match any column (try a column letter like 'A' or a header name)",
                limits.group_by.as_deref().unwrap_or("")
            );
            blocks.push(spreadsheet::warning_paragraph(
                block_index,
                &format!("warning: {msg}"),
            ));
            block_index += 1;
            warnings.push(msg);
        }

        if let Some(m) = sheets_meta.last_mut() {
            if let Some(o) = m.as_object_mut() {
                o.insert(
                    "emitted_tables".to_string(),
                    serde_json::json!(emitted_tables),
                );
            }
        }
    }

    let mut extra = serde_json::json!({
        "sheets": sheets_meta,
    });
    if !warnings.is_empty() {
        extra["warnings"] = serde_json::json!(warnings);
    }
    if !formulas_meta.is_empty() {
        extra["formulas"] = serde_json::json!(formulas_meta);
    }

    Ok(ParsedOfficeDocument {
        blocks,
        images: Vec::new(),
        metadata_json: extra,
    })
}

fn cell_display(
    dt: Option<&calamine::Data>,
    formula: Option<String>,
    sheet_name: &str,
    a1: &str,
    formulas_meta: &mut Vec<serde_json::Value>,
    limits: &spreadsheet::Limits,
) -> spreadsheet::CellDisplay {
    use calamine::Data;

    let mut out = spreadsheet::CellDisplay::default();
    if let Some(dt) = dt {
        match dt {
            Data::Empty => {}
            Data::String(s) => {
                out.text = spreadsheet::normalize_cell_text(s.clone(), limits.max_cell_chars)
            }
            Data::Float(f) => {
                out.text = if (f.fract()).abs() < 1e-9 {
                    format!("{}", *f as i64)
                } else {
                    f.to_string()
                };
            }
            Data::Int(i) => out.text = i.to_string(),
            Data::Bool(b) => out.text = if *b { "TRUE" } else { "FALSE" }.to_string(),
            Data::DateTime(v) => out.text = v.to_string(),
            Data::DateTimeIso(v) => out.text = v.clone(),
            Data::DurationIso(v) => out.text = v.clone(),
            Data::Error(e) => {
                out.is_error = true;
                out.text = format!("#ERR:{e}");
            }
        }
    }

    if let Some(f) = formula {
        let f = f.trim().to_string();
        if !f.is_empty() {
            out.formula = Some(f.clone());
            out.has_cached_value = !out.text.trim().is_empty();
            formulas_meta.push(serde_json::json!({
                "sheet": sheet_name,
                "cell": a1,
                "formula": f,
                "has_cached_value": out.has_cached_value,
                "is_error": out.is_error,
            }));

            let f = out.formula.clone().unwrap();
            if out.has_cached_value || out.is_error {
                out.text = format!("{} ({})", out.text.trim(), f);
            } else {
                out.text = format!("={f}");
            }
        }
    }
    out
}

fn ods_formulas(
    bytes: &[u8],
    limits: &spreadsheet::Limits,
) -> Result<std::collections::HashMap<String, std::collections::HashMap<String, String>>> {
    use quick_xml::events::Event;

    fn local_name(n: &[u8]) -> &[u8] {
        n.split(|b| *b == b':').last().unwrap_or(n)
    }

    let xml = super::read_zip_file_utf8(bytes, "content.xml").context("read content.xml")?;
    let mut out: std::collections::HashMap<String, std::collections::HashMap<String, String>> =
        Default::default();

    let mut r = quick_xml::Reader::from_str(&xml);
    r.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut cur_sheet: Option<String> = None;
    let mut row0: usize = 0;
    let mut col0: usize = 0;
    let mut row_repeat: usize = 1;

    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"table" => {
                cur_sheet = None;
                row0 = 0;
                col0 = 0;
                for a in e.attributes().flatten() {
                    if local_name(a.key.as_ref()) == b"name" {
                        cur_sheet = Some(a.unescape_value().unwrap_or_default().to_string());
                        break;
                    }
                }
            }
            Ok(Event::Empty(e)) if local_name(e.name().as_ref()) == b"table" => {
                cur_sheet = None;
                row0 = 0;
                col0 = 0;
                for a in e.attributes().flatten() {
                    if local_name(a.key.as_ref()) == b"name" {
                        cur_sheet = Some(a.unescape_value().unwrap_or_default().to_string());
                        break;
                    }
                }
            }

            Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"table-row" => {
                col0 = 0;
                row_repeat = 1;
                for a in e.attributes().flatten() {
                    if local_name(a.key.as_ref()) == b"number-rows-repeated" {
                        row_repeat = a
                            .unescape_value()
                            .ok()
                            .and_then(|s| s.parse::<usize>().ok())
                            .unwrap_or(1)
                            .max(1);
                        break;
                    }
                }
            }
            Ok(Event::Empty(e)) if local_name(e.name().as_ref()) == b"table-row" => {
                // Completely empty row.
                let mut rr = 1usize;
                for a in e.attributes().flatten() {
                    if local_name(a.key.as_ref()) == b"number-rows-repeated" {
                        rr = a
                            .unescape_value()
                            .ok()
                            .and_then(|s| s.parse::<usize>().ok())
                            .unwrap_or(1)
                            .max(1);
                        break;
                    }
                }
                row0 = row0.saturating_add(rr);
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"table-row" => {
                row0 = row0.saturating_add(row_repeat);
                row_repeat = 1;
            }

            Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"table-cell" => {
                let mut cr = 1usize;
                let mut f: Option<String> = None;
                for a in e.attributes().flatten() {
                    let k = local_name(a.key.as_ref());
                    if k == b"number-columns-repeated" {
                        cr = a
                            .unescape_value()
                            .ok()
                            .and_then(|s| s.parse::<usize>().ok())
                            .unwrap_or(1)
                            .max(1);
                    }
                    if k == b"formula" {
                        f = Some(a.unescape_value().unwrap_or_default().to_string());
                    }
                }

                if let (Some(sheet), Some(formula)) = (cur_sheet.as_deref(), f) {
                    let formula = formula.trim().trim_start_matches("of:=").to_string();
                    if !formula.is_empty() {
                        let sheet_map = out.entry(sheet.to_string()).or_default();
                        for dr in 0..row_repeat {
                            let rr = row0 + dr;
                            if rr >= limits.max_rows_per_sheet {
                                break;
                            }
                            for dc in 0..cr {
                                let cc = col0 + dc;
                                if cc >= limits.max_cols_per_sheet {
                                    break;
                                }
                                let a1 = spreadsheet::a1_ref(rr, cc);
                                sheet_map.insert(a1, formula.clone());
                            }
                        }
                    }
                }

                col0 = col0.saturating_add(cr);
            }

            Ok(Event::Empty(e)) if local_name(e.name().as_ref()) == b"table-cell" => {
                let mut cr = 1usize;
                let mut f: Option<String> = None;
                for a in e.attributes().flatten() {
                    let k = local_name(a.key.as_ref());
                    if k == b"number-columns-repeated" {
                        cr = a
                            .unescape_value()
                            .ok()
                            .and_then(|s| s.parse::<usize>().ok())
                            .unwrap_or(1)
                            .max(1);
                    }
                    if k == b"formula" {
                        f = Some(a.unescape_value().unwrap_or_default().to_string());
                    }
                }

                if let (Some(sheet), Some(formula)) = (cur_sheet.as_deref(), f) {
                    let formula = formula.trim().trim_start_matches("of:=").to_string();
                    if !formula.is_empty() {
                        let sheet_map = out.entry(sheet.to_string()).or_default();
                        for dr in 0..row_repeat {
                            let rr = row0 + dr;
                            if rr >= limits.max_rows_per_sheet {
                                break;
                            }
                            for dc in 0..cr {
                                let cc = col0 + dc;
                                if cc >= limits.max_cols_per_sheet {
                                    break;
                                }
                                let a1 = spreadsheet::a1_ref(rr, cc);
                                sheet_map.insert(a1, formula.clone());
                            }
                        }
                    }
                }

                col0 = col0.saturating_add(cr);
            }

            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!(e).context("parse ods formulas")),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}
