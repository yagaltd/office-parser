use anyhow::{Context, Result};

use crate::formats::{ParsedOfficeDocument, spreadsheet};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum DateTimeKind {
    Date,
    Time,
    DateTime,
}

pub fn parse(bytes: &[u8]) -> crate::Result<crate::Document> {
    let parsed = parse_xlsx(bytes, crate::spreadsheet::ParseOptions::default())
        .map_err(|e| crate::Error::from_parse(crate::Format::Xlsx, e))?;
    Ok(super::finalize(crate::Format::Xlsx, parsed))
}

pub fn parse_with_options(
    bytes: &[u8],
    opts: crate::spreadsheet::ParseOptions,
) -> crate::Result<crate::Document> {
    let parsed =
        parse_xlsx(bytes, opts).map_err(|e| crate::Error::from_parse(crate::Format::Xlsx, e))?;
    Ok(super::finalize(crate::Format::Xlsx, parsed))
}

fn parse_xlsx(bytes: &[u8], limits: spreadsheet::Limits) -> Result<ParsedOfficeDocument> {
    use calamine::{Data, Reader, Xlsx};
    use std::collections::HashMap;
    use std::io::Cursor;

    let mut wb: Xlsx<_> = Xlsx::new(Cursor::new(bytes)).context("open xlsx")?;

    let name_to_path = xlsx_sheet_name_to_path(bytes).unwrap_or_default();
    let date_kinds_by_style = xlsx_date_time_kinds_by_xf(bytes).unwrap_or_default();

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

        let formula_by_a1: HashMap<String, String> = name_to_path
            .get(&name)
            .and_then(|p| xlsx_sheet_formulas(bytes, p).ok())
            .unwrap_or_default();

        let sheet_path = name_to_path.get(&name).cloned();

        let date_cells = sheet_path
            .as_deref()
            .and_then(|p| xlsx_sheet_date_time_cells(bytes, p, &date_kinds_by_style).ok())
            .unwrap_or_default();

        let named_tables = sheet_path
            .as_deref()
            .and_then(|p| xlsx_named_tables_for_sheet(bytes, p).ok())
            .unwrap_or_default();

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
                let dt: Option<&Data> = range.get((actual_r, c));
                let a1 = spreadsheet::a1_ref(actual_r, c);
                let formula = formula_by_a1.get(&a1).cloned();
                let cd = cell_display(
                    dt,
                    formula,
                    &name,
                    &a1,
                    date_cells.get(&a1).copied(),
                    &mut formulas_meta,
                    &limits,
                );
                if !cd.is_empty() {
                    max_r_used = Some(max_r_used.map(|v| v.max(r)).unwrap_or(r));
                    max_c_used = Some(max_c_used.map(|v| v.max(c)).unwrap_or(c));
                }
                grid[r][c] = cd;
            }
        }

        // Apply style-derived date/time formatting as a post-pass as well, since calamine values
        // can be type-erased to floats and we want deterministic formatting.
        if !date_cells.is_empty() {
            for r in 0..rows_cap {
                let actual_r = row_map[r];
                for c in 0..cols_cap {
                    let a1 = spreadsheet::a1_ref(actual_r, c);
                    let Some(kind) = date_cells.get(&a1).copied() else {
                        continue;
                    };
                    let cell = &mut grid[r][c];
                    if let Ok(v) = cell.text.trim().parse::<f64>() {
                        if excel_serial_looks_like_datetime(v) {
                            cell.text = excel_serial_to_iso(v, kind);
                        }
                    }
                }
            }
        }

        // Best-effort date/time inference from header text when XLSX styles are missing/unavailable.
        if rows_cap >= 2 {
            let mut inferred: Vec<Option<DateTimeKind>> = vec![None; cols_cap];
            for c in 0..cols_cap {
                let header = grid
                    .get(0)
                    .and_then(|r| r.get(c))
                    .map(|c| c.text.trim().to_ascii_lowercase())
                    .unwrap_or_default();
                inferred[c] = if header.contains("timestamp")
                    || header.contains("datetime")
                    || header.contains("time")
                    || header.contains("when")
                {
                    Some(DateTimeKind::DateTime)
                } else if header.contains("date") {
                    Some(DateTimeKind::Date)
                } else {
                    None
                };
            }

            for r in 1..rows_cap {
                for c in 0..cols_cap {
                    let Some(kind) = inferred[c] else {
                        continue;
                    };
                    let cell = &mut grid[r][c];
                    if cell.text.trim().is_empty() {
                        continue;
                    }
                    if let Ok(v) = cell.text.trim().parse::<f64>() {
                        if excel_serial_looks_like_datetime(v) {
                            cell.text = excel_serial_to_iso(v, kind);
                        }
                    }
                }
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
            "warnings": sheet_warnings_for_meta,
            "named_tables": named_tables.iter().map(|t| serde_json::json!({
                "name": t.name.clone(),
                "ref": t.ref_a1.clone(),
                "rows": t.rows,
                "cols": t.cols,
            })).collect::<Vec<_>>()
        }));

        let mut emitted_tables = 0usize;
        let mut group_by_used = false;

        let mut named_bounds: Vec<(usize, usize, usize, usize)> = Vec::new();
        for t in &named_tables {
            let mut runs: Vec<(usize, usize)> = Vec::new();
            let mut cur: Option<(usize, usize)> = None;
            for (si, &actual_r) in row_map.iter().enumerate() {
                if actual_r >= t.row_start && actual_r < t.row_end_excl {
                    match cur.as_mut() {
                        Some((_, e)) if *e == si => *e = si + 1,
                        Some((s, e)) => {
                            runs.push((*s, *e));
                            *s = si;
                            *e = si + 1;
                        }
                        None => cur = Some((si, si + 1)),
                    }
                }
            }
            if let Some((s, e)) = cur {
                runs.push((s, e));
            }

            let runs_len = runs.len();

            if runs.is_empty() {
                continue;
            }

            let c_start = t.col_start.min(used_cols);
            let c_end = t.col_end_excl.min(used_cols);
            if c_end <= c_start {
                continue;
            }

            blocks.push(spreadsheet::named_table_heading(block_index, &t.name));
            block_index += 1;

            for (rs, re) in runs.into_iter() {
                let rows_slice = &grid[rs..re];
                let has_header = t.header_rows > 0 && row_map.get(rs).copied() == Some(t.row_start);
                let (group_by_col, group_label) = limits
                    .group_by
                    .as_deref()
                    .and_then(|spec| {
                        spreadsheet::resolve_group_by_col_and_label(
                            rows_slice, has_header, spec, c_start, c_end,
                        )
                    })
                    .map(|(c, l)| (Some(c), Some(l)))
                    .unwrap_or((None, None));
                if group_by_col.is_some() {
                    group_by_used = true;
                }

                let segs = spreadsheet::semantic_segments(
                    rows_slice,
                    has_header,
                    limits.max_table_rows_per_segment,
                    group_by_col,
                );
                let segs_len = segs.len();

                if runs_len > 1 {
                    let start1 = row_map[rs] + 1;
                    let end1 = row_map[re - 1] + 1;
                    blocks.push(spreadsheet::rows_heading(block_index, start1, end1));
                    block_index += 1;
                }

                for seg in segs {
                    let ds = seg.start;
                    let de = seg.end;
                    let hdr = if has_header {
                        rows_slice.first().map(|r| r.as_slice())
                    } else {
                        None
                    };
                    let data: Vec<&Vec<spreadsheet::CellDisplay>> =
                        rows_slice[ds..de].iter().collect();
                    if runs_len == 1 && segs_len > 1 {
                        let start1 = row_map[rs + ds] + 1;
                        let end1 = row_map[rs + de - 1] + 1;
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
                    blocks.push(spreadsheet::build_table_block_from_parts(
                        block_index,
                        hdr,
                        &data,
                        c_start,
                        c_end,
                        limits.drop_empty_cols,
                    ));
                    block_index += 1;
                    emitted_tables += 1;
                    named_bounds.push((rs + ds, rs + de, c_start, c_end));
                }
            }
        }

        let regions = spreadsheet::split_regions(&grid, limits.empty_row_run);
        let regions_len = regions.len();
        let mut table_idx = 0usize;
        for (rs, re) in regions {
            let region_rows = &grid[rs..re];
            let region_row_map = &row_map[rs..re];
            let Some((cs, ce)) = spreadsheet::region_col_bounds(region_rows) else {
                continue;
            };

            let overlaps_named = named_bounds
                .iter()
                .any(|(r0, r1, c0, c1)| rs < *r1 && re > *r0 && cs < *c1 && ce > *c0);
            if overlaps_named {
                continue;
            }

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
                if regions_len > 1 || col_splits_len > 1 || !named_tables.is_empty() || segs_len > 1
                {
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
    date_kind: Option<DateTimeKind>,
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
                if let Some(k) = date_kind {
                    out.text = excel_serial_to_iso(*f, k);
                } else {
                    out.text = if (f.fract()).abs() < 1e-9 {
                        format!("{}", *f as i64)
                    } else {
                        f.to_string()
                    };
                }
            }
            Data::Int(i) => {
                if let Some(k) = date_kind {
                    out.text = excel_serial_to_iso(*i as f64, k);
                } else {
                    out.text = i.to_string();
                }
            }
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

    out
}

fn xlsx_sheet_name_to_path(bytes: &[u8]) -> Result<std::collections::HashMap<String, String>> {
    use quick_xml::events::Event;

    fn local_name(n: &[u8]) -> &[u8] {
        n.split(|b| *b == b':').last().unwrap_or(n)
    }

    let workbook =
        super::read_zip_file_utf8(bytes, "xl/workbook.xml").context("read workbook.xml")?;
    let rels = super::read_zip_file_utf8(bytes, "xl/_rels/workbook.xml.rels")
        .context("read workbook.xml.rels")?;

    let mut rid_by_name: std::collections::HashMap<String, String> = Default::default();
    {
        let mut r = quick_xml::Reader::from_str(&workbook);
        r.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match r.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"sheet" => {
                    let mut name: Option<String> = None;
                    let mut rid: Option<String> = None;
                    for a in e.attributes().flatten() {
                        let k = local_name(a.key.as_ref());
                        let v = a.unescape_value().unwrap_or_default().to_string();
                        match k {
                            b"name" => name = Some(v),
                            b"id" => rid = Some(v),
                            _ => {}
                        }
                    }
                    if let (Some(n), Some(rid)) = (name, rid) {
                        rid_by_name.insert(n, rid);
                    }
                }
                Ok(Event::Empty(e)) if local_name(e.name().as_ref()) == b"sheet" => {
                    let mut name: Option<String> = None;
                    let mut rid: Option<String> = None;
                    for a in e.attributes().flatten() {
                        let k = local_name(a.key.as_ref());
                        let v = a.unescape_value().unwrap_or_default().to_string();
                        match k {
                            b"name" => name = Some(v),
                            b"id" => rid = Some(v),
                            _ => {}
                        }
                    }
                    if let (Some(n), Some(rid)) = (name, rid) {
                        rid_by_name.insert(n, rid);
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(anyhow::anyhow!(e).context("parse workbook.xml")),
                _ => {}
            }
            buf.clear();
        }
    }

    let mut target_by_rid: std::collections::HashMap<String, String> = Default::default();
    {
        let mut r = quick_xml::Reader::from_str(&rels);
        r.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match r.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"Relationship" => {
                    let mut id: Option<String> = None;
                    let mut target: Option<String> = None;
                    for a in e.attributes().flatten() {
                        let k = a.key.as_ref();
                        let v = a.unescape_value().unwrap_or_default().to_string();
                        match k {
                            b"Id" => id = Some(v),
                            b"Target" => target = Some(v),
                            _ => {}
                        }
                    }
                    if let (Some(id), Some(target)) = (id, target) {
                        target_by_rid.insert(id, target);
                    }
                }
                Ok(Event::Empty(e)) if e.name().as_ref() == b"Relationship" => {
                    let mut id: Option<String> = None;
                    let mut target: Option<String> = None;
                    for a in e.attributes().flatten() {
                        let k = a.key.as_ref();
                        let v = a.unescape_value().unwrap_or_default().to_string();
                        match k {
                            b"Id" => id = Some(v),
                            b"Target" => target = Some(v),
                            _ => {}
                        }
                    }
                    if let (Some(id), Some(target)) = (id, target) {
                        target_by_rid.insert(id, target);
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(anyhow::anyhow!(e).context("parse workbook.xml.rels")),
                _ => {}
            }
            buf.clear();
        }
    }

    let mut out: std::collections::HashMap<String, String> = Default::default();
    for (name, rid) in rid_by_name {
        if let Some(target) = target_by_rid.get(&rid) {
            let target = target.trim_start_matches("/");
            let full = if target.starts_with("xl/") {
                target.to_string()
            } else {
                format!("xl/{target}")
            };
            out.insert(name, full);
        }
    }
    Ok(out)
}

fn xlsx_sheet_formulas(
    bytes: &[u8],
    sheet_path: &str,
) -> Result<std::collections::HashMap<String, String>> {
    use quick_xml::events::Event;

    let xml = super::read_zip_file_utf8(bytes, sheet_path)
        .with_context(|| format!("read {sheet_path}"))?;

    let mut out: std::collections::HashMap<String, String> = Default::default();
    let mut r = quick_xml::Reader::from_str(&xml);
    r.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut cur_cell: Option<String> = None;
    let mut in_f = false;
    let mut f_text = String::new();

    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"c" => {
                cur_cell = None;
                for a in e.attributes().flatten() {
                    if a.key.as_ref() == b"r" {
                        cur_cell = Some(a.unescape_value().unwrap_or_default().to_string());
                    }
                }
            }
            Ok(Event::Start(e)) if e.name().as_ref() == b"f" => {
                in_f = true;
                f_text.clear();
            }
            Ok(Event::End(e)) if e.name().as_ref() == b"f" => {
                in_f = false;
                if let Some(cell) = cur_cell.clone() {
                    let f = f_text.trim();
                    if !f.is_empty() {
                        out.insert(cell, f.to_string());
                    }
                }
            }
            Ok(Event::Text(t)) => {
                if in_f {
                    if let Ok(c) = t.xml_content() {
                        f_text.push_str(&c);
                    } else if let Ok(c) = t.decode() {
                        f_text.push_str(&c);
                    }
                }
            }
            Ok(Event::CData(t)) => {
                if in_f {
                    if let Ok(c) = t.decode() {
                        f_text.push_str(&c);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!(e).context("parse worksheet for formulas")),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn xlsx_date_time_kinds_by_xf(bytes: &[u8]) -> Result<Vec<Option<DateTimeKind>>> {
    use quick_xml::events::Event;

    fn local_name(n: &[u8]) -> &[u8] {
        n.split(|b| *b == b':').last().unwrap_or(n)
    }

    let xml = match super::read_zip_file_utf8(bytes, "xl/styles.xml") {
        Ok(s) => s,
        Err(_) => return Ok(Vec::new()),
    };

    let mut custom: std::collections::HashMap<u32, String> = Default::default();
    let mut xfs: Vec<Option<DateTimeKind>> = Vec::new();

    let mut r = quick_xml::Reader::from_str(&xml);
    r.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut in_numfmts = false;
    let mut in_cellxfs = false;

    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"numFmts" => {
                in_numfmts = true;
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"numFmts" => {
                in_numfmts = false;
            }
            Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"cellXfs" => {
                in_cellxfs = true;
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"cellXfs" => {
                in_cellxfs = false;
            }
            Ok(Event::Empty(e)) if in_numfmts && local_name(e.name().as_ref()) == b"numFmt" => {
                let mut id: Option<u32> = None;
                let mut code: Option<String> = None;
                for a in e.attributes().flatten() {
                    match local_name(a.key.as_ref()) {
                        b"numFmtId" => id = a.unescape_value().ok().and_then(|v| v.parse().ok()),
                        b"formatCode" => {
                            code = Some(a.unescape_value().unwrap_or_default().to_string())
                        }
                        _ => {}
                    }
                }
                if let (Some(id), Some(code)) = (id, code) {
                    custom.insert(id, code);
                }
            }
            Ok(Event::Empty(e)) if in_cellxfs && local_name(e.name().as_ref()) == b"xf" => {
                let mut num_fmt_id: Option<u32> = None;
                for a in e.attributes().flatten() {
                    if local_name(a.key.as_ref()) == b"numFmtId" {
                        num_fmt_id = a.unescape_value().ok().and_then(|v| v.parse().ok());
                    }
                }
                let kind = num_fmt_id
                    .and_then(|id| classify_numfmt(id, custom.get(&id).map(|s| s.as_str())));
                xfs.push(kind);
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!(e).context("parse styles.xml")),
            _ => {}
        }
        buf.clear();
    }

    Ok(xfs)
}

fn classify_numfmt(id: u32, code: Option<&str>) -> Option<DateTimeKind> {
    // Excel built-in date/time formats.
    // This is incomplete but covers the common ones.
    match id {
        14 | 15 | 16 | 17 | 27 | 28 | 29 | 30 | 31 | 36 | 50 | 57 => {
            return Some(DateTimeKind::Date);
        }
        18 | 19 | 20 | 21 | 45 | 46 | 47 => return Some(DateTimeKind::Time),
        22 => return Some(DateTimeKind::DateTime),
        _ => {}
    }

    let code = code?;
    classify_format_code(code)
}

fn classify_format_code(code: &str) -> Option<DateTimeKind> {
    let mut out = String::new();
    let mut in_quote = false;
    let mut in_bracket = 0u32;
    let mut escape = false;

    for ch in code.chars() {
        if escape {
            escape = false;
            continue;
        }
        if ch == '\\' {
            escape = true;
            continue;
        }
        if in_quote {
            if ch == '"' {
                in_quote = false;
            }
            continue;
        }
        if ch == '"' {
            in_quote = true;
            continue;
        }
        if ch == '[' {
            in_bracket += 1;
            continue;
        }
        if ch == ']' {
            in_bracket = in_bracket.saturating_sub(1);
            continue;
        }
        if in_bracket > 0 {
            continue;
        }
        out.push(ch.to_ascii_lowercase());
    }

    let has_y = out.contains('y');
    let has_d = out.contains('d');
    let has_h = out.contains('h');
    let has_s = out.contains('s');
    let has_ampm = out.contains("am/pm");
    let has_time = has_h || has_s || has_ampm;
    let has_date = has_y || has_d || (out.contains('m') && (has_y || has_d));

    match (has_date, has_time) {
        (true, true) => Some(DateTimeKind::DateTime),
        (true, false) => Some(DateTimeKind::Date),
        (false, true) => Some(DateTimeKind::Time),
        _ => None,
    }
}

fn xlsx_sheet_date_time_cells(
    bytes: &[u8],
    sheet_path: &str,
    kinds_by_xf: &[Option<DateTimeKind>],
) -> Result<std::collections::HashMap<String, DateTimeKind>> {
    use quick_xml::events::Event;

    fn local_name(n: &[u8]) -> &[u8] {
        n.split(|b| *b == b':').last().unwrap_or(n)
    }

    let xml = super::read_zip_file_utf8(bytes, sheet_path)
        .with_context(|| format!("read {sheet_path}"))?;

    let mut out: std::collections::HashMap<String, DateTimeKind> = Default::default();
    let mut r = quick_xml::Reader::from_str(&xml);
    r.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if local_name(e.name().as_ref()) == b"c" => {
                let mut cell_ref: Option<String> = None;
                let mut style_idx: Option<usize> = None;
                for a in e.attributes().flatten() {
                    match local_name(a.key.as_ref()) {
                        b"r" => cell_ref = Some(a.unescape_value().unwrap_or_default().to_string()),
                        b"s" => {
                            style_idx = a
                                .unescape_value()
                                .ok()
                                .and_then(|v| v.parse::<usize>().ok())
                        }
                        _ => {}
                    }
                }
                if let (Some(rf), Some(si)) = (cell_ref, style_idx) {
                    if let Some(Some(kind)) = kinds_by_xf.get(si) {
                        out.insert(rf, *kind);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!(e).context("parse worksheet styles")),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn excel_serial_to_iso(serial: f64, kind: DateTimeKind) -> String {
    if !serial.is_finite() {
        return serial.to_string();
    }
    if serial < 0.0 {
        return serial.to_string();
    }

    let mut days = serial.floor() as i64;
    let mut frac = serial - (days as f64);
    if frac < 0.0 {
        frac = 0.0;
    }

    // Excel 1900 leap year bug: serial 60 is the non-existent 1900-02-29.
    if days >= 60 {
        days -= 1;
    }

    // Convert to civil date.
    let base = days_from_civil(1899, 12, 30);
    let z = base + days;
    let (y, m, d) = civil_from_days(z);

    let mut secs = (frac * 86_400.0).round() as i64;
    if secs >= 86_400 {
        secs = 86_399;
    }
    let hh = secs / 3600;
    let mm = (secs % 3600) / 60;
    let ss = secs % 60;

    match kind {
        DateTimeKind::Date => format!("{y:04}-{m:02}-{d:02}"),
        DateTimeKind::Time => format!("{hh:02}:{mm:02}:{ss:02}"),
        DateTimeKind::DateTime => format!("{y:04}-{m:02}-{d:02} {hh:02}:{mm:02}:{ss:02}"),
    }
}

fn excel_serial_looks_like_datetime(serial: f64) -> bool {
    serial.is_finite() && serial >= 10_000.0 && serial <= 100_000.0
}

// Howard Hinnant's civil-from-days algorithm (public domain).
fn days_from_civil(y: i64, m: i64, d: i64) -> i64 {
    let y = y - ((m <= 2) as i64);
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let mp = m + if m > 2 { -3 } else { 9 };
    let doy = (153 * mp + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146097 + doe - 719468
}

fn civil_from_days(z: i64) -> (i64, i64, i64) {
    let z = z + 719468;
    let era = if z >= 0 { z } else { z - 146096 } / 146097;
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = mp + if mp < 10 { 3 } else { -9 };
    let y = y + ((m <= 2) as i64);
    (y, m, d)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formats::finalize;
    use anyhow::Result;
    use std::io::Write;
    use zip::write::FileOptions;

    fn minimal_xlsx_with_style_datetime(header: &str) -> Result<Vec<u8>> {
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(std::io::Cursor::new(&mut buf));
            let opt = FileOptions::default();

            w.start_file("[Content_Types].xml", opt)?;
            w.write_all(
                br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#,
            )?;

            w.add_directory("_rels/", opt)?;
            w.start_file("_rels/.rels", opt)?;
            w.write_all(
                br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
            )?;

            w.add_directory("xl/", opt)?;
            w.add_directory("xl/_rels/", opt)?;
            w.add_directory("xl/worksheets/", opt)?;

            w.start_file("xl/workbook.xml", opt)?;
            w.write_all(
                br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#,
            )?;

            w.start_file("xl/_rels/workbook.xml.rels", opt)?;
            w.write_all(
                br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
            )?;

            w.start_file("xl/styles.xml", opt)?;
            w.write_all(
                br#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>"#,
            )?;

            w.start_file("xl/worksheets/sheet1.xml", opt)?;
            let sheet = format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n  <sheetData>\n    <row r=\"1\">\n      <c r=\"A1\" t=\"inlineStr\"><is><t>{}</t></is></c>\n    </row>\n    <row r=\"2\">\n      <c r=\"A2\" s=\"1\"><v>44128.76593707176</v></c>\n    </row>\n  </sheetData>\n</worksheet>\n",
                header
            );
            w.write_all(sheet.as_bytes())?;

            w.finish()?;
        }
        Ok(buf)
    }

    #[test]
    fn style_xf_and_sheet_cells_are_detected() -> Result<()> {
        let xlsx = minimal_xlsx_with_style_datetime("x")?;
        let map = xlsx_sheet_name_to_path(&xlsx)?;
        assert_eq!(
            map.get("Sheet1").map(|s| s.as_str()),
            Some("xl/worksheets/sheet1.xml")
        );

        let kinds = xlsx_date_time_kinds_by_xf(&xlsx)?;
        assert_eq!(kinds.len(), 2);
        assert_eq!(kinds[1], Some(DateTimeKind::DateTime));

        let sheet_cells = xlsx_sheet_date_time_cells(&xlsx, "xl/worksheets/sheet1.xml", &kinds)?;
        assert_eq!(sheet_cells.get("A2"), Some(&DateTimeKind::DateTime));

        // End-to-end: parse_xlsx should apply the style and render a datetime.
        let parsed = parse_xlsx(&xlsx, crate::spreadsheet::ParseOptions::default())?;
        let doc = finalize(crate::Format::Xlsx, parsed);
        let md = crate::render::to_markdown(&doc);
        assert!(md.contains("2020-10-23"));
        assert!(md.contains("18:22"));
        Ok(())
    }
}

#[derive(Clone, Debug)]
struct NamedTable {
    name: String,
    ref_a1: String,
    row_start: usize,
    row_end_excl: usize,
    col_start: usize,
    col_end_excl: usize,
    header_rows: usize,
    rows: usize,
    cols: usize,
}

fn xlsx_named_tables_for_sheet(bytes: &[u8], sheet_path: &str) -> Result<Vec<NamedTable>> {
    use quick_xml::events::Event;

    fn local_name(n: &[u8]) -> &[u8] {
        n.split(|b| *b == b':').last().unwrap_or(n)
    }

    let sheet_xml = super::read_zip_file_utf8(bytes, sheet_path)
        .with_context(|| format!("read {sheet_path}"))?;

    let rels_path = sheet_rels_path(sheet_path);
    let rels_xml = match super::read_zip_file_utf8(bytes, &rels_path) {
        Ok(s) => s,
        Err(_) => return Ok(Vec::new()),
    };

    let mut target_by_rid: std::collections::HashMap<String, String> = Default::default();
    {
        let mut r = quick_xml::Reader::from_str(&rels_xml);
        r.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match r.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e))
                    if local_name(e.name().as_ref()) == b"Relationship" =>
                {
                    let mut id: Option<String> = None;
                    let mut target: Option<String> = None;
                    for a in e.attributes().flatten() {
                        match local_name(a.key.as_ref()) {
                            b"Id" => id = Some(a.unescape_value().unwrap_or_default().to_string()),
                            b"Target" => {
                                target = Some(a.unescape_value().unwrap_or_default().to_string())
                            }
                            _ => {}
                        }
                    }
                    if let (Some(id), Some(target)) = (id, target) {
                        target_by_rid.insert(id, target);
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(anyhow::anyhow!(e).context("parse sheet rels")),
                _ => {}
            }
            buf.clear();
        }
    }

    let mut table_rids: Vec<String> = Vec::new();
    {
        let mut r = quick_xml::Reader::from_str(&sheet_xml);
        r.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match r.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e))
                    if local_name(e.name().as_ref()) == b"tablePart" =>
                {
                    for a in e.attributes().flatten() {
                        if local_name(a.key.as_ref()) == b"id" {
                            let rid = a.unescape_value().unwrap_or_default().to_string();
                            if !rid.is_empty() {
                                table_rids.push(rid);
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(anyhow::anyhow!(e).context("parse worksheet tableParts")),
                _ => {}
            }
            buf.clear();
        }
    }

    let mut out = Vec::new();
    for rid in table_rids {
        let Some(target) = target_by_rid.get(&rid) else {
            continue;
        };
        let table_path = normalize_xlsx_target(target);
        let table_xml = match super::read_zip_file_utf8(bytes, &table_path) {
            Ok(s) => s,
            Err(_) => continue,
        };
        if let Some(t) = parse_table_xml(&table_xml)? {
            out.push(t);
        }
    }

    Ok(out)
}

fn sheet_rels_path(sheet_path: &str) -> String {
    // xl/worksheets/sheet1.xml -> xl/worksheets/_rels/sheet1.xml.rels
    if let Some((dir, file)) = sheet_path.rsplit_once('/') {
        return format!("{dir}/_rels/{file}.rels");
    }
    format!("_rels/{sheet_path}.rels")
}

fn normalize_xlsx_target(target: &str) -> String {
    let mut t = target.trim_start_matches('/');
    while let Some(rest) = t.strip_prefix("../") {
        t = rest;
    }
    if t.starts_with("xl/") {
        t.to_string()
    } else {
        format!("xl/{t}")
    }
}

fn parse_table_xml(xml: &str) -> Result<Option<NamedTable>> {
    use quick_xml::events::Event;

    fn local_name(n: &[u8]) -> &[u8] {
        n.split(|b| *b == b':').last().unwrap_or(n)
    }

    let mut r = quick_xml::Reader::from_str(xml);
    r.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if local_name(e.name().as_ref()) == b"table" =>
            {
                let mut name: Option<String> = None;
                let mut display_name: Option<String> = None;
                let mut ref_a1: Option<String> = None;
                let mut header_row_count: Option<usize> = None;

                for a in e.attributes().flatten() {
                    let k = local_name(a.key.as_ref());
                    let v = a.unescape_value().unwrap_or_default().to_string();
                    match k {
                        b"name" => name = Some(v),
                        b"displayName" => display_name = Some(v),
                        b"ref" => ref_a1 = Some(v),
                        b"headerRowCount" => header_row_count = v.parse::<usize>().ok(),
                        _ => {}
                    }
                }

                let name = display_name.or(name).unwrap_or_else(|| "Table".to_string());
                let ref_a1 = match ref_a1 {
                    Some(r) if !r.trim().is_empty() => r,
                    _ => return Ok(None),
                };

                let (r0, c0, r1_ex, c1_ex) = match parse_a1_range(&ref_a1) {
                    Some(v) => v,
                    None => return Ok(None),
                };
                let rows = r1_ex.saturating_sub(r0);
                let cols = c1_ex.saturating_sub(c0);
                let header_rows = header_row_count.unwrap_or(1);

                return Ok(Some(NamedTable {
                    name,
                    ref_a1,
                    row_start: r0,
                    row_end_excl: r1_ex,
                    col_start: c0,
                    col_end_excl: c1_ex,
                    header_rows,
                    rows,
                    cols,
                }));
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!(e).context("parse table xml")),
            _ => {}
        }
        buf.clear();
    }
    Ok(None)
}

fn parse_a1_range(s: &str) -> Option<(usize, usize, usize, usize)> {
    let s = s.trim();
    let (a, b) = s.split_once(':').unwrap_or((s, s));
    let (r0, c0) = parse_a1_cell(a)?;
    let (r1, c1) = parse_a1_cell(b)?;
    let r_start = r0.min(r1);
    let c_start = c0.min(c1);
    let r_end_excl = r0.max(r1) + 1;
    let c_end_excl = c0.max(c1) + 1;
    Some((r_start, c_start, r_end_excl, c_end_excl))
}

fn parse_a1_cell(s: &str) -> Option<(usize, usize)> {
    let s = s.trim();
    if s.is_empty() {
        return None;
    }
    let mut letters = String::new();
    let mut digits = String::new();
    for ch in s.chars() {
        if ch.is_ascii_alphabetic() {
            if !digits.is_empty() {
                return None;
            }
            letters.push(ch);
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        } else {
            return None;
        }
    }
    if letters.is_empty() || digits.is_empty() {
        return None;
    }
    let col = excel_col_to_index(&letters)?;
    let row1 = digits.parse::<usize>().ok()?;
    if row1 == 0 {
        return None;
    }
    Some((row1 - 1, col))
}

fn excel_col_to_index(col: &str) -> Option<usize> {
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
