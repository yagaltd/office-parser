use anyhow::{Context, Result};

use crate::formats::{ParsedOfficeDocument, spreadsheet};

pub fn parse(bytes: &[u8]) -> crate::Result<crate::Document> {
    let parsed = parse_delimited(
        bytes,
        b',',
        crate::Format::Csv,
        crate::spreadsheet::ParseOptions::default(),
    )
    .map_err(|e| crate::Error::from_parse(crate::Format::Csv, e))?;
    Ok(super::finalize(crate::Format::Csv, parsed))
}

pub fn parse_with_options(
    bytes: &[u8],
    opts: crate::spreadsheet::ParseOptions,
) -> crate::Result<crate::Document> {
    let parsed = parse_delimited(bytes, b',', crate::Format::Csv, opts)
        .map_err(|e| crate::Error::from_parse(crate::Format::Csv, e))?;
    Ok(super::finalize(crate::Format::Csv, parsed))
}

pub(crate) fn parse_delimited(
    bytes: &[u8],
    delim: u8,
    fmt: crate::Format,
    limits: spreadsheet::Limits,
) -> Result<ParsedOfficeDocument> {
    use csv::ReaderBuilder;

    let mut r = ReaderBuilder::new()
        .has_headers(false)
        .delimiter(delim)
        .from_reader(bytes);

    let mut rows: Vec<Vec<spreadsheet::CellDisplay>> = Vec::new();
    let mut row_map: Vec<usize> = Vec::new();
    let mut tail: std::collections::VecDeque<Vec<spreadsheet::CellDisplay>> = Default::default();
    let mut tail_map: std::collections::VecDeque<usize> = Default::default();
    let mut total_rows: usize = 0;
    let mut truncated = false;

    let tail_rows = limits.tail_rows.min(limits.max_rows_per_sheet / 2);
    let head_rows = limits.max_rows_per_sheet.saturating_sub(tail_rows);

    for rec in r.records() {
        total_rows += 1;
        let actual_r0 = total_rows - 1;
        let rec = rec.context("read csv record")?;
        let mut out_row: Vec<spreadsheet::CellDisplay> = Vec::new();
        for (ci, field) in rec.iter().enumerate() {
            if ci >= limits.max_cols_per_sheet {
                break;
            }
            let t = spreadsheet::normalize_cell_text(field.to_string(), limits.max_cell_chars);
            out_row.push(spreadsheet::CellDisplay {
                text: t,
                ..Default::default()
            });
        }
        if head_rows == limits.max_rows_per_sheet {
            // Head-only.
            if rows.len() < limits.max_rows_per_sheet {
                rows.push(out_row);
                row_map.push(actual_r0);
            } else {
                truncated = true;
            }
        } else {
            // Head+tail.
            if rows.len() < head_rows {
                rows.push(out_row);
                row_map.push(actual_r0);
            } else {
                truncated = true;
                if tail_rows > 0 {
                    if tail.len() == tail_rows {
                        tail.pop_front();
                        tail_map.pop_front();
                    }
                    tail.push_back(out_row);
                    tail_map.push_back(actual_r0);
                }
            }
        }
    }

    if !tail.is_empty() {
        rows.extend(tail.into_iter());
        row_map.extend(tail_map.into_iter());
    }

    let max_cols = rows.iter().map(|r| r.len()).max().unwrap_or(0);
    for r in &mut rows {
        r.resize(max_cols, spreadsheet::CellDisplay::default());
    }

    let mut blocks = Vec::new();
    let mut block_index = 0usize;
    blocks.push(spreadsheet::sheet_heading(block_index, "Sheet1"));
    block_index += 1;

    let mut warnings: Vec<String> = Vec::new();
    if truncated {
        if tail_rows > 0 && head_rows < limits.max_rows_per_sheet {
            warnings.push(format!(
                "rows truncated: {} -> {} (head {} + tail {})",
                total_rows,
                rows.len(),
                head_rows,
                tail_rows
            ));
        } else {
            warnings.push(format!("rows truncated: {} -> {}", total_rows, rows.len()));
        }
    }
    if max_cols >= limits.max_cols_per_sheet {
        warnings.push(format!(
            "cols truncated: {} (limit {})",
            max_cols, limits.max_cols_per_sheet
        ));
    }
    if !warnings.is_empty() {
        blocks.push(spreadsheet::warning_paragraph(
            block_index,
            &format!("warning: {}", warnings.join("; ")),
        ));
        block_index += 1;
    }

    if !rows.is_empty() {
        let has_header = spreadsheet::first_row_looks_like_header(&rows);
        blocks.push(spreadsheet::table_heading(block_index, 1));
        block_index += 1;

        let mut group_by_used = false;
        let (group_by_col, group_label) = limits
            .group_by
            .as_deref()
            .and_then(|spec| {
                spreadsheet::resolve_group_by_col_and_label(&rows, has_header, spec, 0, max_cols)
            })
            .map(|(c, l)| (Some(c), Some(l)))
            .unwrap_or((None, None));
        if group_by_col.is_some() {
            group_by_used = true;
        }

        let segs = spreadsheet::semantic_segments(
            &rows,
            has_header,
            limits.max_table_rows_per_segment,
            group_by_col,
        );
        let segs_len = segs.len();

        for seg in segs {
            let ds = seg.start;
            let de = seg.end;
            if segs_len > 1 {
                let start1 = row_map.get(ds).copied().unwrap_or(0) + 1;
                let end1 = row_map.get(de - 1).copied().unwrap_or(0) + 1;
                if let (Some(label), Some(key)) = (group_label.as_deref(), seg.group_key.as_deref())
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
                rows.first().map(|r| r.as_slice())
            } else {
                None
            };
            let data: Vec<&Vec<spreadsheet::CellDisplay>> = rows[ds..de].iter().collect();
            blocks.push(spreadsheet::build_table_block_from_parts(
                block_index,
                hdr,
                &data,
                0,
                max_cols,
                limits.drop_empty_cols,
            ));
            block_index += 1;
        }

        if limits.group_by.is_some() && !group_by_used {
            warnings.push(format!(
                "group_by '{}' did not match any column (try a column letter like 'A' or a header name)",
                limits.group_by.as_deref().unwrap_or("")
            ));
        }
    }

    Ok(ParsedOfficeDocument {
        blocks,
        images: Vec::new(),
        metadata_json: {
            let mut v = serde_json::json!({
                "format": fmt.as_str(),
                "delimiter": delim,
                "rows": rows.len(),
                "cols": max_cols,
            });
            if !warnings.is_empty() {
                v["warnings"] = serde_json::json!(warnings);
            }
            v
        },
    })
}
