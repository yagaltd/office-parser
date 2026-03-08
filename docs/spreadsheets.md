# Spreadsheets

## ParseOptions

`office_parser::spreadsheet::ParseOptions` controls safety and output segmentation:

- `max_sheets`
- `max_rows_per_sheet`
- `max_cols_per_sheet`
- `max_cell_chars`
- `tail_rows`
- `empty_row_run`
- `drop_empty_cols`
- `max_table_rows_per_segment`
- `group_by`

## Truncation model

When row limits are exceeded:

- keep head rows
- optionally include `tail_rows` from the end

This avoids pathological outputs while preserving representative data.

## Segmentation model

Tables are split into multiple `Block::Table` segments:

- by `max_table_rows_per_segment`
- optionally by `group_by` key change

`group_by` accepts:

- header names (case-insensitive)
- Excel column letters (`A`, `AA`, ...)

## Date/time formatting

- XLSX uses style-aware formatting when style metadata is present.
- A fallback heuristic can format serial values when headers imply date/time semantics.
