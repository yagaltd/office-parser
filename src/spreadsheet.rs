#[derive(Clone, Debug)]
pub struct ParseOptions {
    pub max_sheets: usize,
    pub max_rows_per_sheet: usize,
    pub max_cols_per_sheet: usize,
    pub max_cell_chars: usize,

    /// When truncating rows, include the last N rows as a tail sample.
    ///
    /// Example: if `max_rows_per_sheet=20_000` and `tail_rows=2_000`, we emit
    /// the first 18k rows and the last 2k rows.
    pub tail_rows: usize,

    /// A region ends after N consecutive empty rows.
    pub empty_row_run: usize,

    /// Drop columns that are completely empty within a region/table.
    pub drop_empty_cols: bool,

    /// Maximum number of data rows to include in a single emitted markdown table block.
    ///
    /// This is a semantic splitter (separate from `render::to_chunks`), useful when a single
    /// contiguous sheet region is extremely large.
    pub max_table_rows_per_segment: usize,

    /// If set, split table segments when this key column changes.
    ///
    /// Accepted values:
    /// - Excel column letters: `A`, `B`, `AA`, ...
    /// - Header text (case-insensitive) when a header row is detected.
    pub group_by: Option<String>,
}

impl Default for ParseOptions {
    fn default() -> Self {
        Self {
            max_sheets: 32,
            max_rows_per_sheet: 20_000,
            max_cols_per_sheet: 100,
            max_cell_chars: 2_000,
            tail_rows: 2_000,
            empty_row_run: 1,
            drop_empty_cols: true,
            max_table_rows_per_segment: 500,
            group_by: None,
        }
    }
}
