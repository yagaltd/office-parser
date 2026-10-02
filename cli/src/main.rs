use std::collections::HashMap;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result, anyhow};
use clap::{Parser, ValueEnum};

#[derive(Clone, Debug, ValueEnum)]
enum OutputFormat {
    Markdown,
    Json,
}

#[derive(Clone, Debug, Parser)]
#[command(author, version, about)]
struct Args {
    /// Input document path
    input: PathBuf,

    /// Output directory (creates `<out>/asset/*`)
    #[arg(long, default_value = ".")]
    out: PathBuf,

    /// Output format
    #[arg(long, value_enum, default_value_t = OutputFormat::Markdown)]
    format: OutputFormat,

    /// JSON output: omit embedded base64 image bytes and keep image metadata only.
    #[arg(long)]
    json_no_image_bytes: bool,

    /// Chunk size (chars) for Markdown output. Emits the CognitiveOS
    /// chunked view — this output is NOT MorphEditor-dialect markdown.
    /// By default spreadsheets are rendered without extra generic chunking.
    #[arg(long)]
    chunk_size: Option<usize>,

    /// Spreadsheet: max rows per sheet
    #[arg(long)]
    max_rows: Option<usize>,

    /// Spreadsheet: max cols per sheet
    #[arg(long)]
    max_cols: Option<usize>,

    /// Spreadsheet: include last N rows when truncating to `--max-rows`.
    #[arg(long)]
    tail_rows: Option<usize>,

    /// Spreadsheet: max data rows per emitted table segment (semantic splitting).
    #[arg(long)]
    table_rows: Option<usize>,

    /// Spreadsheet: split when this key column changes (e.g. `order_id` or `A`).
    #[arg(long)]
    group_by: Option<String>,

    /// Slide-change direction for MorphEditor deck mode (PPTX/ODP markdown).
    /// Validated against MorphEditor SLIDE_TRANSITIONS; ignored elsewhere.
    #[arg(long)]
    slide_transition: Option<SlideTransition>,
}

#[derive(Clone, Copy, Debug, ValueEnum)]
enum SlideTransition {
    Up,
    Down,
    Left,
    Right,
    Fade,
}

impl SlideTransition {
    fn as_str(self) -> &'static str {
        match self {
            SlideTransition::Up => "up",
            SlideTransition::Down => "down",
            SlideTransition::Left => "left",
            SlideTransition::Right => "right",
            SlideTransition::Fade => "fade",
        }
    }
}

fn main() -> Result<()> {
    let args = Args::parse();
    run(args)
}

fn run(args: Args) -> Result<()> {
    let input_bytes = std::fs::read(&args.input)
        .with_context(|| format!("read input {}", args.input.display()))?;

    let hint = args
        .input
        .file_name()
        .and_then(|s| s.to_str())
        .unwrap_or("input");

    let format = office_parser::Format::from_hint(hint)
        .ok_or_else(|| anyhow!("unsupported input format: {hint}"))?;

    let sheet_default = office_parser::spreadsheet::ParseOptions::default();
    let sheet_opts = office_parser::spreadsheet::ParseOptions {
        max_rows_per_sheet: args.max_rows.unwrap_or(sheet_default.max_rows_per_sheet),
        max_cols_per_sheet: args.max_cols.unwrap_or(sheet_default.max_cols_per_sheet),
        tail_rows: args.tail_rows.unwrap_or(sheet_default.tail_rows),
        max_table_rows_per_segment: args
            .table_rows
            .unwrap_or(sheet_default.max_table_rows_per_segment),
        group_by: args.group_by.clone().or(sheet_default.group_by.clone()),
        ..sheet_default
    };

    let doc = match format {
        office_parser::Format::Xlsx => {
            office_parser::xlsx::parse_with_options(&input_bytes, sheet_opts)?
        }
        office_parser::Format::Ods => {
            office_parser::ods::parse_with_options(&input_bytes, sheet_opts)?
        }
        office_parser::Format::Csv => {
            office_parser::csv::parse_with_options(&input_bytes, sheet_opts)?
        }
        office_parser::Format::Tsv => {
            office_parser::tsv::parse_with_options(&input_bytes, sheet_opts)?
        }
        _ => office_parser::parse_as(&input_bytes, format)?,
    };

    std::fs::create_dir_all(&args.out)
        .with_context(|| format!("create out dir {}", args.out.display()))?;
    let asset_dir = args.out.join("asset");
    std::fs::create_dir_all(&asset_dir)
        .with_context(|| format!("create asset dir {}", asset_dir.display()))?;

    let (doc, id_to_relpath) = write_assets_and_rewrite_filenames(doc, &asset_dir, &args.out)?;

    let stem = args
        .input
        .file_stem()
        .and_then(|s| s.to_str())
        .filter(|s| !s.is_empty())
        .unwrap_or("output");
    let out_path = match args.format {
        OutputFormat::Markdown => args.out.join(format!("{stem}.md")),
        OutputFormat::Json => args.out.join(format!("{stem}.json")),
    };

    match args.format {
        OutputFormat::Markdown => {
            let default_sheet_chunk = matches!(
                doc.metadata.format,
                office_parser::Format::Xlsx
                    | office_parser::Format::Ods
                    | office_parser::Format::Csv
                    | office_parser::Format::Tsv
            );

            // For spreadsheets we already do semantic splitting into multiple Table blocks.
            // Avoid extra chunking by default so row-range headings remain stable.
            let md = if default_sheet_chunk && args.chunk_size.is_none() {
                office_parser::render::to_morph_markdown(&doc)
            } else if let Some(sz) = args.chunk_size {
                let chunks = office_parser::render::to_chunks(&doc, sz);
                chunks
                    .into_iter()
                    .map(|c| c.content)
                    .collect::<Vec<_>>()
                    .join("\n\n<!-- chunk -->\n\n")
            } else {
                office_parser::render::to_morph_markdown(&doc)
            };

            let mut md = md;
            // Replace Mermaid image placeholders with actual asset paths.
            for (id, rel) in &id_to_relpath {
                md = md.replace(&format!("office-image:{id}"), rel);
            }
            let md = format!("{}{}", okf_frontmatter(&doc, args.slide_transition), md);
            std::fs::write(&out_path, md)
                .with_context(|| format!("write {}", out_path.display()))?;
        }
        OutputFormat::Json => {
            let v = office_parser::render::to_json_value_with_options(
                &doc,
                office_parser::render::JsonRenderOptions {
                    include_image_bytes: !args.json_no_image_bytes,
                },
            );
            let s = serde_json::to_string_pretty(&v).context("serialize json")?;
            std::fs::write(&out_path, s)
                .with_context(|| format!("write {}", out_path.display()))?;
        }
    }

    Ok(())
}

fn sanitize_filename_component(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        let ok = ch.is_ascii_alphanumeric() || ch == '.' || ch == '-' || ch == '_' || ch == ' ';
        out.push(if ok { ch } else { '_' });
    }
    let out = out.trim().trim_matches('.').to_string();
    if out.is_empty() {
        "file".to_string()
    } else {
        out
    }
}

fn ext_from_mime(m: &str) -> &'static str {
    match m {
        "image/png" => "png",
        "image/jpeg" => "jpg",
        "image/webp" => "webp",
        "image/bmp" => "bmp",
        "image/tiff" => "tiff",
        _ => "bin",
    }
}

fn basename_only(p: &str) -> &str {
    p.rsplit('/').next().unwrap_or(p)
}

fn write_assets_and_rewrite_filenames(
    mut doc: office_parser::Document,
    asset_dir: &Path,
    out_dir: &Path,
) -> Result<(office_parser::Document, HashMap<String, String>)> {
    let mut id_to_rel: HashMap<String, String> = HashMap::new();
    let mut used: HashMap<String, usize> = HashMap::new();

    for (i, img) in doc.images.iter_mut().enumerate() {
        let suggested = if let Some(src) = img.source_ref.as_deref() {
            let slide_idx = src
                .strip_prefix("slide:")
                .and_then(|s| s.strip_suffix(":snapshot"))
                .and_then(|s| s.parse::<u32>().ok());
            if let Some(idx) = slide_idx {
                Some(format!("slide_{idx:04}.{}", ext_from_mime(&img.mime_type)))
            } else {
                img.filename
                    .as_deref()
                    .map(basename_only)
                    .map(sanitize_filename_component)
            }
        } else {
            img.filename
                .as_deref()
                .map(basename_only)
                .map(sanitize_filename_component)
        };

        let base_name = if let Some(s) = suggested {
            s
        } else {
            let short = img.id.strip_prefix("sha256:").unwrap_or(img.id.as_str());
            let short = &short[..short.len().min(12)];
            format!("image_{i:04}_{short}.{}", ext_from_mime(&img.mime_type))
        };

        let n = used.entry(base_name.clone()).or_insert(0);
        *n += 1;
        let final_name = if *n == 1 {
            base_name
        } else {
            format!("{i:04}_{base_name}")
        };

        let disk_path = asset_dir.join(&final_name);
        std::fs::write(&disk_path, &img.bytes)
            .with_context(|| format!("write asset {}", disk_path.display()))?;

        let rel_s = format!("asset/{final_name}");

        img.filename = Some(rel_s.clone());
        id_to_rel.insert(img.id.clone(), rel_s);
    }

    // Rewrite image blocks to use the `asset/...` path in their `filename` field.
    for b in doc.blocks.iter_mut() {
        if let office_parser::document_ast::Block::Image { id, filename, .. } = b {
            if let Some(rel) = id_to_rel.get(id).cloned() {
                *filename = Some(rel);
            }
        }
    }

    let _ = out_dir;
    Ok((doc, id_to_rel))
}

/// OKF (Open Knowledge Format) frontmatter for Markdown output:
/// required `type`, optional `title`/`description`, format as `tag`.
/// Spec: https://okf.md/spec — view layer only, JSON output is unchanged.
fn okf_frontmatter(
    doc: &office_parser::Document,
    slide_transition: Option<SlideTransition>,
) -> String {
    fn quote(s: &str) -> String {
        let one_line = s.split_whitespace().collect::<Vec<_>>().join(" ");
        format!("\"{}\"", one_line.replace('\\', "\\\\").replace('"', "\\\""))
    }
    let mut fm = String::from("---\ntype: document\n");
    if let Some(t) = doc
        .metadata
        .title
        .as_deref()
        .map(str::trim)
        .filter(|t| !t.is_empty())
    {
        fm.push_str(&format!("title: {}\n", quote(t)));
    }
    // Prefer the first paragraph that differs from the title; fall back to
    // any first heading/paragraph (OKF description should not echo title).
    let title_trim = doc.metadata.title.as_deref().map(str::trim);
    let description = doc
        .blocks
        .iter()
        .filter_map(|b| match b {
            office_parser::document_ast::Block::Paragraph { text, .. } => Some(text.trim()),
            _ => None,
        })
        .find(|t| !t.is_empty() && Some(*t) != title_trim)
        .or_else(|| {
            doc.blocks.iter().find_map(|b| match b {
                office_parser::document_ast::Block::Heading { text, .. }
                | office_parser::document_ast::Block::Paragraph { text, .. } => {
                    let t = text.trim();
                    (!t.is_empty()).then_some(t)
                }
                _ => None,
            })
        });
    if let Some(d) = description {
        fm.push_str(&format!("description: {}\n", quote(d)));
    }
    fm.push_str(&format!("tags: [{}]\n", doc.metadata.format.as_str()));
    if matches!(
        doc.metadata.format,
        office_parser::Format::Pptx | office_parser::Format::Odp
    ) {
        // MorphEditor deck mode: `slides:` in the leading frontmatter fence
        // selects SLIDE_RATIOS deck rendering (absent key -> 16:9 default).
        let ratio = doc
            .metadata
            .extra
            .get("slide_ratio")
            .and_then(|v| v.as_str())
            .unwrap_or("16:9");
        fm.push_str(&format!("slides: {ratio}\n"));
        if let Some(t) = slide_transition {
            fm.push_str(&format!("transition: {}\n", t.as_str()));
        }
    }
    fm.push_str("---\n\n");
    fm
}
