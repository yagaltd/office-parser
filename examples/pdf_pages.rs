//! The PDF → pages pipeline: page text + embedded images from the parser,
//! page PNGs from the renderer, ready for OCR/vision downstream.
//!
//!   cargo run --features pdf,render-pdfium --example pdf_pages -- <doc.pdf> <out-dir>
//!   OFFICE_PARSER_BACKEND=pdfium   # force the pdfium leg (default: poppler)

fn main() {
    let mut args = std::env::args().skip(1);
    let pdf_path = args.next().expect("usage: pdf_pages <doc.pdf> <out-dir>");
    let out_dir = args.next().expect("usage: pdf_pages <doc.pdf> <out-dir>");
    std::fs::create_dir_all(&out_dir).expect("out dir");

    let bytes = std::fs::read(&pdf_path).expect("pdf read");

    // ── parse: text layer + embedded images ─────────────────────────────
    let document = office_parser::formats::pdf::parse(&bytes).expect("parse");
    let text: String = document
        .blocks
        .iter()
        .filter_map(|b| match b {
            office_parser::document_ast::Block::Heading { text, .. }
            | office_parser::document_ast::Block::Paragraph { text, .. }
            | office_parser::document_ast::Block::Note { text, .. } => Some(text.clone()),
            office_parser::document_ast::Block::List { items, .. } => Some(
                items
                    .iter()
                    .map(|i| i.text.clone())
                    .collect::<Vec<_>>()
                    .join("\n"),
            ),
            office_parser::document_ast::Block::Table { rows, .. } => Some(
                rows.iter()
                    .map(|row| {
                        row.iter()
                            .map(|c| c.text.as_str())
                            .collect::<Vec<_>>()
                            .join(" | ")
                    })
                    .collect::<Vec<_>>()
                    .join("\n"),
            ),
            office_parser::document_ast::Block::Image { filename, .. } => {
                Some(format!("[image: {}]", filename.clone().unwrap_or_default()))
            }
            office_parser::document_ast::Block::Link { url, .. } => Some(url.clone()),
        })
        .collect::<Vec<_>>()
        .join("\n");
    println!(
        "parsed: {} text chars, {} embedded images",
        text.chars().count(),
        document.images.len()
    );
    for (_i, image) in document.images.iter().enumerate() {
        let name = format!(
            "{out_dir}/embedded-{}.{}",
            image.id,
            image.mime_type.rsplit('/').next().unwrap_or("bin")
        );
        std::fs::write(&name, &image.bytes).expect("embedded image write");
        println!("  embedded image → {name}");
    }
    let text_path = format!("{out_dir}/text.txt");
    std::fs::write(&text_path, &text).expect("text write");
    println!("  text → {text_path}");

    // ── render: one PNG per page (pdfium or poppler) ────────────────────
    let pages = office_parser::formats::pdf::render_pages_png(&bytes, 150).expect("render");
    for (i, png) in pages.iter().enumerate() {
        let name = format!("{out_dir}/page-{}.png", i + 1);
        std::fs::write(&name, png).expect("page write");
        println!("  page {} → {name} ({} KB)", i + 1, png.len() / 1024);
    }
    println!("done: {} pages", pages.len());
}
