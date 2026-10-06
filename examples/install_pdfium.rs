//! Download the Pdfium dynamic library for this platform into the office-parser
//! cache, so `render_pages_png` can use the pdfium backend (feature
//! `render-pdfium`).
//!
//!   cargo run --features render-pdfium --example install_pdfium

fn main() {
    match office_parser::formats::pdf::install_pdfium() {
        Ok(path) => println!("pdfium installed: {}", path.display()),
        Err(e) => {
            eprintln!("error: {e}");
            std::process::exit(1);
        }
    }
}
