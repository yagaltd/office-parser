mod document;
pub mod document_ast;
pub mod formats;
pub mod render;
pub mod spreadsheet;

pub use document::{Document, DocumentMetadata, ExtractedImage, Format};

#[derive(thiserror::Error, Debug)]
pub enum Error {
    #[error("unsupported format hint: {0}")]
    UnsupportedFormat(String),

    #[error("feature not enabled for format: {0}")]
    FeatureDisabled(&'static str),

    #[error("I/O error")]
    Io(#[from] std::io::Error),

    #[error("zip error")]
    Zip(#[source] anyhow::Error),

    #[error("xml parse error")]
    Xml(#[source] anyhow::Error),

    #[error("pdf parse error")]
    Pdf(#[source] anyhow::Error),

    #[error("spreadsheet parse error")]
    Spreadsheet(#[source] anyhow::Error),

    #[error("render error")]
    Render(#[source] anyhow::Error),

    #[error(transparent)]
    Other(#[from] anyhow::Error),
}

impl Error {
    pub(crate) fn from_parse(format: crate::Format, err: anyhow::Error) -> Self {
        match format {
            crate::Format::Pdf => Self::Pdf(err),
            crate::Format::Docx
            | crate::Format::Odt
            | crate::Format::Pptx
            | crate::Format::Odp
            | crate::Format::Epub
            | crate::Format::Xmind
            | crate::Format::Mmap => {
                let msg = err.to_string();
                if msg.contains("zip") || msg.contains("Zip") {
                    Self::Zip(err)
                } else {
                    Self::Xml(err)
                }
            }
            crate::Format::Xlsx | crate::Format::Ods | crate::Format::Csv | crate::Format::Tsv => {
                Self::Spreadsheet(err)
            }
            crate::Format::Xml => Self::Xml(err),
            crate::Format::Rtf
            | crate::Format::Json
            | crate::Format::Yaml
            | crate::Format::Toml => Self::Other(err),
        }
    }
}

pub type Result<T> = std::result::Result<T, Error>;

/// Parse using a MIME type or extension hint (e.g. `application/pdf` or `report.pdf`).
pub fn parse(bytes: &[u8], hint: &str) -> Result<Document> {
    let fmt = Format::from_hint(hint).ok_or_else(|| Error::UnsupportedFormat(hint.to_string()))?;
    parse_as(bytes, fmt)
}

pub fn parse_as(bytes: &[u8], format: Format) -> Result<Document> {
    match format {
        Format::Docx => docx::parse(bytes),
        Format::Odt => odt::parse(bytes),
        Format::Pptx => pptx::parse(bytes),
        Format::Odp => odp::parse(bytes),
        Format::Xlsx => xlsx::parse(bytes),
        Format::Ods => ods::parse(bytes),
        Format::Csv => csv::parse(bytes),
        Format::Tsv => tsv::parse(bytes),
        Format::Pdf => pdf::parse(bytes),
        Format::Rtf => rtf::parse(bytes),
        Format::Json => json::parse(bytes),
        Format::Yaml => yaml::parse(bytes),
        Format::Toml => toml::parse(bytes),
        Format::Xml => xml::parse(bytes),
        Format::Epub => epub::parse(bytes),
        Format::Xmind => xmind::parse(bytes),
        Format::Mmap => mmap::parse(bytes),
    }
}

pub mod docx {
    use super::{Document, Result};
    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::docx::parse(bytes)
    }
}

pub mod odt {
    use super::{Document, Result};
    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::odt::parse(bytes)
    }
}

pub mod rtf {
    use super::{Document, Result};
    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::rtf::parse(bytes)
    }
}

pub mod pptx {
    use super::{Document, Result};

    #[derive(Clone, Copy, Debug)]
    pub struct ParseOptions {
        /// Reserved for future snapshot support.
        ///
        /// Current behavior: PPTX parsing emits semantic blocks and extracted embedded images,
        /// but does not generate slide snapshot images.
        pub include_slide_snapshots: bool,
    }

    impl Default for ParseOptions {
        fn default() -> Self {
            Self {
                include_slide_snapshots: false,
            }
        }
    }

    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::pptx::parse(bytes)
    }

    pub fn parse_with_options(bytes: &[u8], opts: ParseOptions) -> Result<Document> {
        crate::formats::pptx::parse_with_options(bytes, opts)
    }
}

pub mod odp {
    use super::{Document, Result};

    #[derive(Clone, Copy, Debug)]
    pub struct ParseOptions {
        /// Reserved for future snapshot support.
        ///
        /// Current behavior: ODP parsing emits semantic blocks and extracted embedded images,
        /// but does not generate slide snapshot images.
        pub include_slide_snapshots: bool,
    }

    impl Default for ParseOptions {
        fn default() -> Self {
            Self {
                include_slide_snapshots: false,
            }
        }
    }

    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::odp::parse(bytes)
    }

    pub fn parse_with_options(bytes: &[u8], opts: ParseOptions) -> Result<Document> {
        crate::formats::odp::parse_with_options(bytes, opts)
    }
}

pub mod xlsx {
    use super::{Document, Result};

    pub type ParseOptions = crate::spreadsheet::ParseOptions;

    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::xlsx::parse(bytes)
    }

    pub fn parse_with_options(bytes: &[u8], opts: ParseOptions) -> Result<Document> {
        crate::formats::xlsx::parse_with_options(bytes, opts)
    }
}

pub mod ods {
    use super::{Document, Result};

    pub type ParseOptions = crate::spreadsheet::ParseOptions;

    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::ods_sheet::parse(bytes)
    }

    pub fn parse_with_options(bytes: &[u8], opts: ParseOptions) -> Result<Document> {
        crate::formats::ods_sheet::parse_with_options(bytes, opts)
    }
}

pub mod csv {
    use super::{Document, Result};

    pub type ParseOptions = crate::spreadsheet::ParseOptions;

    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::csv::parse(bytes)
    }

    pub fn parse_with_options(bytes: &[u8], opts: ParseOptions) -> Result<Document> {
        crate::formats::csv::parse_with_options(bytes, opts)
    }
}

pub mod tsv {
    use super::{Document, Result};

    pub type ParseOptions = crate::spreadsheet::ParseOptions;

    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::tsv::parse(bytes)
    }

    pub fn parse_with_options(bytes: &[u8], opts: ParseOptions) -> Result<Document> {
        crate::formats::tsv::parse_with_options(bytes, opts)
    }
}

pub mod pdf {
    use super::{Document, Result};
    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::pdf::parse(bytes)
    }
}

pub mod json {
    use super::{Document, Result};
    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::json::parse(bytes)
    }
}

pub mod yaml {
    use super::{Document, Result};
    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::yaml::parse(bytes)
    }
}

pub mod toml {
    use super::{Document, Result};
    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::toml::parse(bytes)
    }
}

pub mod xml {
    use super::{Document, Result};
    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::xml::parse(bytes)
    }
}

pub mod epub {
    use super::{Document, Result};
    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::epub::parse(bytes)
    }
}

pub mod xmind {
    use super::{Document, Result};
    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::xmind::parse(bytes)
    }
}

pub mod mmap {
    use super::{Document, Result};
    pub fn parse(bytes: &[u8]) -> Result<Document> {
        crate::formats::mmap::parse(bytes)
    }
}
