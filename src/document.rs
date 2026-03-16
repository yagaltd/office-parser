use crate::document_ast::Block;

#[derive(Clone, Debug)]
pub struct Document {
    pub blocks: Vec<Block>,
    pub images: Vec<ExtractedImage>,
    pub metadata: DocumentMetadata,
}

#[derive(Clone, Debug)]
pub struct ExtractedImage {
    pub bytes: Vec<u8>,
    pub mime_type: String,
    pub filename: Option<String>,
    pub source_ref: Option<String>,
    pub id: String,
}

#[derive(Clone, Debug)]
pub struct DocumentMetadata {
    pub format: Format,
    pub title: Option<String>,
    pub page_count: Option<usize>,
    pub slide_count: Option<usize>,
    pub extra: serde_json::Value,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Format {
    Docx,
    Odt,
    Pptx,
    Odp,
    Xlsx,
    Ods,
    Csv,
    Tsv,
    Pdf,
    Rtf,
    Json,
    Yaml,
    Toml,
    Xml,
    Epub,
    Xmind,
    Mmap,
}

impl Format {
    pub const fn as_str(&self) -> &'static str {
        match self {
            Self::Docx => "docx",
            Self::Odt => "odt",
            Self::Pptx => "pptx",
            Self::Odp => "odp",
            Self::Xlsx => "xlsx",
            Self::Ods => "ods",
            Self::Csv => "csv",
            Self::Tsv => "tsv",
            Self::Pdf => "pdf",
            Self::Rtf => "rtf",
            Self::Json => "json",
            Self::Yaml => "yaml",
            Self::Toml => "toml",
            Self::Xml => "xml",
            Self::Epub => "epub",
            Self::Xmind => "xmind",
            Self::Mmap => "mmap",
        }
    }

    pub fn from_hint(hint: &str) -> Option<Self> {
        let h = hint.trim().to_ascii_lowercase();

        // MIME
        match h.as_str() {
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document" => {
                return Some(Self::Docx);
            }
            "application/vnd.oasis.opendocument.text" => return Some(Self::Odt),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation" => {
                return Some(Self::Pptx);
            }
            "application/vnd.oasis.opendocument.presentation" => return Some(Self::Odp),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" => {
                return Some(Self::Xlsx);
            }
            "application/vnd.oasis.opendocument.spreadsheet" => return Some(Self::Ods),
            "text/csv" => return Some(Self::Csv),
            "text/tab-separated-values" => return Some(Self::Tsv),
            "application/pdf" => return Some(Self::Pdf),
            "text/rtf" | "application/rtf" => return Some(Self::Rtf),
            "application/json" => return Some(Self::Json),
            "application/yaml" | "application/x-yaml" | "text/yaml" | "text/x-yaml" => {
                return Some(Self::Yaml);
            }
            "application/toml" | "application/x-toml" | "text/toml" | "text/x-toml" => {
                return Some(Self::Toml);
            }
            "application/xml" | "text/xml" | "application/rss+xml" => return Some(Self::Xml),
            "application/epub+zip" => return Some(Self::Epub),
            "application/vnd.xmind.workbook" => return Some(Self::Xmind),
            "application/x-mindmanager" | "application/vnd.mindjet.mindmanager" => {
                return Some(Self::Mmap);
            }
            _ => {}
        }

        // Extension/path
        let ext = h
            .rsplit('.')
            .next()
            .unwrap_or("")
            .trim_matches(|c: char| !c.is_ascii_alphanumeric());
        match ext {
            "docx" => Some(Self::Docx),
            "odt" => Some(Self::Odt),
            "pptx" => Some(Self::Pptx),
            "odp" => Some(Self::Odp),
            "xlsx" => Some(Self::Xlsx),
            "ods" => Some(Self::Ods),
            "csv" => Some(Self::Csv),
            "tsv" => Some(Self::Tsv),
            "pdf" => Some(Self::Pdf),
            "rtf" => Some(Self::Rtf),
            "json" => Some(Self::Json),
            "yaml" | "yml" => Some(Self::Yaml),
            "toml" => Some(Self::Toml),
            "xml" => Some(Self::Xml),
            "epub" => Some(Self::Epub),
            "xmind" => Some(Self::Xmind),
            "mmap" => Some(Self::Mmap),
            _ => None,
        }
    }
}
