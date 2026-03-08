use anyhow::Result;

use crate::formats::ParsedOfficeDocument;

pub fn parse(bytes: &[u8]) -> crate::Result<crate::Document> {
    let parsed = parse_tsv(bytes).map_err(|e| crate::Error::from_parse(crate::Format::Tsv, e))?;
    Ok(super::finalize(crate::Format::Tsv, parsed))
}

pub fn parse_with_options(
    bytes: &[u8],
    opts: crate::spreadsheet::ParseOptions,
) -> crate::Result<crate::Document> {
    let parsed = crate::formats::csv::parse_delimited(bytes, b'\t', crate::Format::Tsv, opts)
        .map_err(|e| crate::Error::from_parse(crate::Format::Tsv, e))?;
    Ok(super::finalize(crate::Format::Tsv, parsed))
}

fn parse_tsv(bytes: &[u8]) -> Result<ParsedOfficeDocument> {
    crate::formats::csv::parse_delimited(
        bytes,
        b'\t',
        crate::Format::Tsv,
        crate::spreadsheet::ParseOptions::default(),
    )
}
