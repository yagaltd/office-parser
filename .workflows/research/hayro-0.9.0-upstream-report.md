# hayro upstream 0.9.0 vs vendored PdfCraft stack (scout report, 2026-11-09)

Refs: vendored hayro 0.7.1 / syntax 0.7.2 / interpret 0.7.0 / cmap 0.1.0 /
jbig2 0.3.0 / postscript 0.1.0 / vello 0.0.8 (572 PdfCraft patch marks);
upstream main @ 360f4a2 (0.9.0; new crates hayro-font/ccitt/jpeg2000/svg/write,
vello_cpu → crates.io 0.3).

## Wiring discrepancy (report-only per AGENTS.md containment)
[patch.crates-io] wires ONLY hayro/syntax/interpret/cmap — the vendored
hayro-jbig2, hayro-postscript, vello_common, vello_cpu copies are NOT in the
build (registry versions are used). Those vendor dirs are dead weight; their
patch marks do not apply today.

## Patch families still needed upstream (10)
Prev-as-indirect-ref; sync poison recovery; run-length decode limit;
lenient cmap (upstream 0.1.1 is stricter — we skip malformed entries);
renderer DoS bounds (no pixel/allocation caps); offscreen-path skip;
knockout/isolated groups + CMYK group passes (upstream Device IGNORES the
flags — biggest fidelity gap); annotation flags; OCG /VE; content byte
budgets (upstream has only a depth cap).

## Fixed/superseded upstream
xref hardening (u32 ids, prev-chain cycles), reader nesting caps, dash
hardening, image resampling color-bleed (#1507), type3 unicode mapping,
function fixes, tiling/soft-mask/shading families, JBIG2 memory budget
(0.4.0; default 1 GiB vs our 32 MB pixel cap), 0.8 backports by definition.

## Upstream-only wins 0.7→0.9
Crypt filter + passwords, text hinting, glyph caching, aliasing_threshold,
warning_sink, JBIG2 strict mode/Globals, ~40 fidelity PRs.

## API-port risk: MEDIUM
Glue = one ~35-line fn (src/formats/pdf.rs:2714). Survives: re-exports,
Pdf::new, pages(), render_dimensions(), RenderCache. Breaks: render_size()
removed; RenderSettings → PixmapSettings; render_into takes (&mut
RenderContext, Affine) via Pixmap; logging/simd features removed; MSRV 1.92.

## Recommendation
STAY VENDED for OCR rendering (the load-bearing patches — content budgets,
allocation caps, lenient cmap, Prev-as-ref — are exactly what upstream still
lacks). Upgrade = cherry-pick project: rewrite glue, port the 10 families
onto the restructured tree, decide the unwired vendor copies. Benchmarks
(dense 10-page A/B, 2026-11-09): hayro +1.14% chars vs pdfium, ~5% faster —
supports the pure-Rust direction pending a real-corpus WER bake-off.
