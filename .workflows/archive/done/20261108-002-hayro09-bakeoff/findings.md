# bug-hunter-002 findings — plan 20261108-002 review diff (41c04e7..HEAD)

Verdict: no P0; 1 P1, 4 P2, 5 P3. Upgrade fundamentally sound (budgets,
xref-Prev, OCG /VE, annotation gating, offscreen fills, Cargo wiring clean).

## P1-1 Nested erase pass corrupts renderer state
vendor/hayro/src/lib.rs:231-233 knockout_erase_wanted() returns in_knockout()
with no !self.erasing guard; begin/end clobber erase_groups/erasing. Nested
/Group form inside a knockout element re-enters erase mid-pass; nested
end_knockout_erase sets erasing=false mid-pass -> remaining outer content
paints with Compose::Copy; pops fall through to layers they don't own.
Asymmetry proves omission: mask.rs:81 + cmyk.rs:113 guard || self.erasing.
Fix: `!self.erasing && self.in_knockout()`; optionally save/restore erase_groups.

## P2-1 CMYK ink passes mangle images/patterns
cmyk.rs:17-21 documents a fallback for pattern/image, but only
Color::ink_tints()==None sets unsupported (cmyk.rs:39-41); paint.rs Pattern
branch + image.rs never consult ink -> literal RGB read as inverted tints.
Fix: set ink.unsupported when a non-Color paint/image draws during a pass.

## P2-2 CMYK result mislabels straight RGBA as premultiplied
cmyk.rs:167-175 from_parts with default PixelMetadata (AlphaPremultiplied);
built rgba is straight -> AA edges composite bright. Fix:
PixelMetadata::new(ImageAlphaType::Alpha, true).

## P2-3 Lenient cmap resync swallows later valid sections
parse.rs:221-236: failed high-slot read_u32_code continues, eating
endcodespacerange; subsequent sections consumed as garbage. Fix: return the
consumed token; on end-marker treat as block end. Same class at :251/:284/
:318/:354.

## P2-4 Bake-off pdfium sniff guard is dead
run_bakeoff.py:44-47 includes b"pdfium.dylib"; pdf.rs install_pdfium()
carries "lib/libpdfium.dylib" unconditionally (cfg! is runtime) -> sniff
always true. Fix: keep only FPDF_/pdfium-render markers.

## P3-1 render() pre-clamp checks height for BOTH axes (lib.rs:361-374)
## P3-2 hayro arm ignores clamp_budget (pdf.rs:2714 vs 2764-2773)
## P3-3 glyph stroke offscreen extent mixes spaces (glyph.rs:277-285)
## P3-4 image decode runs before the MAX_IMAGE_PIXELS gate (decode/mod.rs:50-77)
## P3-5 WER order assumption undocumented (README)

## Triage (clean)
Content budgets, xref Prev (visited+depth), decode-limit threading (images
unbroken), OCG /VE (depth 32, arity), annotation gating (Table 171 complete),
offscreen fills (fill⊆control-box), sync poison recovery, glue (RenderCache
per-doc, dpi/72 exact), Cargo unification (exact versions; brotli needed for
CJK), scorer micro-avg/empty-page/padding — all sound.
