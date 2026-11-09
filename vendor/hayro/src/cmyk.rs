//! PdfCraft patch: the ink passes of an isolated transparency group that blends in
//! `DeviceCMYK` (ISO 32000-2 §11.3.4; see `hayro_interpret::CmykGroup`).

use crate::Renderer;
use kurbo::{Affine, Rect};
use std::sync::Arc;
use vello_cpu::peniko::ImageQuality;
use vello_cpu::Pixmap;

/// PdfCraft patch: the largest CMYK group drawn in ink passes, in pixels: the passes render on
/// canvases of their own, and a larger group is drawn as before.
const MAX_CMYK_GROUP_PIXELS: u64 = 1 << 24;

/// PdfCraft patch (27): set while this renderer draws one ink pass of a group blending in
/// `DeviceCMYK`.
#[derive(Clone, Copy, Default)]
pub(crate) struct InkState {
    pass: Option<InkPass>,
    /// Something in the group has no ink tints (an RGB colour, a pattern, an image): the
    /// group is drawn as before instead.
    pub(crate) unsupported: bool,
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum InkPass {
    /// Cyan, magenta and yellow as red, green and blue.
    Cmy,
    /// Black as grey.
    K,
}

impl InkState {
    /// PdfCraft patch (bug-hunter-002 P2-1): is an ink pass live? (image.rs
    /// needs the answer without touching the private field.)
    pub(crate) fn is_passing(&self) -> bool {
        self.pass.is_some()
    }

    /// The colour a pass paints for `color`; transparent, with the group marked unsupported,
    /// when the colour has no ink tints.
    pub(crate) fn paint(
        &mut self,
        color: &hayro_interpret::color::Color,
    ) -> hayro_interpret::color::AlphaColor {
        let Some((tints, alpha)) = color.ink_tints() else {
            self.unsupported = true;
            return hayro_interpret::color::AlphaColor::TRANSPARENT;
        };
        match self.pass {
            Some(InkPass::Cmy) => hayro_interpret::color::AlphaColor::new([
                1.0 - tints[0],
                1.0 - tints[1],
                1.0 - tints[2],
                alpha,
            ]),
            Some(InkPass::K) => hayro_interpret::color::AlphaColor::new([
                1.0 - tints[3],
                1.0 - tints[3],
                1.0 - tints[3],
                alpha,
            ]),
            // PdfCraft patch: no pass under way; the interpreter colour as RGBA.
            None => color.to_rgba(),
        }
    }
}

impl Renderer<'_> {
    /// PdfCraft patch (27): one ink pass of `group` on a `width` × `height` canvas whose corner
    /// is device point `offset`; `None` when the group holds something the pass has no tints
    /// for.
    pub(crate) fn ink_pass(
        &self,
        group: &hayro_interpret::CmykGroup<'_, '_>,
        pass: InkPass,
        offset: (f64, f64),
        width: u16,
        height: u16,
    ) -> Option<Pixmap> {
        let mut ctx = self.child_context(width, height);
        let mut renderer = Renderer {
            global: self.global,
            ctx: &mut ctx,
            inside_pattern: self.inside_pattern,
            soft_mask_cache: Default::default(),
            in_type3_glyph: false,
            knockout: Vec::new(),
            erasing: false,
            erase_groups: 0,
            ink: InkState {
                pass: Some(pass),
                unsupported: false,
            },
        };
        // The pass draws inside a layer of its own, so it composites over the transparent
        // canvas instead of blending with whatever it covers.
        renderer.ctx.push_layer(None, None, None, None, None);
        group.interpret(&mut renderer, offset);
        renderer.ctx.pop_layer();
        if renderer.ink.unsupported {
            return None;
        }
        let mut pix = Pixmap::new(width, height);
        renderer.ctx.flush();
        let mut resources = vello_cpu::Resources::default();
        renderer.ctx.render(&mut pix, &mut resources);
        Some(pix)
    }

    /// PdfCraft patch (27): an isolated group that blends in `DeviceCMYK`. Its contents are
    /// drawn twice on canvases of their own, as cyan-magenta-yellow and as black, and the tints
    /// those passes leave are converted to sRGB the way the device draws solid CMYK colour. The
    /// result goes into the layer the interpreter pushed for the group, which composites it
    /// with the group's opacity, soft mask and blend mode as before.
    pub(crate) fn draw_isolated_cmyk_group(
        &mut self,
        group: &hayro_interpret::CmykGroup<'_, '_>,
    ) -> bool {
        if self.ink.pass.is_some() || self.erasing {
            return false;
        }
        let canvas = Rect::new(
            0.0,
            0.0,
            f64::from(self.ctx.width()),
            f64::from(self.ctx.height()),
        );
        let area = group.device_bbox().intersect(canvas);
        if !(area.width() > 0.0 && area.height() > 0.0) {
            // Nothing of the group is on the canvas (a NaN box lands here too).
            return true;
        }
        let (x0, y0) = (area.x0.floor(), area.y0.floor());
        // Within the canvas, so the sides fit its u16 ones.
        let width = (area.x1.ceil() - x0) as u16;
        let height = (area.y1.ceil() - y0) as u16;
        if width == 0 || height == 0 {
            return true;
        }
        if u64::from(width) * u64::from(height) > MAX_CMYK_GROUP_PIXELS {
            return false;
        }

        let Some(cmy) = self.ink_pass(group, InkPass::Cmy, (x0, y0), width, height) else {
            return false;
        };
        let Some(k) = self.ink_pass(group, InkPass::K, (x0, y0), width, height) else {
            return false;
        };

        let Some(cmy_pixels) = straight_rgba(&cmy) else {
            return false;
        };
        let Some(k_pixels) = straight_rgba(&k) else {
            return false;
        };

        // A tint is the pass's colour complemented; each pass's own alpha carries its coverage.
        let mut tints = Vec::with_capacity(usize::from(width) * usize::from(height) * 4);
        for (p, q) in cmy_pixels.chunks_exact(4).zip(k_pixels.chunks_exact(4)) {
            tints.extend_from_slice(&[255 - p[0], 255 - p[1], 255 - p[2], 255 - q[0]]);
        }
        let mut rgb = vec![0u8; tints.len() / 4 * 3];
        // PdfCraft patch: converted the same way the device draws solid DeviceCMYK colour.
        if hayro_interpret::color::cmyk_to_srgb8(&tints, &mut rgb).is_none() {
            return false;
        }

        let mut rgba = Vec::with_capacity(rgb.len() / 3 * 4);
        for (rgb, p) in rgb.chunks_exact(3).zip(cmy_pixels.chunks_exact(4)) {
            rgba.extend_from_slice(&[rgb[0], rgb[1], rgb[2], p[3]]);
        }
        let result = Pixmap::from_parts(
            rgba,
            width,
            height,
            vello_cpu::PixelMetadata::new(
                    // PdfCraft patch (bug-hunter-002 P2-2): the composited
                    // rgba above is STRAIGHT alpha; the default metadata is
                    // AlphaPremultiplied, which made AA edges composite
                    // bright when the pixmap is drawn.
                vello_cpu::peniko::ImageAlphaType::Alpha,
                true,
            ),
        );

        // The group's own soft mask and blend mode are on its layer; what the contents set
        // must not apply to the finished pixels a second time.
        self.ctx.reset_mask();
        self.ctx
            .set_blend_mode(vello_cpu::peniko::BlendMode::default());
        self.ctx.set_paint_transform(Affine::IDENTITY);
        self.draw_pixmap(
            Arc::new(result),
            ImageQuality::Low,
            kurbo::Affine::translate((x0, y0)),
        );
        true
    }

    /// PdfCraft patch (27): a nested isolated group that names its own (non-CMYK) blending
    /// space can no longer be drawn inside this pass.
    pub(crate) fn group_blends_outside_device_cmyk(&mut self) {
        if self.ink.pass.is_some() {
            self.ink.unsupported = true;
        }
    }
}

/// PdfCraft patch: straight (non-premultiplied) RGBA8 pixels of a rendered canvas, through the
/// one pixel path the vendored tree may rely on (`into_png` + the crate's own `image`
/// dependency), because vello_cpu 0.3 exposes no verified direct pixel accessor.
fn straight_rgba(pix: &Pixmap) -> Option<Vec<u8>> {
    let png = pix.clone().into_png().ok()?;
    let img = image::load_from_memory(&png).ok()?.into_rgba8();
    Some(img.into_raw())
}
