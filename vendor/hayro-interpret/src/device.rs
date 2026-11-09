use crate::CmykGroup;
use crate::font::GlyphRun;
use crate::x_object::soft_mask::SoftMask;
use crate::{BlendMode, ClipPath, FillRule, Image};
use crate::{DrawMode, DrawProps, ImageDrawProps};
use hayro_syntax::object::Dict;
use kurbo::{BezPath, Rect, Shape};

/// A trait for a device that can be used to process PDF drawing instructions.
pub trait Device<'a> {
    /// Draw a path.
    fn draw_path(&mut self, path: &BezPath, props: DrawProps<'a>, draw_mode: &DrawMode);
    /// Push a new clip path to the clip stack.
    fn push_clip_path(&mut self, clip_path: &ClipPath);
    /// Push a rectangular clip to the clip stack.
    fn push_clip_rect(&mut self, rect: &Rect) {
        self.push_clip_path(&ClipPath {
            path: rect.to_path(0.1),
            fill: FillRule::NonZero,
        });
    }
    /// Push a new transparency group to the blend stack.
    fn push_transparency_group(
        &mut self,
        opacity: f32,
        mask: Option<SoftMask<'a>>,
        blend_mode: BlendMode,
        isolated: bool,
        knockout: bool,
    );
    /// Draw a run of positioned glyphs.
    fn draw_glyph_run(
        &mut self,
        glyph_run: &GlyphRun<'_, 'a>,
        props: DrawProps<'a>,
        draw_mode: &DrawMode,
    );
    /// Draw an image.
    fn draw_image(&mut self, image: Image<'a, '_>, props: ImageDrawProps<'a>);
    /// Pop the last clip path or clip rectangle from the clip stack.
    fn pop_clip(&mut self);
    /// Pop the last transparency group from the blend stack.
    fn pop_transparency_group(&mut self);
    /// Draw a rectangle directly, without going through the general path pipeline.
    fn draw_rect(&mut self, rect: &Rect, props: DrawProps<'a>, draw_mode: &DrawMode) {
        self.draw_path(&rect.to_path(0.1), props, draw_mode);
    }
    /// Called at the beginning of a marked content sequence (BMC/BDC).
    ///
    /// The tag is the marked content tag (e.g. b"P", b"Span"). The properties are the property list
    /// dictionary.
        pub(crate) fn draw(
        &self,
        resources: &Resources<'a>,
        context: &mut Context<'a>,
        device: &mut impl Device<'a>,
    ) {
        if !context.ocg_state.is_visible() {
            return;
        }

        if !context.begin_nested_interpretation() {
            return;
        }

        let has_oc = xobject_oc(&self.dict, context);
        if !context.ocg_state.is_visible() {
            if has_oc {
                context.ocg_state.end_marked_content();
            }
            context.end_nested_interpretation();
            return;
        }

        let iter = TypedIter::new(self.decoded.as_ref());

        context.path_mut().truncate(0);
        context.save_state();
        context.pre_concat_affine(self.matrix);
        context.push_root_transform();

        // PdfCraft patch: a transparency group that is neither isolated nor knockout composites
        // its elements with the backdrop already on the page (§11.4.7), so when it is entered at
        // full opacity with no soft mask and Normal blending there is nothing for a layer of its
        // own to add, and drawing straight onto the page is the exact result. A layer here was
        // wrong, not just slower: a blend mode inside the group saw the layer's transparent
        // backdrop instead of the content beneath, and the layer then covered the page as an
        // opaque block. An isolated group starts from a blank backdrop by definition, and a
        // knockout group or one entered at partial opacity, under a soft mask or with a blend
        // mode keeps its layer.
        let entered_plainly = {
            let state = context.fill_paint();
            let stroke = context.stroke_paint();
            state.alpha >= 1.0
                && stroke.alpha >= 1.0
                && context.compositing().soft_mask.is_none()
                && context.compositing().blend_mode == crate::BlendMode::Normal
        };
        let own_layer = self.is_transparency_group && (self.isolated || self.knockout || !entered_plainly);

        // PdfCraft patch: a group drawn as an element of a knockout group replaces what that
        // group drew under its shape. The shape is cleared by drawing the group's content once
        // in the device's erase mode; the group then composites over the cleared area in a
        // layer of its own.
        let knockout_element = self.is_transparency_group && device.knockout_erase_wanted();
        if knockout_element {
            context.save_state();
            device.begin_knockout_erase();
            device.push_clip_path(&ClipPath {
                path: context.ctm()
                    * Rect::new(
                        self.bbox[0] as f64,
                        self.bbox[1] as f64,
                        self.bbox[2] as f64,
                        self.bbox[3] as f64,
                    )
                    .to_path(0.1),
                fill: FillRule::NonZero,
            });
            interpret(
                iter.clone(),
                self.resources.as_ref().unwrap_or(resources),
                context,
                device,
            );
            device.pop_clip();
            device.end_knockout_erase();
            context.restore_state(device);
            context.path_mut().truncate(0);
        }

        if own_layer {
            device.push_transparency_group(
                context.fill_paint().alpha,
                std::mem::take(&mut context.compositing_mut().soft_mask),
                std::mem::take(&mut context.compositing_mut().blend_mode),
                self.isolated,
                self.knockout,
            );

            context.fill_paint_mut().alpha = 1.0;
            context.stroke_paint_mut().alpha = 1.0;
        }

        // PdfCraft patch: an isolated group that names another blending space blends in that
        // space (§11.3.4); a device drawing an enclosing group in CMYK is told, so it can
        // decline that group. A non-isolated group's /CS is ignored: it blends in its parent's
        // space (ISO 32000-2 Table 145, CS).
        if self.is_transparency_group
            && self.isolated
            && names_a_blending_space(&self.dict)
            && !blends_in_device_cmyk(&self.dict)
        {
            device.group_blends_outside_device_cmyk();
        }

        // PdfCraft patch: an isolated group blending in DeviceCMYK starts from a backdrop of
        // its own, so a device able to blend in CMYK computes its interior apart and composites
        // the result through the layer pushed above (see `cmyk_group`). Knockout groups are
        // left out: their elements composite with the group's initial backdrop, not with each
        // other.
        let blended_in_cmyk = own_layer
            && self.isolated
            && !self.knockout
            && blends_in_device_cmyk(&self.dict)
            && device.draw_isolated_cmyk_group(&CmykGroup::new(
                self,
                self.resources.as_ref().unwrap_or(resources),
                context,
            ));

        if !blended_in_cmyk {
            device.push_clip_path(&ClipPath {
                path: context.ctm()
                    * Rect::new(
                        self.bbox[0] as f64,
                        self.bbox[1] as f64,
                        self.bbox[2] as f64,
                        self.bbox[3] as f64,
                    )
                    .to_path(0.1),
                fill: FillRule::NonZero,
            });

            interpret(
                iter,
                self.resources.as_ref().unwrap_or(resources),
                context,
                device,
            );

            device.pop_clip();
        }

        if own_layer {
            device.pop_transparency_group();
        }

        context.pop_root_transform();
        context.restore_state(device);

        if has_oc {
            context.ocg_state.end_marked_content();
        }

        context.end_nested_interpretation();
    }

fn begin_marked_content(&mut self, _tag: &[u8], _properties: Option<&Dict<'_>>) {}
    /// Called at the end of a marked content sequence (EMC).
    fn end_marked_content(&mut self) {}
}

/// A device that discards all drawing operations.
pub struct DummyDevice;

impl Device<'_> for DummyDevice {
    fn draw_path(&mut self, _: &BezPath, _: DrawProps<'_>, _: &DrawMode) {}
    fn push_clip_path(&mut self, _: &ClipPath) {}
    fn push_transparency_group(
        &mut self,
        _: f32,
        _: Option<SoftMask<'_>>,
        _: BlendMode,
        _: bool,
        _: bool,
    ) {
    }
    fn draw_glyph_run(&mut self, _: &GlyphRun<'_, '_>, _: DrawProps<'_>, _: &DrawMode) {}
    fn draw_image(&mut self, _: Image<'_, '_>, _: ImageDrawProps<'_>) {}
    fn pop_clip(&mut self) {}
    fn pop_transparency_group(&mut self) {}
}
