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
    /// PdfCraft patch: whether a transparency group drawn now is an element of a knockout
    /// group that the device wants cleared under the group's shape first. If so the
    /// interpreter draws the group's content once between `begin_knockout_erase` and
    /// `end_knockout_erase`, in which every paint clears what is under its shape instead of
    /// painting, and then draws the group as usual, in a layer of its own.
    fn knockout_erase_wanted(&mut self) -> bool {
        false
    }
    /// PdfCraft patch: see `knockout_erase_wanted`.
    fn begin_knockout_erase(&mut self) {}
    /// PdfCraft patch: see `knockout_erase_wanted`.
    fn end_knockout_erase(&mut self) {}
    /// PdfCraft patch: draw the contents of an isolated transparency group that blends in
    /// `DeviceCMYK` (ISO 32000-2 §11.3.4), already inside the layer it is composited
    /// through. Return `true` when the device drew them; on `false` (the default) the
    /// interpreter draws the contents into the device as for any other group.
    fn draw_isolated_cmyk_group(&mut self, _group: &CmykGroup<'a, '_>) -> bool {
        false
    }
    /// PdfCraft patch: an isolated transparency group that names a blending colour space
    /// other than `DeviceCMYK` is about to be drawn. A device drawing an enclosing group in
    /// CMYK can no longer draw that group right, since this one blends in its own space.
    fn group_blends_outside_device_cmyk(&mut self) {}
    /// Called at the beginning of a marked content sequence (BMC/BDC).
    ///
    /// The tag is the marked content tag (e.g. b"P", b"Span"). The properties are the property list
    /// dictionary.
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
