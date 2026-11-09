/*!
A crate for rendering PDF files.

This crate allows you to render pages of a PDF file into bitmaps. It is supposed to be relatively
lightweight, since we do not have any dependencies on the GPU. All the rendering happens on the CPU.

The ultimate goal of this crate is to be a *feature-complete* and *performant* PDF rasterizer.
With that said, we are currently still very far away from reaching that goal: So far, no effort
has been put into performance optimizations, as we are still working on implementing missing features.
However, this crate is currently the most comprehensive and feature-complete
implementation of a PDF rasterizer in pure Rust. This claim is supported by the fact that we currently
include over 1000 PDF files in our regression test suite. The majority of those have been scraped
from the `pdf.js` and `PDFBOX` test suites and therefore represent a very large and diverse sample
of PDF files.

As mentioned, there are still some serious limitations, including lack of support for
encrypted/password-protected PDF files, blending and isolation, knockout groups as well as a range
of smaller features such as color key masking. But you should be able to render the vast majority
of PDF files without too many issues.

## Safety
This crate forbids unsafe code via a crate-level attribute.

## Examples
For usage examples, see the [example](https://github.com/LaurenzV/hayro/tree/master/hayro/examples) in
the GitHub repository.

## Cargo features
This crate has one optional feature:
- `embed-fonts`: See the description of [`hayro-interpret`](https://docs.rs/hayro-interpret/latest/hayro_interpret/#cargo-features) for more information.
*/

#![forbid(unsafe_code)]
#![deny(missing_docs)]

use fearless_simd::Level;
use glyph_cache::GlyphCache;
use hayro_interpret::Device;
use hayro_interpret::FillRule;
use hayro_interpret::InterpreterCache;
use hayro_interpret::InterpreterSettings;
use hayro_interpret::font::{GlyphRun, HintingCache};
use hayro_interpret::hayro_syntax::Pdf;
use hayro_interpret::hayro_syntax::page::Page;
use hayro_interpret::util::{RectExt, TransformExt};
use hayro_interpret::{BlendMode, Context, DrawMode, DrawProps, ImageDrawProps, SoftMask};
use hayro_interpret::{ClipPath, interpret_page};
use kurbo::{Affine, BezPath, Rect};
use pic_scale::Scaler;
use rustc_hash::FxHashMap;
use std::cell::RefCell;
use std::ops::RangeInclusive;
use std::rc::Rc;

pub use hayro_interpret;
pub use hayro_interpret::hayro_syntax;
pub use kurbo;
pub use vello_cpu;

use vello_cpu::color::palette::css::{TRANSPARENT, WHITE};
use vello_cpu::color::{AlphaColor, Srgb};
use vello_cpu::peniko::{Compose, Fill, Mix};
use vello_cpu::{Mask, Pixmap, RenderContext, peniko};

mod clip;
mod cmyk;
mod glyph;
mod glyph_cache;
mod image;
mod mask;
mod paint;
mod path;
mod tiling;

type OutlineCache = FxHashMap<(u128, u32), Option<Rc<BezPath>>>;

pub(crate) struct GlobalState {
    level: Level,
    outline_cache: Rc<RefCell<OutlineCache>>,
    glyph_cache: Rc<RefCell<GlyphCache>>,
    hinting_cache: HintingCache,
    scaler: Scaler,
    force_image_interpolation: bool,
    aliasing_threshold: Option<u8>,
    hinting: bool,
    glyph_caching: bool,
}

impl GlobalState {
    pub(crate) fn new(cache: &RenderCache<'_>, settings: &RenderSettings) -> Self {
        Self {
            level: Level::new(),
            outline_cache: cache.outline_cache.clone(),
            glyph_cache: cache.glyph_cache.clone(),
            hinting_cache: cache.hinting_cache.clone(),
            scaler: Scaler::new(image::RESAMPLING_FUNCTION),
            force_image_interpolation: settings.force_image_interpolation,
            aliasing_threshold: settings.aliasing_threshold,
            hinting: settings.hinting,
            glyph_caching: settings.glyph_caching,
        }
    }
}

pub(crate) struct Renderer<'a> {
    global: &'a GlobalState,
    pub(crate) ctx: &'a mut RenderContext,
    pub(crate) inside_pattern: bool,
    pub(crate) soft_mask_cache: FxHashMap<u128, Mask>,
    pub(crate) in_type3_glyph: bool,
    // PdfCraft patch: knockout machinery (see `set_knockout`) and the ink pass of a group
    // blending in DeviceCMYK (see `cmyk`).
    pub(crate) knockout: Vec<bool>,
    pub(crate) erasing: bool,
    pub(crate) erase_groups: usize,
    pub(crate) ink: cmyk::InkState,
}

impl<'r> Renderer<'r> {
    pub(crate) fn new(ctx: &'r mut RenderContext, global: &'r GlobalState) -> Self {
        Self {
            global,
            ctx,
            inside_pattern: false,
            soft_mask_cache: FxHashMap::default(),
            in_type3_glyph: false,
            knockout: Vec::new(),
            erasing: false,
            erase_groups: 0,
            ink: cmyk::InkState::default(),
        }
    }

    fn misses_target(&self, bbox: Rect) -> bool {
        let target = Rect::new(0.0, 0.0, self.ctx.width() as f64, self.ctx.height() as f64);
        !bbox.overlaps(target)
    }

    // PdfCraft patch: whether the innermost transparency group is a knockout group.
    pub(crate) fn in_knockout(&self) -> bool {
        self.knockout.last().copied().unwrap_or(false)
    }

    fn child_context(&self, width: u16, height: u16) -> RenderContext {
        let mut ctx =
            RenderContext::new_with(width, height, derive_settings(self.ctx.render_settings()));
        ctx.set_aliasing_threshold(self.global.aliasing_threshold);
        ctx
    }

    fn apply_draw_props(&mut self, props: &DrawProps<'_>) {
        self.ctx.set_transform(props.transform);
        // PdfCraft patch: inside a knockout group every element replaces what it covers; in an
        // erase pass it clears what is under its shape instead of painting.
        if self.erasing {
            self.ctx.set_blend_mode(knockout_blend_mode());
        } else if self.in_knockout() {
            self.ctx.set_blend_mode(knockout_blend_mode());
        } else {
            self.ctx.set_blend_mode(convert_blend_mode(props.blend_mode));
        }
        self.apply_soft_mask(props.soft_mask.as_ref());
    }

    fn apply_image_props(&mut self, props: &ImageDrawProps<'_>) {
        self.ctx.set_transform(props.transform);
        // PdfCraft patch: see `apply_draw_props`.
        if self.erasing || self.in_knockout() {
            self.ctx.set_blend_mode(knockout_blend_mode());
        } else {
            self.ctx.set_blend_mode(convert_blend_mode(props.blend_mode));
        }
        self.apply_soft_mask(props.soft_mask.as_ref());
    }
}

impl<'a, 'r> Device<'a> for Renderer<'r> {
    fn draw_image(&mut self, image: hayro_interpret::Image<'a, '_>, props: ImageDrawProps<'a>) {
        Self::draw_pdf_image(self, image, props);
    }

    fn push_clip_path(&mut self, clip_path: &ClipPath) {
        Self::push_clip_path(self, clip_path);
    }

    fn push_clip_rect(&mut self, rect: &Rect) {
        Self::push_clip_rect(self, rect);
    }

    fn push_transparency_group(
        &mut self,
        opacity: f32,
        mask: Option<SoftMask<'a>>,
        blend_mode: BlendMode,
        isolated: bool,
        knockout: bool,
    ) {
        // PdfCraft patch: nothing is layered in a knockout erase pass.
        if self.erasing {
            self.erase_groups += 1;
            return;
        }
        let _ = isolated; // PdfCraft patch: group layers here already start from a blank
                          // backdrop, which is the isolated behaviour.
        Self::push_transparency_group(self, opacity, mask, blend_mode, knockout);
    }

    fn pop_clip(&mut self) {
        Self::pop_clip(self);
    }

    fn pop_transparency_group(&mut self) {
        // PdfCraft patch: nothing is layered in a knockout erase pass.
        if self.erasing && self.erase_groups > 0 {
            self.erase_groups -= 1;
            return;
        }
        self.knockout.pop();

        Self::pop_transparency_group(self);
    }

    fn draw_isolated_cmyk_group(&mut self, group: &hayro_interpret::CmykGroup<'a, '_>) -> bool {
        Self::draw_isolated_cmyk_group(self, group)
    }

    fn group_blends_outside_device_cmyk(&mut self) {
        Self::group_blends_outside_device_cmyk(self)
    }

    fn knockout_erase_wanted(&mut self) -> bool {
        // PdfCraft patch (bug-hunter-002 P1-1): a nested /Group form inside a
        // knockout element must not re-enter the erase pass — mask.rs and
        // cmyk.rs already guard on `erasing`; the renderer-side gate is the
        // same. Without it a nested end_knockout_erase clears `erasing`
        // mid-pass and the outer content paints with Compose::Copy.
        !self.erasing && self.in_knockout()
    }

    fn begin_knockout_erase(&mut self) {
        self.erasing = true;
        self.erase_groups = 0;
    }

    fn end_knockout_erase(&mut self) {
        self.erasing = false;
        self.erase_groups = 0;
        if self.in_knockout() {
            self.ctx.set_blend_mode(knockout_blend_mode());
        }
    }

    fn draw_path(&mut self, path: &BezPath, props: DrawProps<'a>, draw_mode: &DrawMode) {
        Self::draw_path(self, path, props, draw_mode);
    }

    fn draw_rect(&mut self, rect: &Rect, props: DrawProps<'a>, draw_mode: &DrawMode) {
        Self::draw_rect(self, rect, props, draw_mode);
    }

    fn draw_glyph_run(
        &mut self,
        glyph_run: &GlyphRun<'_, 'a>,
        props: DrawProps<'a>,
        draw_mode: &DrawMode,
    ) {
        Self::draw_glyph_run(self, glyph_run, props, draw_mode);
    }
}

/// A cache used by the renderer.
///
/// Ideally, such a cache should be constructed once per PDF and then reused across
/// multiple render invocations on the same document.
// TODO: Allow clearing the cache.
#[derive(Default)]
pub struct RenderCache<'a> {
    pub(crate) interpreter_cache: InterpreterCache<'a>,
    pub(crate) outline_cache: Rc<RefCell<OutlineCache>>,
    pub(crate) glyph_cache: Rc<RefCell<GlyphCache>>,
    pub(crate) hinting_cache: HintingCache,
}

impl<'a> RenderCache<'a> {
    /// Create a new render cache.
    pub fn new() -> Self {
        Self {
            interpreter_cache: InterpreterCache::new(),
            outline_cache: Rc::new(RefCell::new(FxHashMap::default())),
            glyph_cache: Rc::new(RefCell::new(GlyphCache::default())),
            hinting_cache: HintingCache::new(),
        }
    }
}

/// Settings for rendering a page.
#[derive(Clone, Copy)]
pub struct RenderSettings {
    /// Whether all images should forcibly be rendered with bilinear interpolation.
    pub force_image_interpolation: bool,
    /// The aliasing threshold for fills and strokes.
    ///
    /// `None` (the default) enables anti-aliasing; `Some(threshold)` paints pixels
    /// whose coverage exceeds the threshold.
    pub aliasing_threshold: Option<u8>,
    /// Whether text should be hinted.
    pub hinting: bool,
    /// Whether rasterized glyphs should be cached and reused.
    pub glyph_caching: bool,
}

impl Default for RenderSettings {
    fn default() -> Self {
        Self {
            force_image_interpolation: false,
            aliasing_threshold: None,
            hinting: true,
            glyph_caching: true,
        }
    }
}

/// Settings for the output pixmap.
#[derive(Clone, Copy)]
pub struct PixmapSettings {
    /// Horizontal scale factor.
    pub x_scale: f32,
    /// Vertical scale factor.
    pub y_scale: f32,
    /// Background color.
    pub bg_color: AlphaColor<Srgb>,
}

impl Default for PixmapSettings {
    fn default() -> Self {
        Self {
            x_scale: 1.0,
            y_scale: 1.0,
            bg_color: TRANSPARENT,
        }
    }
}

/// Render a page to a pixmap with the given settings.
///
/// This is a "simplified" method in case all you want to achieve is rendering a
/// page with a certain scale factor.
///
/// If you want to have more control over how exactly the page is rendered (e.g. with
/// an arbitrary transform or into a custom pixel buffer), use [`render_into`]
pub fn render<'a>(
    page: &'a Page<'a>,
    cache: &RenderCache<'a>,
    interpreter_settings: &InterpreterSettings,
    render_settings: &RenderSettings,
    pixmap_settings: &PixmapSettings,
) -> Pixmap {
    let (width, height) = page.render_dimensions();
    // PdfCraft patch: the pixel dimensions are pre-clamped in u32 before the u16 casts, so a
    // hostile page box cannot produce a 65535 × 65535 canvas: the scales are lowered until the
    // output fits both the backend's sides and a 2^28-pixel budget (the renderer's image cap).
    const MAX_RENDER_PIXELS: u64 = 1 << 28;
    let mut x_scale = pixmap_settings.x_scale;
    let mut y_scale = pixmap_settings.y_scale;
    if x_scale.is_finite() && y_scale.is_finite() && x_scale > 0.0 && y_scale > 0.0 {
        // PdfCraft patch (bug-hunter-002 P3-1): each axis clamps against its
        // own dimension — the old closure checked height for both, letting a
        // width>65535pt page saturate the u16 cast.
        let fit_side = |dim: f32, s: f32| {
            let max = u64::from(u16::MAX);
            if (dim.max(0.0) as f64 * s as f64) as u64 > max {
                max as f32 / dim.max(1.0) as f32
            } else {
                s
            }
        };
        x_scale = fit_side(width.max(0.0) as f32, x_scale);
        y_scale = fit_side(height.max(0.0) as f32, y_scale);
        let pixels = (width.max(0.0) as f64 * x_scale as f64) as u64
            * (height.max(0.0) as f64 * y_scale as f64) as u64;
        if pixels > MAX_RENDER_PIXELS {
            let shrink = (MAX_RENDER_PIXELS as f64 / pixels.max(1) as f64).sqrt();
            x_scale *= shrink as f32;
            y_scale *= shrink as f32;
        }
    } else {
        x_scale = 0.0;
        y_scale = 0.0;
    }
    let mut ctx = RenderContext::new(
        (width.max(0.0) * x_scale) as u16,
        (height.max(0.0) * y_scale) as u16,
    );
    let transform = Affine::scale_non_uniform(x_scale as f64, y_scale as f64)
        * page.initial_transform(true).to_kurbo();
    render_into(
        page,
        cache,
        interpreter_settings,
        render_settings,
        &mut ctx,
        transform,
    );
    ctx.flush();

    let mut pixmap = Pixmap::new(ctx.width(), ctx.height());
    ctx.render_with(
        &mut pixmap,
        &mut vello_cpu::Resources::default(),
        vello_cpu::RasterizerSettings {
            target_init: vello_cpu::TargetInit::Clear(pixmap_settings.bg_color),
            ..Default::default()
        },
    );
    pixmap
}

/// Render a page into the given [`RenderContext`] with the
/// given transform.
///
/// See the [following example](https://github.com/LaurenzV/hayro/blob/main/hayro/examples/render.rs)
/// if you are unsure how to call this method.
pub fn render_into<'a>(
    page: &'a Page<'a>,
    cache: &RenderCache<'a>,
    interpreter_settings: &InterpreterSettings,
    render_settings: &RenderSettings,
    ctx: &mut RenderContext,
    transform: Affine,
) {
    let target = Rect::new(0.0, 0.0, ctx.width() as f64, ctx.height() as f64);
    let mut state = Context::new(
        transform,
        target,
        &cache.interpreter_cache,
        page.xref(),
        interpreter_settings.clone(),
    );
    ctx.take_current_state();
    ctx.reset_mask();
    ctx.reset_filter_effect();
    ctx.set_aliasing_threshold(render_settings.aliasing_threshold);

    let global = GlobalState::new(cache, render_settings);
    let mut device = Renderer::new(ctx, &global);
    let crop_box = page.intersected_crop_box().to_kurbo();
    let [a, b, c, d, _, _] = transform.as_coeffs();
    let axis_aligned = (b == 0.0 && c == 0.0) || (a == 0.0 && d == 0.0);
    let mapped_crop_box = transform
        .transform_rect_bbox(crop_box)
        .inflate(1.0 / 256.0, 1.0 / 256.0);
    let clip_crop_box = !axis_aligned || !mapped_crop_box.contains_rect(target);

    if clip_crop_box {
        device.ctx.set_transform(transform);
        device.ctx.push_clip_rect(&crop_box);
    }

    interpret_page(page, &mut state, &mut device);

    if clip_crop_box {
        device.pop_clip();
    }
}

// Just a convenience method for testing.
#[doc(hidden)]
pub fn render_pdf(
    pdf: &Pdf,
    scale: f32,
    settings: InterpreterSettings,
    render_settings: &RenderSettings,
    range: Option<RangeInclusive<usize>>,
) -> Option<Vec<Pixmap>> {
    let cache = RenderCache::new();
    let rendered = pdf
        .pages()
        .iter()
        .enumerate()
        .flat_map(|(idx, page)| {
            if range.clone().is_some_and(|range| !range.contains(&idx)) {
                return None;
            }

            let pixmap = render(
                page,
                &cache,
                &settings,
                render_settings,
                &PixmapSettings {
                    x_scale: scale,
                    y_scale: scale,
                    bg_color: WHITE,
                },
            );

            Some(pixmap)
        })
        .collect();

    Some(rendered)
}

pub(crate) fn derive_settings(settings: &vello_cpu::RenderSettings) -> vello_cpu::RenderSettings {
    vello_cpu::RenderSettings {
        num_threads: 0,
        ..*settings
    }
}

/// PdfCraft patch: inside a knockout group every element replaces what it covers: elements are
/// drawn with `Compose::Copy`, which (weighted by coverage) replaces what earlier elements drew
/// where the element's shape is, so each element composites with the group's initial backdrop
/// (transparent: the non-isolated approximation every group layer here already makes) instead of
/// with the elements before it. Against a transparent backdrop the element's own blend mode has
/// nothing to blend with, so the mix is Normal. This covers paths, strokes, glyph outlines,
/// images and shadings drawn directly; elements drawn into a layer of their own are preceded by
/// a knockout erase pass (see `Device::begin_knockout_erase`).
pub(crate) fn knockout_blend_mode() -> peniko::BlendMode {
    peniko::BlendMode::new(Mix::Normal, Compose::Copy)
}

fn convert_fill_rule(fill_rule: FillRule) -> Fill {
    match fill_rule {
        FillRule::NonZero => Fill::NonZero,
        FillRule::EvenOdd => Fill::EvenOdd,
    }
}

fn convert_blend_mode(blend_mode: BlendMode) -> peniko::BlendMode {
    let mix = match blend_mode {
        BlendMode::Normal => Mix::Normal,
        BlendMode::Multiply => Mix::Multiply,
        BlendMode::Screen => Mix::Screen,
        BlendMode::Overlay => Mix::Overlay,
        BlendMode::Darken => Mix::Darken,
        BlendMode::Lighten => Mix::Lighten,
        BlendMode::ColorDodge => Mix::ColorDodge,
        BlendMode::ColorBurn => Mix::ColorBurn,
        BlendMode::HardLight => Mix::HardLight,
        BlendMode::SoftLight => Mix::SoftLight,
        BlendMode::Difference => Mix::Difference,
        BlendMode::Exclusion => Mix::Exclusion,
        BlendMode::Hue => Mix::Hue,
        BlendMode::Saturation => Mix::Saturation,
        BlendMode::Color => Mix::Color,
        BlendMode::Luminosity => Mix::Luminosity,
    };

    peniko::BlendMode::new(mix, Compose::SrcOver)
}
