use crate::{Renderer, knockout_blend_mode};
use fearless_simd::{Level, Select, Simd, SimdBase, SimdInto, mask8x32, u8x32, u16x16};
use hayro_interpret::util::x_y_advances;
use hayro_interpret::{FillRule, ImageData, ImageDrawProps, LumaData, Paint, RgbData};
use kurbo::{Affine, Rect, Shape};
use pic_scale::{
    ImageSize, ImageStore, ImageStoreMut, PicScaleError, Resampling, ResamplingFunction, Scaler,
};
use std::mem::size_of;
use std::sync::Arc;
use vello_cpu::peniko::{Compose, Fill, ImageQuality, ImageSampler, Mix};
use vello_cpu::{Image, ImageSource, Mask, PixelMetadata, Pixmap, peniko};

// Previously, we used `CatmullRom`. The problem with that one is that it
// can have negative weights. If we pass a premultiplied buffer to
// `pic-scale`, it can happen that any of the RGB variants end up
// slightly larger than the alpha channel, leading to pixel artifacts when
// rendering. In order to avoid having to do another pass over the buffer
// to clamp, we instead use Hermite, which has similar quality but doesn't
// have this problem.
pub(super) const RESAMPLING_FUNCTION: ResamplingFunction = ResamplingFunction::Hermite;

// PdfCraft patch: largest image drawn, in pixels (2^28 ≈ 268 MP, well above real page images).
const MAX_IMAGE_PIXELS: u64 = 1 << 28;

// PdfCraft patch: source axes allowed into resampling planning. Unlike the final pixmap,
// sources need not fit u16: a 65536-pixel strip can safely shrink to a 16-pixel image.
const MAX_RESAMPLING_SOURCE_SIDE: u32 = 1 << 20;
const MAX_RESAMPLING_PLAN_BYTES: usize = 128 << 20;

/// PdfCraft patch: source, intermediate and RGBA byte lengths, independent of backend sides.
fn source_byte_len(width: u32, height: u32, channels: usize, max_pixels: u64) -> Option<usize> {
    if width == 0 || height == 0 || !matches!(channels, 1 | 3 | 4) {
        return None;
    }
    let pixels = u64::from(width).checked_mul(u64::from(height))?;
    if pixels > max_pixels {
        return None;
    }
    let pixels = usize::try_from(pixels).ok()?;
    let rgba_len = pixels.checked_mul(4)?;
    if rgba_len > isize::MAX as usize {
        return None;
    }
    pixels.checked_mul(channels)
}

/// PdfCraft patch: only final pixmaps/padded glyphs need to fit the backend's u16 sides.
fn image_byte_len(width: u32, height: u32, channels: usize, max_pixels: u64) -> Option<usize> {
    if width > u32::from(u16::MAX) || height > u32::from(u16::MAX) {
        return None;
    }
    source_byte_len(width, height, channels, max_pixels)
}

/// PdfCraft patch: conservative resampling planning work for one axis in PicScale 0.7.12. Its
/// four-tap kernels grow on downscaling. Include float/32-aligned i16 weights, temporary
/// float/f64/usize tap arrays and two copies of the bounds. This is a size estimate, not a
/// fallible reserve inside PicScale or a bound on total renderer memory.
fn filter_plan_bytes(source: u32, target: u32) -> Option<usize> {
    if source == target {
        return Some(0);
    }
    let (source, target) = (usize::try_from(source).ok()?, usize::try_from(target).ok()?);
    if target == 0 {
        return None;
    }
    let taps = source.checked_mul(4)?.div_ceil(target).max(4);
    let aligned_taps = taps.div_ceil(32).checked_mul(32)?;
    let float_weights = taps.checked_mul(target)?.checked_mul(size_of::<f32>())?;
    let fixed_weights = aligned_taps
        .checked_mul(target)?
        .checked_mul(size_of::<i16>())?;
    let tap_scratch = taps.checked_mul(size_of::<f32>() + size_of::<f64>() + size_of::<usize>())?;
    let bounds = target.checked_mul(4)?.checked_mul(size_of::<usize>())?;
    float_weights
        .checked_add(fixed_weights)?
        .checked_add(tap_scratch)?
        .checked_add(bounds)
}

/// PdfCraft patch: alpha-aware RGBA plans also need source-sized premultiplication scratch.
/// Bound the combined scratch before planning, including a conservative crossed intermediate
/// (or four rows for a short target) when both convolution axes change.
fn resampling_scratch_byte_len(
    source: (u32, u32),
    target: (u32, u32),
    channels: usize,
    max_pixels: u64,
) -> Option<usize> {
    if source == target {
        return Some(0);
    }
    let alpha = if channels == 4 {
        source_byte_len(source.0, source.1, channels, max_pixels)?
    } else {
        0
    };
    let filter = if source.0 != target.0 && source.1 != target.1 {
        source_byte_len(source.0, target.1.max(4), channels, max_pixels)?
    } else {
        0
    };
    let len = alpha.checked_add(filter)?;
    let limit = max_pixels.saturating_mul(4).min(isize::MAX as u64);
    if u64::try_from(len).ok()? > limit {
        return None;
    }
    Some(len)
}

/// PdfCraft patch: validate output, crossed intermediate and filter-work bounds before planning.
/// Source width × target height also bounds PicScale's smaller single-threaded four-row scratch.
fn resampling_byte_len(
    source: (u32, u32),
    target: (u32, u32),
    channels: usize,
    data_len: usize,
    max_pixels: u64,
) -> Option<usize> {
    if source_byte_len(source.0, source.1, channels, max_pixels)? != data_len {
        return None;
    }
    let target_len = image_byte_len(target.0, target.1, channels, max_pixels)?;
    resampling_scratch_byte_len(source, target, channels, max_pixels)?;
    source_byte_len(source.0, target.1, channels, max_pixels)?;
    if source.0 > MAX_RESAMPLING_SOURCE_SIDE || source.1 > MAX_RESAMPLING_SOURCE_SIDE {
        return None;
    }
    let plan_bytes = filter_plan_bytes(source.0, target.0)?
        .checked_add(filter_plan_bytes(source.1, target.1)?)?;
    if plan_bytes > MAX_RESAMPLING_PLAN_BYTES {
        return None;
    }
    Some(target_len)
}

/// PdfCraft patch: allocation failures in image conversion/resampling keep the renderer alive.
fn image_buffer(len: usize) -> Option<Vec<u8>> {
    let mut out = Vec::new();
    out.try_reserve_exact(len).ok()?;
    Some(out)
}

#[derive(Clone, Copy)]
enum ImagePixelFormat {
    Luma,
    Rgb,
    PremultipliedRgba,
}

struct SolidColorImage {
    color: [u8; 3],
    width: u32,
    height: u32,
    interpolate: bool,
}

enum RenderImageData {
    Rgb(RgbData),
    Luma(LumaData),
    Solid(SolidColorImage),
}

impl From<ImageData> for RenderImageData {
    fn from(value: ImageData) -> Self {
        match value {
            ImageData::Rgb(rgb) => Self::Rgb(rgb),
            ImageData::Luma(luma) => Self::Luma(luma),
        }
    }
}

impl From<RgbData> for RenderImageData {
    fn from(value: RgbData) -> Self {
        Self::Rgb(value)
    }
}

impl RenderImageData {
    fn width(&self) -> u32 {
        match self {
            Self::Rgb(d) => d.width,
            Self::Luma(d) => d.width,
            Self::Solid(d) => d.width,
        }
    }

    fn height(&self) -> u32 {
        match self {
            Self::Rgb(d) => d.height,
            Self::Luma(d) => d.height,
            Self::Solid(d) => d.height,
        }
    }

    fn interpolate(&self) -> bool {
        match self {
            Self::Rgb(d) => d.interpolate,
            Self::Luma(d) => d.interpolate,
            Self::Solid(d) => d.interpolate,
        }
    }
}

impl Renderer<'_> {
    fn draw_image_with_alpha_mask(&mut self, image_data: RenderImageData, alpha_data: LumaData) {
        let img_width = image_data.width();
        let img_height = image_data.height();
        let image_transform = *self.ctx.transform();
        // PdfCraft patch: this element composites through layers of its own, which cannot be
        // copied; clear what the knockout group drew under the image's placement first.
        if self.in_knockout() && !self.erasing {
            let rect = Rect::new(
                0.0,
                0.0,
                f64::from(image_data.width()),
                f64::from(image_data.height()),
            );
            let outline = image_transform * rect.to_path(0.1);
            self.knockout_erase_shape(&outline);
        }
        let mask_transform = image_transform
            * Affine::scale_non_uniform(
                img_width as f64 / alpha_data.width as f64,
                img_height as f64 / alpha_data.height as f64,
            );
        let mask_image = SolidColorImage {
            color: [0; 3],
            width: alpha_data.width,
            height: alpha_data.height,
            interpolate: alpha_data.interpolate,
        };

        self.ctx.push_layer(None, None, None, None, None);
        self.ctx.set_transform(image_transform);
        self.draw_image(image_data, None);

        self.ctx.push_layer(
            None,
            Some(peniko::BlendMode::new(Mix::Normal, Compose::DestIn)),
            None,
            None,
            None,
        );
        self.ctx.set_transform(mask_transform);
        // Note that there is a circle between `draw_image` and `draw_image_with_alpha_mask`,
        // but `draw_image_with_alpha_mask` is only called if the dimensions or interpolate
        // values between alpha_data and rgb_data don't match. Here we use a
        // `SolidColorImage` so it doesn't affect it.
        self.draw_image(RenderImageData::Solid(mask_image), Some(alpha_data));
        self.ctx.pop_layer();
        self.ctx.pop_layer();
        self.ctx.set_transform(image_transform);
    }

    fn draw_image(&mut self, image_data: impl Into<RenderImageData>, alpha_data: Option<LumaData>) {
        // PdfCraft patch (bug-hunter-002 P2-1): images cannot be tinted
        // per-ink — mark a live CMYK pass unsupported (documented fallback)
        // instead of compositing literal RGB into ink canvases.
        if self.ink.is_passing() {
            self.ink.unsupported = true;
        }
        let image_data = image_data.into();
        // PdfCraft patch: in a knockout erase pass an image clears its placement rectangle
        // instead of painting.
        if self.erasing {
            let cur_transform = *self.ctx.transform();
            let rect = Rect::new(
                0.0,
                0.0,
                f64::from(image_data.width()),
                f64::from(image_data.height()),
            );
            let outline = cur_transform * rect.to_path(0.1);
            self.ctx.set_blend_mode(knockout_blend_mode());
            self.knockout_erase_shape(&outline);
            return;
        }
        let cur_transform = *self.ctx.transform();
        let mut additional_transform = Affine::IDENTITY;

        let (x_scale, y_scale) = {
            let (x, y) = x_y_advances(&cur_transform);
            (x.length() as f32, y.length() as f32)
        };
        let mut img_width = image_data.width();
        let mut img_height = image_data.height();
        let interpolate = image_data.interpolate() || self.global.force_image_interpolation;

        if let Some(a) = &alpha_data
            && (a.width != img_width
                || a.height != img_height
                || (a.interpolate || self.global.force_image_interpolation) != interpolate)
        {
            return self.draw_image_with_alpha_mask(image_data, alpha_data.unwrap());
        }

        let mut quality = if interpolate {
            ImageQuality::Medium
        } else {
            ImageQuality::Low
        };

        let has_alpha = alpha_data.is_some();
        let mut may_have_transparency = has_alpha;
        let needs_resize = x_scale < 1.0 || y_scale < 1.0;
        let (new_width, new_height) = if needs_resize {
            let w = (img_width as f32 * x_scale)
                .ceil()
                .max(1.0)
                .min((u16::MAX / 2) as f32) as u32;
            let h = (img_height as f32 * y_scale)
                .ceil()
                .max(1.0)
                .min((u16::MAX / 2) as f32) as u32;
            if self.in_type3_glyph {
                quality = ImageQuality::High;
            }
            (w, h)
        } else {
            (img_width, img_height)
        };

        // For luma images without alpha, we can resize as single-channel and
        // expand to RGBA afterwards, which is ~4x faster.
        let mut needs_premultiplication = has_alpha;
        let mut rgba_data = if matches!(&image_data, RenderImageData::Solid(_)) && has_alpha {
            let RenderImageData::Solid(solid) = image_data else {
                unreachable!()
            };
            let alpha = alpha_data.unwrap();

            let alpha_data = if !needs_resize {
                alpha.data
            } else {
                // PdfCraft patch: on failure retain the source alpha.
                match resize_image_data(
                    &self.global.scaler,
                    alpha.data.clone(),
                    img_width,
                    img_height,
                    new_width,
                    new_height,
                    ImagePixelFormat::Luma,
                ) {
                    Some(resized_alpha) => {
                        additional_transform = Affine::scale_non_uniform(
                            img_width as f64 / new_width as f64,
                            img_height as f64 / new_height as f64,
                        );
                        img_width = new_width;
                        img_height = new_height;
                        resized_alpha
                    }
                    None => alpha.data,
                }
            };

            let [r, g, b] = solid.color;
            alpha_data
                .iter()
                .flat_map(|a| [r, g, b, *a])
                .collect::<Vec<_>>()
        } else if matches!(&image_data, RenderImageData::Luma(_)) && !has_alpha {
            // We cannot lift this up due to borrowing issues.
            let RenderImageData::Luma(luma) = image_data else {
                unreachable!()
            };

            let luma_data = if !needs_resize {
                luma.data
            } else {
                // PdfCraft patch: on failure retain the source luma.
                match resize_image_data(
                    &self.global.scaler,
                    luma.data.clone(),
                    img_width,
                    img_height,
                    new_width,
                    new_height,
                    ImagePixelFormat::Luma,
                ) {
                    Some(resized) => {
                        additional_transform = Affine::scale_non_uniform(
                            img_width as f64 / new_width as f64,
                            img_height as f64 / new_height as f64,
                        );
                        img_width = new_width;
                        img_height = new_height;
                        resized
                    }
                    None => luma.data,
                }
            };

            let mut out = vec![0; luma_data.len() * 4];
            for (out, g) in out.as_chunks_mut::<4>().0.iter_mut().zip(&luma_data) {
                *out = [*g, *g, *g, 255];
            }
            out
        } else if matches!(&image_data, RenderImageData::Luma(_)) && has_alpha {
            let RenderImageData::Luma(luma) = image_data else {
                unreachable!()
            };
            let alpha = alpha_data.unwrap();

            if !needs_resize {
                luma.data
                    .iter()
                    .zip(&alpha.data)
                    .flat_map(|(g, a)| [*g, *g, *g, *a])
                    .collect::<Vec<_>>()
            } else {
                // PdfCraft patch: on failure expand at the source resolution instead.
                match downscale_with_alpha::<1>(
                    self.global.level,
                    &self.global.scaler,
                    luma.data.clone(),
                    alpha.data.clone(),
                    img_width,
                    img_height,
                    new_width,
                    new_height,
                ) {
                    Some(resized) => {
                        additional_transform = Affine::scale_non_uniform(
                            img_width as f64 / new_width as f64,
                            img_height as f64 / new_height as f64,
                        );
                        img_width = new_width;
                        img_height = new_height;
                        needs_premultiplication = false;
                        resized
                    }
                    None => luma
                        .data
                        .iter()
                        .zip(&alpha.data)
                        .flat_map(|(g, a)| [*g, *g, *g, *a])
                        .collect::<Vec<_>>(),
                }
            }
        } else if matches!(&image_data, RenderImageData::Rgb(_)) && needs_resize {
            let RenderImageData::Rgb(rgb) = image_data else {
                unreachable!()
            };

            // PdfCraft patch: on failure expand at the source resolution instead.
            let resized = if let Some(alpha) = alpha_data {
                match downscale_with_alpha::<3>(
                    self.global.level,
                    &self.global.scaler,
                    rgb.data.clone(),
                    alpha.data.clone(),
                    img_width,
                    img_height,
                    new_width,
                    new_height,
                ) {
                    Some(resized) => {
                        needs_premultiplication = false;
                        additional_transform = Affine::scale_non_uniform(
                            img_width as f64 / new_width as f64,
                            img_height as f64 / new_height as f64,
                        );
                        img_width = new_width;
                        img_height = new_height;
                        resized
                    }
                    None => rgb
                        .data
                        .as_chunks::<3>()
                        .0
                        .iter()
                        .zip(&alpha.data)
                        .flat_map(|(rgb, a)| [rgb[0], rgb[1], rgb[2], *a])
                        .collect::<Vec<_>>(),
                }
            } else {
                match resize_image_data(
                    &self.global.scaler,
                    rgb.data.clone(),
                    img_width,
                    img_height,
                    new_width,
                    new_height,
                    ImagePixelFormat::Rgb,
                ) {
                    Some(resized) => {
                        additional_transform = Affine::scale_non_uniform(
                            img_width as f64 / new_width as f64,
                            img_height as f64 / new_height as f64,
                        );
                        img_width = new_width;
                        img_height = new_height;
                        resized
                            .as_chunks::<3>()
                            .0
                            .iter()
                            .flat_map(|rgb| [rgb[0], rgb[1], rgb[2], 255])
                            .collect::<Vec<_>>()
                    }
                    None => rgb
                        .data
                        .as_chunks::<3>()
                        .0
                        .iter()
                        .flat_map(|rgb| [rgb[0], rgb[1], rgb[2], 255])
                        .collect::<Vec<_>>(),
                }
            };
            resized
        } else {
            let (rgb_data, alpha_data) = match image_data {
                RenderImageData::Rgb(rgb) => (rgb.data, alpha_data.map(|a| a.data)),
                RenderImageData::Luma(luma) => {
                    let mut rgb = vec![0; luma.data.len() * 3];
                    for (rgb, g) in rgb.as_chunks_mut::<3>().0.iter_mut().zip(&luma.data) {
                        *rgb = [*g, *g, *g];
                    }
                    (rgb, alpha_data.map(|a| a.data))
                }
                RenderImageData::Solid(solid) => {
                    let rgb = solid
                        .color
                        .repeat(solid.width as usize * solid.height as usize);
                    (rgb, alpha_data.map(|a| a.data))
                }
            };

            let mut rgba_data = match alpha_data {
                None => rgb_data
                    .as_chunks::<3>()
                    .0
                    .iter()
                    .flat_map(|rgb| [rgb[0], rgb[1], rgb[2], 255])
                    .collect::<Vec<_>>(),
                Some(alpha) => rgb_data
                    .as_chunks::<3>()
                    .0
                    .iter()
                    .zip(alpha)
                    .flat_map(|(rgb, a)| [rgb[0], rgb[1], rgb[2], a])
                    .collect::<Vec<_>>(),
            };

            if !needs_resize {
                rgba_data
            } else {
                premultiply_rgba(self.global.level, &mut rgba_data);

                // PdfCraft patch: on failure keep the premultiplied source pixels.
                match resize_image_data(
                    &self.global.scaler,
                    rgba_data.clone(),
                    img_width,
                    img_height,
                    new_width,
                    new_height,
                    ImagePixelFormat::PremultipliedRgba,
                ) {
                    Some(resized) => {
                        needs_premultiplication = false;
                        additional_transform = Affine::scale_non_uniform(
                            img_width as f64 / new_width as f64,
                            img_height as f64 / new_height as f64,
                        );
                        img_width = new_width;
                        img_height = new_height;
                        resized
                    }
                    None => rgba_data,
                }
            }
        };

        if needs_premultiplication {
            premultiply_rgba(self.global.level, &mut rgba_data);
        }

        // The problem is that by default, when applying a bilinear or bicubic scaling, we will
        // sample pixels using an extend (pad/reflect/repeat). For glyphs, this is undesirable
        // as the glyphs will look very bold. Therefore, for glyphs it is more desirable to sample
        // a transparent pixel when reaching the border. Thus, we wrap glyphs in a transparent frame
        // of pixel width 2.
        if self.in_type3_glyph {
            let mut padded_image = vec![];
            padded_image.extend(vec![0; (4 * img_width as usize + 16) * 2]);

            for row in rgba_data.chunks_exact(img_width as usize * 4) {
                padded_image.extend([0; 8]);
                padded_image.extend(row);
                padded_image.extend([0; 8]);
            }

            padded_image.extend(vec![0; (4 * img_width as usize + 16) * 2]);
            img_width += 4;
            img_height += 4;
            additional_transform *= Affine::translate((-2.0, -2.0));
            may_have_transparency = true;

            rgba_data = padded_image;
        }

        let pixmap = Pixmap::from_parts(
            rgba_data,
            img_width as u16,
            img_height as u16,
            PixelMetadata {
                may_have_transparency,
                ..Default::default()
            },
        );

        self.draw_pixmap(
            Arc::new(pixmap),
            quality,
            cur_transform * additional_transform,
        );
    }

    pub(crate) fn draw_pixmap(&mut self, pixmap: Arc<Pixmap>, quality: ImageQuality, transform: Affine) {
        let (width, height) = (pixmap.width(), pixmap.height());
        let image = Image {
            image: ImageSource::Pixmap(pixmap),
            sampler: ImageSampler {
                x_extend: peniko::Extend::Pad,
                y_extend: peniko::Extend::Pad,
                quality,
                alpha: 1.0,
            },
        };

        self.ctx.set_transform(transform);
        self.ctx.set_paint(image);
        self.ctx
            .fill_rect(&Rect::new(0.0, 0.0, width as f64, height as f64));
    }

    pub(super) fn draw_pdf_image<'a>(
        &mut self,
        image: hayro_interpret::Image<'a, '_>,
        props: ImageDrawProps<'a>,
    ) {
        let local = Rect::new(0.0, 0.0, image.width() as f64, image.height() as f64);
        if self.misses_target(props.transform.transform_rect_bbox(local)) {
            return;
        }

        self.apply_image_props(&props);
        let mut transform = props.transform;
        self.ctx.set_paint_transform(Affine::IDENTITY);
        self.ctx.set_aliasing_threshold(Some(1));

        let (x, y) = x_y_advances(&transform);
        let target_width = (x.length() * image.width() as f64).ceil() as u32;
        let target_height = (y.length() * image.height() as f64).ceil() as u32;

        match image {
            hayro_interpret::Image::Stencil(s) => {
                s.with_stencil(
                    |stencil, paint| {
                        transform *= Affine::scale_non_uniform(
                            stencil.scale_factors.0 as f64,
                            stencil.scale_factors.1 as f64,
                        );

                        match paint {
                            Paint::Color(c) => {
                                let color = c.to_rgba().to_rgba8();
                                let alpha = color[3];

                                let blend_mode = self.ctx.blend_mode();
                                let push_layer =
                                    alpha != 255 || blend_mode != peniko::BlendMode::default();
                                self.ctx.set_transform(transform);
                                if push_layer {
                                    self.ctx.push_layer(
                                        None,
                                        Some(blend_mode),
                                        Some(alpha as f32 / 255.0),
                                        None,
                                        None,
                                    );
                                }
                                let old_rule = *self.ctx.fill_rule();
                                self.ctx.set_fill_rule(Fill::NonZero);

                                self.draw_image(
                                    RenderImageData::Solid(SolidColorImage {
                                        color: [color[0], color[1], color[2]],
                                        width: stencil.width,
                                        height: stencil.height,
                                        interpolate: stencil.interpolate,
                                    }),
                                    Some(stencil),
                                );

                                if push_layer {
                                    self.ctx.pop_layer();
                                }

                                self.ctx.set_fill_rule(old_rule);
                            }
                            Paint::Pattern(_) => {
                                let (width, height) = (self.ctx.width(), self.ctx.height());
                                let stencil_rect = Rect::new(
                                    0.0,
                                    0.0,
                                    stencil.width as f64,
                                    stencil.height as f64,
                                );
                                let mask_pix = {
                                    let rgb_bytes = ImageData::Rgb(RgbData {
                                        data: vec![
                                            255;
                                            stencil.width as usize
                                                * stencil.height as usize
                                                * 3
                                        ],
                                        width: stencil.width,
                                        height: stencil.height,
                                        interpolate: stencil.interpolate,
                                        scale_factors: stencil.scale_factors,
                                    });
                                    let mut ctx = self.child_context(width, height);
                                    let mut sub_renderer = Renderer::new(&mut ctx, self.global);
                                    let mut sub_pix = Pixmap::new(width, height);
                                    sub_renderer.ctx.set_transform(transform);
                                    sub_renderer.draw_image(rgb_bytes, Some(stencil));
                                    sub_renderer.ctx.flush();
                                    let mut resources = vello_cpu::Resources::default();
                                    sub_renderer.ctx.render(&mut sub_pix, &mut resources);
                                    sub_pix
                                };

                                self.ctx.push_layer(
                                    None,
                                    Some(self.ctx.blend_mode()),
                                    None,
                                    Some(Mask::new_luminance(&mask_pix)),
                                    None,
                                );
                                self.ctx.set_transform(transform);

                                let clip_path = self.set_paint(paint, || stencil_rect, false);
                                if let Some(clip_path) = clip_path.as_ref() {
                                    self.push_clip_path_inner(clip_path, FillRule::NonZero);
                                }
                                self.ctx.fill_rect(&stencil_rect);
                                if clip_path.is_some() {
                                    self.ctx.pop_clip();
                                }

                                self.ctx.pop_layer();
                            }
                        };
                    },
                    Some((target_width, target_height)),
                );
            }
            hayro_interpret::Image::Raster(r) => {
                r.with_rgba(
                    |image, alpha| {
                        let (sx, sy) = image.scale_factors();
                        transform *= Affine::scale_non_uniform(sx as f64, sy as f64);
                        self.ctx.set_transform(transform);
                        self.draw_image(image, alpha);
                    },
                    Some((target_width, target_height)),
                );
            }
        }

        self.ctx
            .set_aliasing_threshold(self.global.aliasing_threshold);
    }
}

trait Splat4thExt {
    fn splat_4th(self) -> Self;
}

impl<S: Simd> Splat4thExt for u8x32<S> {
    #[inline(always)]
    fn splat_4th(self) -> Self {
        [
            self[3], self[3], self[3], self[3], self[7], self[7], self[7], self[7], self[11],
            self[11], self[11], self[11], self[15], self[15], self[15], self[15], self[19],
            self[19], self[19], self[19], self[23], self[23], self[23], self[23], self[27],
            self[27], self[27], self[27], self[31], self[31], self[31], self[31],
        ]
        .simd_into(self.simd)
    }
}

fn resize_image_data(
    scaler: &Scaler,
    data: Vec<u8>,
    src_width: u32,
    src_height: u32,
    new_width: u32,
    new_height: u32,
    pixel_format: ImagePixelFormat,
) -> Option<Vec<u8>> {
    match pixel_format {
        ImagePixelFormat::Luma => resize_image_data_impl::<1>(
            scaler,
            data,
            src_width,
            src_height,
            new_width,
            new_height,
            |scaler, source_size, target_size| {
                scaler.plan_planar_resampling(source_size, target_size)
            },
        ),
        ImagePixelFormat::Rgb => resize_image_data_impl::<3>(
            scaler,
            data,
            src_width,
            src_height,
            new_width,
            new_height,
            |scaler, source_size, target_size| scaler.plan_rgb_resampling(source_size, target_size),
        ),
        ImagePixelFormat::PremultipliedRgba => resize_image_data_impl::<4>(
            scaler,
            data,
            src_width,
            src_height,
            new_width,
            new_height,
            |scaler, source_size, target_size| {
                scaler.plan_rgba_resampling(source_size, target_size, false)
            },
        ),
    }
}

/// PdfCraft patch: image dimensions are validated and buffers reserved before planning; any
/// failure hands the owned input back, so the caller may retain the valid source.
fn resize_image_data_impl<const N: usize>(
    scaler: &Scaler,
    data: Vec<u8>,
    src_width: u32,
    src_height: u32,
    new_width: u32,
    new_height: u32,
    plan: impl FnOnce(&Scaler, ImageSize, ImageSize) -> Result<Arc<Resampling<u8, N>>, PicScaleError>,
) -> Option<Vec<u8>> {
    let resized = (|| {
        let len = resampling_byte_len(
            (src_width, src_height),
            (new_width, new_height),
            N,
            data.len(),
            MAX_IMAGE_PIXELS,
        )?;
        let source_size = ImageSize::new(src_width as usize, src_height as usize);
        let target_size = ImageSize::new(new_width as usize, new_height as usize);
        let src =
            ImageStore::<u8, N>::from_slice(&data, src_width as usize, src_height as usize).ok()?;
        let plan = plan(scaler, source_size, target_size).ok()?;
        let mut out = image_buffer(len)?;
        out.resize(len, 0);
        let mut dst = ImageStoreMut::<u8, N>::from_slice(
            &mut out,
            new_width as usize,
            new_height as usize,
        )
        .ok()?;
        plan.resample(&src, &mut dst).ok()?;
        Some(out)
    })();
    resized
}

fn downscale_with_alpha<const N: usize>(
    level: Level,
    scaler: &Scaler,
    mut color: Vec<u8>,
    alpha: Vec<u8>,
    src_width: u32,
    src_height: u32,
    new_width: u32,
    new_height: u32,
) -> Option<Vec<u8>> {
    premultiply_with_alpha::<N>(level, &mut color, &alpha);

    let color_format = if N == 1 {
        ImagePixelFormat::Luma
    } else {
        ImagePixelFormat::Rgb
    };
    let color = resize_image_data(
        scaler,
        color,
        src_width,
        src_height,
        new_width,
        new_height,
        color_format,
    )?;
    let alpha = resize_image_data(
        scaler,
        alpha,
        src_width,
        src_height,
        new_width,
        new_height,
        ImagePixelFormat::Luma,
    )?;

    Some(
        color
            .as_chunks::<N>()
            .0
            .iter()
            .zip(alpha)
            .flat_map(|(pixel, a)| match pixel.as_slice() {
                [g] => [*g, *g, *g, a],
                [r, g, b] => [*r, *g, *b, a],
                _ => unreachable!(),
            })
            .collect(),
    )
}

fn premultiply_with_alpha<const N: usize>(level: Level, color: &mut [u8], alpha: &[u8]) {
    let simd_pixels = alpha.len() / 32 * 32;
    let (simd_color, tail_color) = color.split_at_mut(simd_pixels * N);
    let (simd_alpha, tail_alpha) = alpha.split_at(simd_pixels);

    #[inline(always)]
    fn premultiply_luma_simd<S: Simd>(simd: S, luma: &mut [u8], alpha: &[u8]) {
        for (luma, alpha) in luma
            .as_chunks_mut::<32>()
            .0
            .iter_mut()
            .zip(alpha.as_chunks::<32>().0)
        {
            let premultiplied = normalized_mul(
                simd,
                u8x32::from_slice(simd, luma),
                u8x32::from_slice(simd, alpha),
            );
            premultiplied.store_slice(luma);
        }
    }

    #[inline(always)]
    fn premultiply_rgb_simd<S: Simd>(simd: S, rgb: &mut [u8], alpha: &[u8]) {
        const ALPHA_INDICES: [[u8; 16]; 3] = [
            [0, 0, 0, 1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4, 5],
            [5, 5, 6, 6, 6, 7, 7, 7, 8, 8, 8, 9, 9, 9, 10, 10],
            [
                10, 11, 11, 11, 12, 12, 12, 13, 13, 13, 14, 14, 14, 15, 15, 15,
            ],
        ];

        let [i0, i1, i2] = ALPHA_INDICES;
        let (i0, i1, i2) = (i0.simd_into(simd), i1.simd_into(simd), i2.simd_into(simd));

        for (rgb, alpha) in rgb
            .as_chunks_mut::<96>()
            .0
            .iter_mut()
            .zip(alpha.as_chunks::<32>().0)
        {
            let (low, high) = simd.split_u8x32(u8x32::from_slice(simd, alpha));
            let alphas = [
                simd.combine_u8x16(
                    simd.swizzle_dyn_u8x16(low, i0),
                    simd.swizzle_dyn_u8x16(low, i1),
                ),
                simd.combine_u8x16(
                    simd.swizzle_dyn_u8x16(low, i2),
                    simd.swizzle_dyn_u8x16(high, i0),
                ),
                simd.combine_u8x16(
                    simd.swizzle_dyn_u8x16(high, i1),
                    simd.swizzle_dyn_u8x16(high, i2),
                ),
            ];

            for (rgb, alpha) in rgb.as_chunks_mut::<32>().0.iter_mut().zip(alphas) {
                normalized_mul(simd, u8x32::from_slice(simd, rgb), alpha).store_slice(rgb);
            }
        }
    }

    if N == 1 {
        fearless_simd::dispatch!(level, simd => premultiply_luma_simd(simd, simd_color, simd_alpha));
    } else {
        fearless_simd::dispatch!(level, simd => premultiply_rgb_simd(simd, simd_color, simd_alpha));
    }

    for (pixel, a) in tail_color.as_chunks_mut::<N>().0.iter_mut().zip(tail_alpha) {
        for c in pixel {
            *c = div_255(u16::from(*c) * u16::from(*a)) as u8;
        }
    }
}

fn premultiply_rgba(level: Level, data: &mut [u8]) {
    let simd_len = data.len() / 32 * 32;
    let (simd_data, tail) = data.split_at_mut(simd_len);

    #[inline(always)]
    fn premultiply_rgba_simd<S: Simd>(simd: S, data: &mut [u8]) {
        let alpha_lanes: mask8x32<S> = [
            0, 0, 0, -1, 0, 0, 0, -1, 0, 0, 0, -1, 0, 0, 0, -1, 0, 0, 0, -1, 0, 0, 0, -1, 0, 0, 0,
            -1, 0, 0, 0, -1,
        ]
        .simd_into(simd);
        for chunk in data.as_chunks_mut::<32>().0 {
            let rgba = u8x32::from_slice(simd, chunk);
            let premultiplied = normalized_mul(simd, rgba, rgba.splat_4th());
            alpha_lanes.select(rgba, premultiplied).store_slice(chunk);
        }
    }

    fearless_simd::dispatch!(level, simd => premultiply_rgba_simd(simd, simd_data));

    for pixel in tail.as_chunks_mut::<4>().0 {
        let alpha = u16::from(pixel[3]);
        pixel[0] = div_255(u16::from(pixel[0]) * alpha) as u8;
        pixel[1] = div_255(u16::from(pixel[1]) * alpha) as u8;
        pixel[2] = div_255(u16::from(pixel[2]) * alpha) as u8;
    }
}

#[inline(always)]
fn normalized_mul<S: Simd>(simd: S, a: u8x32<S>, b: u8x32<S>) -> u8x32<S> {
    let (a_low, a_high) = simd.widen_u8x32(a);
    let (b_low, b_high) = simd.widen_u8x32(b);
    simd.narrow_u16x16((a_low * b_low).div_255(), (a_high * b_high).div_255())
}

trait Div255Ext {
    fn div_255(self) -> Self;
}

impl<S: Simd> Div255Ext for u16x16<S> {
    #[inline(always)]
    fn div_255(self) -> Self {
        (self + Self::splat(self.simd, 255)) >> 8
    }
}

#[inline(always)]
const fn div_255(value: u16) -> u16 {
    (value + 255) >> 8
}
