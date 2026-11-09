use std::borrow::Cow;
use hayro_syntax::object::Stream;
use crate::cache::{Cache, CacheKey};
use crate::color::{Color, ColorSpace};
use crate::convert::convert_transform;
use crate::font::{FontCache, PositionedGlyph, StandardFont};
use crate::interpret::stack::StateStack;
use crate::interpret::state::{
    ClipType, CompositingState, PaintData, PaintState, State, TextState, TextStateFont,
};
use crate::ocg::OcgState;
use crate::util::{BezPathExt, Float64Ext};
use crate::{
    ClipPath, Device, DrawProps, FillRule, InterpreterSettings, InterpreterWarning, Paint,
    StrokeProps,
};
use hayro_syntax::content::ops::Transform;
use hayro_syntax::object::Dict;
use hayro_syntax::object::Name;
use hayro_syntax::page::Resources;
use hayro_syntax::xref::XRef;
use kurbo::{Affine, BezPath, PathEl, Point, Rect, Shape};
use smallvec::smallvec;

/// Maximum nesting depth for interpreting `XObject`'s/patterns/streams.
pub(crate) const MAX_NESTED_INTERPRETATION_DEPTH: u32 = 50;
std::thread_local! {
    static CONTENT_LEFT: core::cell::Cell<usize> = const { core::cell::Cell::new(MAX_PAGE_CONTENT) };
    // PdfCraft patch: whether this page's content budget ran out.
    static CONTENT_TRUNCATED: core::cell::Cell<bool> = const { core::cell::Cell::new(false) };
}

/// PdfCraft patch: decoded content-stream bytes one page may hold at once: its contents, and
/// the forms, tiling patterns and Type 3 glyph procedures being painted. Each stream already
/// stops at `MAX_DECODED_STREAM`, but forms nest fifty deep and every level keeps its content
/// while the next one runs. A level gives its bytes back when it finishes, so content painted
/// again and again doesn't use the budget up. Ordinary pages hold a few megabytes.
pub(crate) const MAX_PAGE_CONTENT: usize = 512 << 20;

/// PdfCraft patch: a content stream decoded within the page's [`MAX_PAGE_CONTENT`]. The bytes
/// it decoded count against the budget until it is dropped; a stream without filters is
/// borrowed from the file and costs nothing.
pub(crate) struct DecodedContent<'a> {
    data: Cow<'a, [u8]>,
    held: usize,
}

impl AsRef<[u8]> for DecodedContent<'_> {
    fn as_ref(&self) -> &[u8] {
        &self.data
    }
}

impl Drop for DecodedContent<'_> {
    fn drop(&mut self) {
        let held = self.held;
        CONTENT_LEFT.with(|c| c.set(c.get().saturating_add(held).min(MAX_PAGE_CONTENT)));
    }
}

/// PdfCraft patch: decode a content stream within what is left of this page's
/// [`MAX_PAGE_CONTENT`]. `None` once the budget is spent (the content is skipped).
pub(crate) fn decode_content<'a>(stream: &Stream<'a>) -> Option<DecodedContent<'a>> {
    let left = CONTENT_LEFT.with(core::cell::Cell::get);
    if left == 0 {
        CONTENT_TRUNCATED.with(|truncated| truncated.set(true));
        return None;
    }
    let data = stream
        .decoded_within(&hayro_syntax::object::stream::ImageDecodeParams::default(), left)
        .ok()?
        .data;
    let held = match &data {
        Cow::Owned(v) => v.len(),
        Cow::Borrowed(_) => 0,
    };
    charge_content(held);
    Some(DecodedContent { data, held })
}

/// PdfCraft patch: count `len` bytes held for the rest of the page (its own contents).
pub(crate) fn charge_content(len: usize) {
    CONTENT_LEFT.with(|c| c.set(c.get().saturating_sub(len)));
}

/// PdfCraft patch: whether content was skipped because the page's budget ran out.
pub(crate) fn content_was_truncated() -> bool {
    CONTENT_TRUNCATED.with(core::cell::Cell::get)
}


/// A cache used by the interpreter.
///
/// Ideally, such a cache should be constructed once per PDF and then reused across
/// multiple interpreter invocations on the same document.
// TODO: Cache keys are only unique within a document (streams, for example, are keyed by
// their object identifier), so a cache shared across documents can mix up their objects.
#[derive(Clone)]
pub struct InterpreterCache<'a> {
    pub(crate) font_cache: FontCache<'a>,
    pub(crate) object_cache: Cache,
}

impl<'a> Default for InterpreterCache<'a> {
    fn default() -> Self {
        Self::new()
    }
}

impl<'a> InterpreterCache<'a> {
    /// Create a new interpreter cache.
    pub fn new() -> Self {
        Self {
            font_cache: FontCache::new(),
            object_cache: Cache::new(),
        }
    }
}

/// A per-page interpretation context that borrows shared data from an [`InterpreterCache`].
pub struct Context<'a> {
    stack: StateStack<'a>,
    path: BezPath,
    sub_path_start: Point,
    last_point: Point,
    clip: Option<FillRule>,
    root_transforms: Vec<Affine>,
    bbox: Vec<Rect>,
    pub(crate) settings: InterpreterSettings,
    pub(crate) interpreter_cache: InterpreterCache<'a>,
    pub(crate) xref: &'a XRef,
    pub(crate) ocg_state: OcgState,
    nesting_depth: u32,
    pub(crate) glyph_scratch: Vec<PositionedGlyph<'a>>,
}

impl<'a> Context<'a> {
    /// Create a new context.
    pub fn new(
        initial_transform: Affine,
        bbox: Rect,
        cache: &InterpreterCache<'a>,
        xref: &'a XRef,
        settings: InterpreterSettings,
    ) -> Self {
        let state = State::new(initial_transform);
        // PdfCraft patch: a page (or other top-level content) starts with a fresh content budget.
        CONTENT_LEFT.with(|c| c.set(MAX_PAGE_CONTENT));
        CONTENT_TRUNCATED.with(|truncated| truncated.set(false));

        Self::new_with(initial_transform, bbox, cache, xref, settings, state, 0)
    }

    pub(crate) fn new_with(
        initial_transform: Affine,
        bbox: Rect,
        cache: &InterpreterCache<'a>,
        xref: &'a XRef,
        settings: InterpreterSettings,
        state: State<'a>,
        nesting_depth: u32,
    ) -> Self {
        let ocg_state = {
            let root_ref = xref.root_id();
            xref.get::<Dict<'_>>(root_ref)
                .map(|catalog| OcgState::from_catalog(&catalog))
                .unwrap_or_default()
        };

        Self {
            stack: StateStack::new(state),
            settings,
            xref,
            root_transforms: vec![initial_transform],
            last_point: Point::default(),
            sub_path_start: Point::default(),
            clip: None,
            bbox: vec![bbox],
            path: BezPath::new(),
            interpreter_cache: cache.clone(),
            ocg_state,
            nesting_depth,
            glyph_scratch: Vec::new(),
        }
    }

    pub(crate) fn save_state(&mut self) {
        self.stack.save();
    }

    pub(crate) fn bbox(&self) -> Rect {
        self.bbox
            .last()
            .copied()
            .unwrap_or_else(|| Rect::new(0.0, 0.0, 1.0, 1.0))
    }

    fn push_bbox(&mut self, bbox: Rect) {
        let new = self.bbox().intersect(bbox);
        self.bbox.push(new);
    }

    pub(crate) fn push_clip_path(
        &mut self,
        clip_path: BezPath,
        fill: FillRule,
        device: &mut impl Device<'a>,
    ) {
        if let Some(clip_rect) = path_as_rect(&clip_path) {
            let cur_bbox = self.bbox();

            // If the clip path is a rect and completely covers the current bbox, don't emit it.
            if cur_bbox
                .min_x()
                .is_nearly_greater_or_equal(clip_rect.min_x())
                && cur_bbox
                    .min_y()
                    .is_nearly_greater_or_equal(clip_rect.min_y())
                && cur_bbox.max_x().is_nearly_less_or_equal(clip_rect.max_x())
                && cur_bbox.max_y().is_nearly_less_or_equal(clip_rect.max_y())
            {
                self.stack.push_clip(ClipType::Dummy);
                return;
            }

            device.push_clip_rect(&clip_rect);
            self.push_bbox(clip_rect);
            self.stack.push_clip(ClipType::Real);
            return;
        }

        let bbox = clip_path.bounding_box();
        device.push_clip_path(&ClipPath {
            path: clip_path,
            fill,
        });
        self.push_bbox(bbox);
        self.stack.push_clip(ClipType::Real);
    }

    pub(crate) fn pop_clip(&mut self, device: &mut impl Device<'a>) {
        if let Some(ClipType::Real) = self.stack.pop_clip() {
            device.pop_clip();
            self.pop_bbox();
        }
    }

    fn pop_bbox(&mut self) {
        self.bbox.pop();
    }

    pub(crate) fn push_root_transform(&mut self) {
        self.root_transforms.push(self.ctm());
    }

    pub(crate) fn pop_root_transform(&mut self) {
        self.root_transforms.pop();
    }

    pub(crate) fn root_transform(&self) -> Affine {
        self.root_transforms
            .last()
            .copied()
            .unwrap_or(Affine::IDENTITY)
    }

    pub(crate) fn restore_state(&mut self, device: &mut impl Device<'a>) {
        let Some(target_clips) = self.stack.restore_target_clips() else {
            return;
        };

        while self.stack.num_clips() > target_clips {
            self.pop_clip(device);
        }

        self.stack.restore();
    }

    pub(crate) fn draw_props(&self, is_stroke: bool) -> DrawProps<'a> {
        DrawProps {
            transform: self.ctm(),
            paint: self.get_paint(is_stroke),
            soft_mask: self.compositing().soft_mask.clone(),
            blend_mode: self.compositing().blend_mode,
        }
    }

    pub(crate) fn get_paint(&self, is_stroke: bool) -> Paint<'a> {
        let paint = if is_stroke {
            self.stroke_paint()
        } else {
            self.fill_paint()
        };
        let data = PaintData {
            alpha: paint.alpha,
            color: paint.color.clone(),
            color_space: paint.color_space.clone(),
            pattern: paint.pattern.clone(),
            transfer_function: self.compositing().transfer_function.clone(),
        };

        if data.color_space.is_pattern() || data.pattern.is_some() {
            if let Some(mut pattern) = data.pattern {
                pattern.set_opacity(data.alpha);

                if let Some(tf) = &data.transfer_function {
                    pattern.set_transfer_function(tf.clone());
                }

                pattern.pre_concat_transform(self.root_transform());

                Paint::Pattern(Box::new(pattern))
            } else {
                // Pattern was likely invalid, use transparent paint.
                Paint::Color(Color::new(ColorSpace::device_gray(), smallvec![0.0], 0.0))
            }
        } else {
            let color = Color::new(data.color_space, data.color, data.alpha);

            if let Some(tf) = &data.transfer_function {
                Paint::Color(Color::from_rgba(tf.apply(&color.to_rgba())))
            } else {
                Paint::Color(color)
            }
        }
    }

    pub(crate) fn path(&self) -> &BezPath {
        &self.path
    }

    pub(crate) fn path_mut(&mut self) -> &mut BezPath {
        &mut self.path
    }

    pub(crate) fn sub_path_start(&self) -> &Point {
        &self.sub_path_start
    }

    pub(crate) fn sub_path_start_mut(&mut self) -> &mut Point {
        &mut self.sub_path_start
    }

    pub(crate) fn last_point(&self) -> &Point {
        &self.last_point
    }

    pub(crate) fn last_point_mut(&mut self) -> &mut Point {
        &mut self.last_point
    }

    pub(crate) fn clip_mut(&mut self) -> &mut Option<FillRule> {
        &mut self.clip
    }

    pub(crate) fn ctm(&self) -> Affine {
        *self.stack.ctm()
    }

    pub(crate) fn ctm_mut(&mut self) -> &mut Affine {
        self.stack.ctm_mut()
    }

    pub(crate) fn stroke_paint(&self) -> &PaintState<'a> {
        self.stack.stroke()
    }

    pub(crate) fn stroke_paint_mut(&mut self) -> &mut PaintState<'a> {
        self.stack.stroke_mut()
    }

    pub(crate) fn fill_paint(&self) -> &PaintState<'a> {
        self.stack.fill()
    }

    pub(crate) fn fill_paint_mut(&mut self) -> &mut PaintState<'a> {
        self.stack.fill_mut()
    }

    pub(crate) fn stroke_props(&self) -> &StrokeProps {
        self.stack.stroke_props()
    }

    pub(crate) fn stroke_props_mut(&mut self) -> &mut StrokeProps {
        self.stack.stroke_props_mut()
    }

    pub(crate) fn text_state(&self) -> &TextState<'a> {
        self.stack.text()
    }

    pub(crate) fn text_state_mut(&mut self) -> &mut TextState<'a> {
        self.stack.text_mut()
    }

    pub(crate) fn compositing(&self) -> &CompositingState<'a> {
        self.stack.compositing()
    }

    pub(crate) fn compositing_mut(&mut self) -> &mut CompositingState<'a> {
        self.stack.compositing_mut()
    }

    /// A copy of the current graphics state.
    pub(crate) fn snapshot(&self) -> State<'a> {
        self.stack.snapshot()
    }

    pub(crate) fn pre_concat_transform(&mut self, transform: Transform) {
        self.pre_concat_affine(convert_transform(transform));
    }

    pub(crate) fn pre_concat_affine(&mut self, transform: Affine) {
        *self.ctm_mut() *= transform;
    }

    pub(crate) fn get_color_space(
        &mut self,
        resources: &Resources<'_>,
        name: &Name<'_>,
    ) -> Option<ColorSpace> {
        let cs_object = resources.get_color_space(name)?;
        self.interpreter_cache
            .object_cache
            .get_or_insert_with(cs_object.cache_key(), || {
                ColorSpace::new(cs_object.clone(), &self.interpreter_cache.object_cache)
            })
    }

    pub(crate) fn num_states(&self) -> usize {
        self.stack.len()
    }

    pub(crate) fn nesting_depth(&self) -> u32 {
        self.nesting_depth
    }

    /// PdfCraft patch: a copy of the current graphics state (for re-interpreting a form's
    /// contents on their own, see `cmyk_group`).
    pub(crate) fn snapshot_state(&self) -> State<'a> {
        self.stack.snapshot()
    }


    pub(crate) fn begin_nested_interpretation(&mut self) -> bool {
        if self.nesting_depth >= MAX_NESTED_INTERPRETATION_DEPTH {
            return false;
        }

        self.nesting_depth += 1;

        true
    }

    pub(crate) fn end_nested_interpretation(&mut self) {
        self.nesting_depth = self.nesting_depth.saturating_sub(1);
    }

    pub(crate) fn fallback_font(&self) -> Option<TextStateFont<'a>> {
        self.interpreter_cache
            .font_cache
            .standard_font(StandardFont::Helvetica, &self.settings.font_resolver)
            .map(TextStateFont::Fallback)
    }

    pub(crate) fn resolve_font(&mut self, font_dict: &Dict<'a>) -> Option<TextStateFont<'a>> {
        let resolved = self.interpreter_cache.font_cache.font(
            font_dict,
            &self.settings.font_resolver,
            &self.settings.cmap_resolver,
        );

        if let Some(resolved) = resolved {
            Some(TextStateFont::Font(resolved))
        } else {
            (self.settings.warning_sink)(InterpreterWarning::UnsupportedFont);

            self.fallback_font()
        }
    }
}

pub(crate) fn path_as_rect(path: &BezPath) -> Option<Rect> {
    let points = match path.elements() {
        [
            PathEl::MoveTo(first),
            PathEl::LineTo(second),
            PathEl::LineTo(third),
            PathEl::LineTo(fourth),
            PathEl::ClosePath,
        ] => [*first, *second, *third, *fourth, *first],
        [
            PathEl::MoveTo(first),
            PathEl::LineTo(second),
            PathEl::LineTo(third),
            PathEl::LineTo(fourth),
            PathEl::LineTo(last),
        ] if first.x.is_nearly_equal(last.x) && first.y.is_nearly_equal(last.y) => {
            [*first, *second, *third, *fourth, *last]
        }
        _ => return None,
    };

    let mut previous_axis = None;
    for edge in points.windows(2) {
        let same_x = edge[0].x.is_nearly_equal(edge[1].x);
        let same_y = edge[0].y.is_nearly_equal(edge[1].y);
        if same_x == same_y || previous_axis == Some(same_x) {
            return None;
        }
        previous_axis = Some(same_x);
    }

    Some(path.fast_bounding_box())
}
