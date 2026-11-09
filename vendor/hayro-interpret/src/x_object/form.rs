use super::xobject_oc;
use crate::cmyk_group::{CmykGroup, blends_in_device_cmyk, names_a_blending_space};
use crate::context::Context;
use crate::device::Device;
use crate::{ClipPath, FillRule, interpret};
use hayro_syntax::content::TypedIter;
use hayro_syntax::object::Dict;
use hayro_syntax::object::Stream;
use hayro_syntax::object::dict::keys::*;
use hayro_syntax::page::Resources;
use kurbo::{Affine, Rect, Shape};

pub(crate) struct FormXObject<'a> {
    pub(crate) decoded: crate::context::DecodedContent<'a>,
    pub(crate) matrix: Affine,
    pub(crate) bbox: [f32; 4],
    is_transparency_group: bool,
    isolated: bool,
    knockout: bool,
    pub(crate) dict: Dict<'a>,
    pub(crate) resources: Option<Resources<'a>>,
}

impl<'a> FormXObject<'a> {
    pub(crate) fn new(stream: &Stream<'a>) -> Option<Self> {
        let dict = stream.dict();

        // PdfCraft patch: see `context::MAX_PAGE_CONTENT`.
        let decoded = crate::context::decode_content(stream)?;
        let resources = dict.get(RESOURCES).map(Resources::new);

        let matrix = Affine::new(
            dict.get::<[f64; 6]>(MATRIX)
                .unwrap_or([1.0, 0.0, 0.0, 1.0, 0.0, 0.0]),
        );
        let bbox = dict.get::<[f32; 4]>(BBOX)?;
        let group = dict.get::<Dict<'_>>(GROUP);
        let flag = |key| {
            group
                .as_ref()
                .and_then(|g| g.get::<bool>(key))
                .unwrap_or(false)
        };

        Some(Self {
            decoded,
            matrix,
            is_transparency_group: group.is_some(),
            isolated: flag(I),
            knockout: flag(K),
            bbox,
            dict: dict.clone(),
            resources,
        })
    }

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
}
