use std::collections::hash_map::Entry;

use eframe::egui::{
    self, epaint::CornerRadius, Align2, Color32, FontFamily, FontId, Rect, Stroke, StrokeKind,
};

use crate::{
    app::{CanvasState, ChangeHistory, ImageMoveDrag, ImageResizeDrag, ResizeHandle},
    document::{
        DocumentImage, DocumentState, ImageLayoutMode, ImageRendering, MIN_IMAGE_SIZE_POINTS,
    },
    layout::document_points_to_screen_points,
};

use super::{palette, ImageLayout};

pub(super) fn paint_image_on_page(
    ui: &mut egui::Ui,
    canvas: &mut CanvasState,
    painter: &egui::Painter,
    image: &ImageLayout,
    image_rect: Rect,
    palette: &palette::CanvasPalette,
    alpha_multiplier: f32,
) {
    if let Some(texture) = texture_for_image(ui.ctx(), canvas, &image.image) {
        let alpha = (image.image.opacity * alpha_multiplier * 255.0)
            .round()
            .clamp(0.0, 255.0) as u8;
        let tint = Color32::from_white_alpha(alpha);
        painter.image(
            texture.id(),
            image_rect,
            Rect::from_min_max(egui::Pos2::ZERO, egui::pos2(1.0, 1.0)),
            tint,
        );
    } else {
        painter.rect_filled(image_rect, CornerRadius::same(4), palette.footer_bg);
        painter.rect_stroke(
            image_rect,
            CornerRadius::same(4),
            Stroke::new(1.0, palette.footer_stroke),
            StrokeKind::Outside,
        );
        painter.text(
            image_rect.center(),
            Align2::CENTER_CENTER,
            &image.image.alt_text,
            FontId::new(12.0, FontFamily::Proportional),
            palette.footer_text,
        );
    }
}

pub(super) fn resize_handle_rects(image_rect: Rect) -> [(ResizeHandle, Rect); 8] {
    const H: f32 = 5.0;
    let sq =
        |x: f32, y: f32| Rect::from_center_size(egui::pos2(x, y), egui::vec2(H * 2.0, H * 2.0));
    let cx = image_rect.center().x;
    let cy = image_rect.center().y;
    [
        (ResizeHandle::NW, sq(image_rect.left(), image_rect.top())),
        (ResizeHandle::N, sq(cx, image_rect.top())),
        (ResizeHandle::NE, sq(image_rect.right(), image_rect.top())),
        (ResizeHandle::E, sq(image_rect.right(), cy)),
        (
            ResizeHandle::SE,
            sq(image_rect.right(), image_rect.bottom()),
        ),
        (ResizeHandle::S, sq(cx, image_rect.bottom())),
        (ResizeHandle::SW, sq(image_rect.left(), image_rect.bottom())),
        (ResizeHandle::W, sq(image_rect.left(), cy)),
    ]
}

pub(super) fn paint_image_selection(painter: &egui::Painter, image_rect: Rect) {
    const SELECTION_COLOR: Color32 = Color32::from_rgb(54, 116, 206);
    painter.rect_stroke(
        image_rect,
        CornerRadius::ZERO,
        Stroke::new(2.0, SELECTION_COLOR),
        StrokeKind::Outside,
    );
    for (_, handle_rect) in &resize_handle_rects(image_rect) {
        painter.rect_filled(*handle_rect, CornerRadius::ZERO, Color32::WHITE);
        painter.rect_stroke(
            *handle_rect,
            CornerRadius::ZERO,
            Stroke::new(1.5, SELECTION_COLOR),
            StrokeKind::Outside,
        );
    }
}

pub(super) fn image_drag_preview_rect(
    image_id: usize,
    image_rect: Rect,
    move_preview: Option<(usize, Rect, egui::Vec2)>,
) -> Rect {
    match move_preview {
        Some((dragged_id, start_rect, offset)) if dragged_id == image_id => {
            start_rect.translate(offset)
        }
        _ => image_rect,
    }
}

pub(super) fn handle_image_interaction(
    ui: &mut egui::Ui,
    response: &egui::Response,
    canvas: &mut CanvasState,
    document: &mut DocumentState,
    history: &mut ChangeHistory,
) -> (bool, bool) {
    const HANDLE_HIT_PADDING: f32 = 6.0;
    const DRAG_THRESHOLD_POINTS: f32 = 2.0;
    let mut document_changed = false;

    if response.clicked() || response.drag_started() {
        response.request_focus();
    }

    // Change cursor icon when hovering over a resize handle
    if let Some(hover_pos) = ui.ctx().pointer_hover_pos() {
        let mut cursor_icon = None;
        if let Some((_, handle)) = image_handle_hit(canvas, hover_pos, HANDLE_HIT_PADDING) {
            cursor_icon = Some(match handle {
                ResizeHandle::NW | ResizeHandle::SE => egui::CursorIcon::ResizeNwSe,
                ResizeHandle::NE | ResizeHandle::SW => egui::CursorIcon::ResizeNeSw,
                ResizeHandle::N | ResizeHandle::S => egui::CursorIcon::ResizeSouth,
                ResizeHandle::E | ResizeHandle::W => egui::CursorIcon::ResizeEast,
            });
        }
        if cursor_icon.is_none()
            && canvas
                .image_rects
                .iter()
                .any(|(_, rect)| rect.contains(hover_pos))
        {
            cursor_icon = Some(egui::CursorIcon::Grab);
        }
        if let Some(icon) = cursor_icon {
            ui.ctx().set_cursor_icon(icon);
        }
    }

    // Finalize any active image drag when mouse released
    if !response.dragged() {
        if let Some(move_drag) = canvas.move_drag.take() {
            let zoom = canvas.zoom.max(0.01);
            let dx = (move_drag.current_ptr.x - move_drag.start_ptr.x) / zoom;
            let dy = (move_drag.current_ptr.y - move_drag.start_ptr.y) / zoom;

            let was_inline = document
                .image_by_id(move_drag.image_id)
                .map(|image| image.layout_mode == ImageLayoutMode::Inline)
                .unwrap_or(false);

            if was_inline {
                if dx.abs() > DRAG_THRESHOLD_POINTS || dy.abs() > DRAG_THRESHOLD_POINTS {
                    // Auto-convert inline image to floating on drag
                    document.set_image_layout_mode(move_drag.image_id, ImageLayoutMode::Floating);
                    document.set_image_offset_by_id(
                        move_drag.image_id,
                        move_drag.start_x_points + dx,
                        move_drag.start_y_points + dy,
                    );
                    document_changed = true;
                }
            } else {
                document.set_image_offset_by_id(
                    move_drag.image_id,
                    move_drag.start_x_points + dx,
                    move_drag.start_y_points + dy,
                );
                document_changed = true;
            }
        }
        canvas.resize_drag = None;
    }

    // Continue active resize drag
    if response.dragged() {
        let drag_data = canvas.resize_drag.as_ref().map(|d| {
            (
                d.image_id,
                d.handle,
                d.start_ptr,
                d.start_width_points,
                d.start_height_points,
                d.start_x_points,
                d.start_y_points,
            )
        });
        if let Some((image_id, handle, start_ptr, start_w, start_h, start_x, start_y)) = drag_data {
            if let Some(pointer_pos) = response.interact_pointer_pos() {
                let zoom = canvas.zoom;
                let dx = (pointer_pos.x - start_ptr.x) / zoom;
                let dy = (pointer_pos.y - start_ptr.y) / zoom;
                let shift = ui.input(|i| i.modifiers.shift);
                let lock_aspect = document
                    .image_by_id(image_id)
                    .map(|image| image.lock_aspect_ratio)
                    .unwrap_or(false);
                let lock_ratio = lock_aspect ^ shift;

                let (new_w, new_h, offset_dx, offset_dy) =
                    resized_image_geometry(handle, start_w, start_h, dx, dy, lock_ratio);

                document.resize_image_by_id(image_id, new_w, new_h);
                document.set_image_offset_by_id(image_id, start_x + offset_dx, start_y + offset_dy);
                document_changed = true;
            }
            return (true, document_changed);
        }

        if let Some(move_drag) = canvas.move_drag.as_mut() {
            if let Some(pointer_pos) = response.interact_pointer_pos() {
                move_drag.current_ptr = pointer_pos;
                ui.ctx().set_cursor_icon(egui::CursorIcon::Grabbing);
            }
            return (true, document_changed);
        }
    }

    let Some(pointer_pos) = response.interact_pointer_pos() else {
        return (false, document_changed);
    };

    // Start a new resize drag when dragging from a handle
    if response.drag_started() {
        if let Some((image_id, handle)) = image_handle_hit(canvas, pointer_pos, HANDLE_HIT_PADDING)
        {
            if let Some(image) = document.image_by_id(image_id) {
                let now = ui.input(|i| i.time);
                history.checkpoint(document, now);
                canvas.selected_image_id = Some(image_id);
                canvas.active_table_cell = None;
                canvas.move_drag = None;
                canvas.resize_drag = Some(ImageResizeDrag {
                    image_id,
                    handle,
                    start_ptr: pointer_pos,
                    start_width_points: image.width_points,
                    start_height_points: image.height_points,
                    start_x_points: image.offset_x_points(),
                    start_y_points: image.offset_y_points(),
                });
            }
            return (true, document_changed);
        }

        if let Some((image_id, image_rect)) = image_body_hit(canvas, pointer_pos) {
            if image_rect.contains(pointer_pos) {
                let now = ui.input(|i| i.time);
                history.checkpoint(document, now);
                let offset = document
                    .image_by_id(image_id)
                    .map(|image| (image.offset_x_points(), image.offset_y_points()))
                    .unwrap_or((0.0, 0.0));
                canvas.selected_image_id = Some(image_id);
                canvas.active_table_cell = None;
                canvas.resize_drag = None;
                canvas.move_drag = Some(ImageMoveDrag {
                    image_id,
                    start_ptr: pointer_pos,
                    current_ptr: pointer_pos,
                    start_rect: image_rect,
                    start_x_points: offset.0,
                    start_y_points: offset.1,
                });
                return (true, document_changed);
            }
        }

        if let Some(image_id) = canvas
            .image_rects
            .iter()
            .find(|(_, rect)| rect.contains(pointer_pos))
            .map(|(id, _)| *id)
        {
            canvas.selected_image_id = Some(image_id);
            canvas.active_table_cell = None;
            canvas.resize_drag = None;
            canvas.move_drag = None;
            return (true, document_changed);
        }
    }

    // Click on image body → select it; click elsewhere → deselect
    if response.clicked() {
        let hit = canvas
            .image_rects
            .iter()
            .find(|(_, rect)| rect.contains(pointer_pos))
            .map(|(id, _)| *id);
        canvas.selected_image_id = hit;
        if hit.is_some() {
            canvas.active_table_cell = None;
        }
        if hit.is_none() {
            canvas.resize_drag = None;
            canvas.move_drag = None;
            return (false, document_changed);
        }
        return (true, document_changed);
    }

    (false, document_changed)
}

pub(super) fn resized_image_geometry(
    handle: ResizeHandle,
    start_w: f32,
    start_h: f32,
    dx: f32,
    dy: f32,
    lock_ratio: bool,
) -> (f32, f32, f32, f32) {
    let start_w = start_w.max(MIN_IMAGE_SIZE_POINTS);
    let start_h = start_h.max(MIN_IMAGE_SIZE_POINTS);

    if lock_ratio {
        return resized_image_geometry_locked(handle, start_w, start_h, dx, dy);
    }

    let mut left = 0.0;
    let mut top = 0.0;
    let mut right = start_w;
    let mut bottom = start_h;

    match handle {
        ResizeHandle::NW => {
            left = dx;
            top = dy;
        }
        ResizeHandle::N => top = dy,
        ResizeHandle::NE => {
            right = start_w + dx;
            top = dy;
        }
        ResizeHandle::E => right = start_w + dx,
        ResizeHandle::SE => {
            right = start_w + dx;
            bottom = start_h + dy;
        }
        ResizeHandle::S => bottom = start_h + dy,
        ResizeHandle::SW => {
            left = dx;
            bottom = start_h + dy;
        }
        ResizeHandle::W => left = dx,
    }

    if right - left < MIN_IMAGE_SIZE_POINTS {
        if matches!(
            handle,
            ResizeHandle::NW | ResizeHandle::W | ResizeHandle::SW
        ) {
            left = right - MIN_IMAGE_SIZE_POINTS;
        } else {
            right = left + MIN_IMAGE_SIZE_POINTS;
        }
    }
    if bottom - top < MIN_IMAGE_SIZE_POINTS {
        if matches!(
            handle,
            ResizeHandle::NW | ResizeHandle::N | ResizeHandle::NE
        ) {
            top = bottom - MIN_IMAGE_SIZE_POINTS;
        } else {
            bottom = top + MIN_IMAGE_SIZE_POINTS;
        }
    }

    (right - left, bottom - top, left, top)
}

pub(super) fn resized_image_geometry_locked(
    handle: ResizeHandle,
    start_w: f32,
    start_h: f32,
    dx: f32,
    dy: f32,
) -> (f32, f32, f32, f32) {
    let aspect = start_h / start_w.max(1.0);
    let (new_w, new_h) = match handle {
        ResizeHandle::E | ResizeHandle::W => {
            aspect_size_from_width(start_w + signed_dx(handle, dx), aspect)
        }
        ResizeHandle::N | ResizeHandle::S => {
            aspect_size_from_height(start_h + signed_dy(handle, dy), aspect)
        }
        ResizeHandle::NW | ResizeHandle::NE | ResizeHandle::SE | ResizeHandle::SW => {
            let from_w = (start_w + signed_dx(handle, dx)).max(MIN_IMAGE_SIZE_POINTS);
            let from_h = (start_h + signed_dy(handle, dy)).max(MIN_IMAGE_SIZE_POINTS);
            let width_change = (from_w / start_w - 1.0).abs();
            let height_change = (from_h / start_h - 1.0).abs();
            if width_change >= height_change {
                aspect_size_from_width(from_w, aspect)
            } else {
                aspect_size_from_height(from_h, aspect)
            }
        }
    };

    let left = match handle {
        ResizeHandle::NW | ResizeHandle::W | ResizeHandle::SW => start_w - new_w,
        ResizeHandle::N | ResizeHandle::S => (start_w - new_w) * 0.5,
        _ => 0.0,
    };
    let top = match handle {
        ResizeHandle::NW | ResizeHandle::N | ResizeHandle::NE => start_h - new_h,
        ResizeHandle::E | ResizeHandle::W => (start_h - new_h) * 0.5,
        _ => 0.0,
    };

    (new_w, new_h, left, top)
}

fn signed_dx(handle: ResizeHandle, dx: f32) -> f32 {
    if matches!(
        handle,
        ResizeHandle::NW | ResizeHandle::W | ResizeHandle::SW
    ) {
        -dx
    } else {
        dx
    }
}

fn signed_dy(handle: ResizeHandle, dy: f32) -> f32 {
    if matches!(
        handle,
        ResizeHandle::NW | ResizeHandle::N | ResizeHandle::NE
    ) {
        -dy
    } else {
        dy
    }
}

fn aspect_size_from_width(width: f32, aspect: f32) -> (f32, f32) {
    let mut width = width.max(MIN_IMAGE_SIZE_POINTS);
    let mut height = width * aspect.max(0.001);
    if height < MIN_IMAGE_SIZE_POINTS {
        height = MIN_IMAGE_SIZE_POINTS;
        width = height / aspect.max(0.001);
    }
    (width, height)
}

fn aspect_size_from_height(height: f32, aspect: f32) -> (f32, f32) {
    let mut height = height.max(MIN_IMAGE_SIZE_POINTS);
    let mut width = height / aspect.max(0.001);
    if width < MIN_IMAGE_SIZE_POINTS {
        width = MIN_IMAGE_SIZE_POINTS;
        height = width * aspect.max(0.001);
    }
    (width, height)
}

pub(super) fn selected_image_rect(canvas: &CanvasState) -> Option<(usize, Rect)> {
    let selected_id = canvas.selected_image_id?;
    canvas
        .image_rects
        .iter()
        .find(|(id, _)| *id == selected_id)
        .copied()
}

pub(super) fn image_body_hit(
    canvas: &CanvasState,
    pointer_pos: egui::Pos2,
) -> Option<(usize, Rect)> {
    canvas
        .image_rects
        .iter()
        .rev()
        .find(|(_, rect)| rect.contains(pointer_pos))
        .copied()
}

pub(super) fn image_handle_hit(
    canvas: &CanvasState,
    pointer_pos: egui::Pos2,
    padding: f32,
) -> Option<(usize, ResizeHandle)> {
    if let Some((image_id, image_rect)) = selected_image_rect(canvas) {
        for &(handle, handle_rect) in &resize_handle_rects(image_rect) {
            if handle_rect.expand(padding).contains(pointer_pos) {
                return Some((image_id, handle));
            }
        }
    }

    None
}

pub(super) fn image_display_size(image: &DocumentImage, _wrap_width: f32, zoom: f32) -> egui::Vec2 {
    let width =
        document_points_to_screen_points(image.width_points.max(MIN_IMAGE_SIZE_POINTS), zoom);
    let height =
        document_points_to_screen_points(image.height_points.max(MIN_IMAGE_SIZE_POINTS), zoom);
    egui::vec2(width, height)
}

pub(super) fn texture_for_image<'a>(
    ctx: &egui::Context,
    canvas: &'a mut CanvasState,
    image: &DocumentImage,
) -> Option<&'a egui::TextureHandle> {
    // Encode rendering mode into the cache key so smooth/crisp use separate textures
    let cache_key = image.id * 2
        + if image.rendering == ImageRendering::Crisp {
            1
        } else {
            0
        };
    let tex_options = if image.rendering == ImageRendering::Crisp {
        egui::TextureOptions::NEAREST
    } else {
        egui::TextureOptions::LINEAR
    };
    match canvas.image_textures.entry(cache_key) {
        Entry::Occupied(entry) => Some(entry.into_mut()),
        Entry::Vacant(entry) => {
            let decoded = ::image::load_from_memory(&image.bytes).ok()?;
            let rgba = decoded.to_rgba8();
            let size = [rgba.width() as usize, rgba.height() as usize];
            let color_image =
                egui::ColorImage::from_rgba_unmultiplied(size, rgba.as_raw().as_slice());
            let texture = ctx.load_texture(
                format!("doc-image-{}-{}", image.id, cache_key & 1),
                color_image,
                tex_options,
            );
            Some(entry.insert(texture))
        }
    }
}
