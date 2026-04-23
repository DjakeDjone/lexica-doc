mod editor_input;
mod page_layout;
mod palette;

use std::collections::hash_map::Entry;
use std::sync::Arc;

use eframe::egui::{
    self,
    epaint::CornerRadius,
    text_selection::visuals::{paint_text_cursor, paint_text_selection},
    Align2, Color32, FontFamily, FontId, Id, Rect, Sense, Stroke, StrokeKind,
};

use crate::{
    app::{CanvasState, ChangeHistory, ImageMoveDrag, ImageResizeDrag, ResizeHandle, ThemeMode},
    document::{
        text_format, CharacterStyle, DocumentImage, DocumentState, ImageRendering, LineSpacingKind,
        ParagraphAlignment, WrapMode, OBJECT_REPLACEMENT_CHAR,
    },
    layout::{
        centered_page_rect, document_points_to_pixels, document_points_to_screen_points,
        fit_page_zoom, page_content_rect,
    },
};

use editor_input::{apply_viewport_input, handle_keyboard_input, handle_pointer_interaction};
use page_layout::layout_page_stack;
use palette::canvas_palette;

struct DocumentLayout {
    galley: Arc<egui::Galley>,
    list_markers: Vec<ListMarkerLayout>,
    images: Vec<ImageLayout>,
    manual_page_break_rows: Vec<usize>,
}

struct ActiveTightWrapFlow {
    pending_top_height: f32,
    remaining_height: f32,
    text_start_x: f32,
    text_width: f32,
}

struct TightWrapZone {
    row_index: usize,
    top: f32,
    bottom: f32,
    text_start_x: f32,
    text_width: f32,
}

struct ListMarkerLayout {
    row_index: usize,
    text: String,
    x: f32,
    font_id: FontId,
    color: Color32,
}

struct ImageLayout {
    row_index: usize,
    size: egui::Vec2,
    offset: egui::Vec2,
    image: DocumentImage,
}

pub fn paint_document_canvas(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    theme_mode: ThemeMode,
    history: &mut ChangeHistory,
) {
    let palette = canvas_palette(theme_mode);
    let viewport = ui.available_rect_before_wrap();
    let editor_id = Id::new("document_canvas");
    let response = ui.interact(viewport, editor_id, Sense::click_and_drag());
    let painter = ui.painter_at(viewport);
    let pixels_per_point = ui.ctx().pixels_per_point();
    apply_viewport_input(ui, &response, canvas);
    if canvas.zoom_mode == crate::app::ZoomMode::FitPage {
        canvas.zoom = fit_page_zoom(viewport, document.page_size);
    }

    painter.rect_filled(viewport, CornerRadius::ZERO, palette.canvas_bg);

    let base_page_rect =
        centered_page_rect(viewport, document.page_size, canvas.zoom, egui::Vec2::ZERO);
    let content_size = page_content_rect(base_page_rect, document.margins, canvas.zoom).size();
    let mut document_layout = layout_document(ui, document, canvas, content_size.x);
    let page_layout = layout_page_stack(
        viewport,
        document,
        canvas,
        &document_layout.galley,
        &document_layout.manual_page_break_rows,
    );

    let has_focus = ui.memory(|mem| mem.has_focus(editor_id));
    if has_focus && handle_keyboard_input(ui, document, canvas, &document_layout.galley, history) {
        document_layout = layout_document(ui, document, canvas, content_size.x);
    }

    if has_focus && !canvas.selection.is_empty() {
        paint_text_selection(
            &mut document_layout.galley,
            ui.visuals(),
            &canvas.selection,
            None,
        );
    }

    let mut new_image_rects: Vec<(usize, Rect)> = Vec::new();

    for page in &page_layout.pages {
        let shadow_offset = egui::vec2(
            document_points_to_screen_points(6.0, canvas.zoom),
            document_points_to_screen_points(8.0, canvas.zoom),
        );
        painter.rect_filled(
            page.page_rect.translate(shadow_offset),
            CornerRadius::same(2),
            palette.page_shadow,
        );
        painter.rect_filled(page.page_rect, CornerRadius::same(2), palette.page_bg);
        painter.rect_stroke(
            page.page_rect,
            CornerRadius::same(2),
            Stroke::new(
                1.0,
                if has_focus {
                    palette.page_focus
                } else {
                    palette.page_border
                },
            ),
            StrokeKind::Outside,
        );

        let visible_content_rect = Rect::from_min_size(
            page.content_rect.min,
            egui::vec2(
                page.content_rect.width(),
                (page.end_y - page.start_y)
                    .max(0.0)
                    .min(page.content_rect.height()),
            ),
        );
        let galley_origin = page.content_rect.min - egui::vec2(0.0, page.start_y);
        let zoom = canvas.zoom;

        // Helper: compute the screen rect for an image layout entry on this page.
        let image_screen_rect = |image: &ImageLayout| -> Option<Rect> {
            let row = document_layout.galley.rows.get(image.row_index)?;
            let image_y = row.pos.y;
            if image_y < page.start_y || image_y > page.end_y {
                return None;
            }
            Some(Rect::from_min_size(
                egui::pos2(
                    page.content_rect.left()
                        + row.pos.x
                        + image.offset.x
                        + document_points_to_screen_points(image.image.offset_x_points(), zoom),
                    page.content_rect.top() + image_y - page.start_y
                        + image.offset.y
                        + document_points_to_screen_points(image.image.offset_y_points(), zoom),
                ),
                image.size,
            ))
        };

        let page_clipped_painter = painter.with_clip_rect(page.page_rect);

        // Layer 1: Paint behind-text images (below everything)
        for image in &document_layout.images {
            if image.image.wrap_mode != WrapMode::BehindText {
                continue;
            }
            let Some(image_rect) = image_screen_rect(image) else {
                continue;
            };
            paint_image_on_page(
                ui,
                canvas,
                &page_clipped_painter,
                image,
                image_rect,
                &palette,
                1.0,
            );
            new_image_rects.push((image.image.id, image_rect));
        }

        // Layer 2: Paint text galley
        painter.with_clip_rect(visible_content_rect).galley(
            galley_origin,
            document_layout.galley.clone(),
            Color32::BLACK,
        );

        // List markers
        let clipped_painter = painter.with_clip_rect(visible_content_rect);
        for marker in &document_layout.list_markers {
            let Some(row) = document_layout.galley.rows.get(marker.row_index) else {
                continue;
            };
            let marker_y = row.pos.y;
            if marker_y < page.start_y || marker_y > page.end_y {
                continue;
            }

            let marker_pos = egui::pos2(
                page.content_rect.left() + marker.x,
                page.content_rect.top() + marker_y - page.start_y,
            );
            clipped_painter.text(
                marker_pos,
                Align2::RIGHT_TOP,
                &marker.text,
                marker.font_id.clone(),
                marker.color,
            );
        }

        // Layer 3: Paint normal images (not behind-text, not in-front-of-text) sorted by z-index
        let mut normal_images: Vec<&ImageLayout> = document_layout
            .images
            .iter()
            .filter(|img| !img.image.wrap_mode.is_no_text_displacement())
            .collect();
        normal_images.sort_by_key(|img| img.image.z_index);

        for image in &normal_images {
            let Some(image_rect) = image_screen_rect(image) else {
                continue;
            };
            paint_image_on_page(
                ui,
                canvas,
                &page_clipped_painter,
                image,
                image_rect,
                &palette,
                1.0,
            );
            new_image_rects.push((image.image.id, image_rect));
        }

        // Layer 4: Paint in-front-of-text images (above everything)
        let mut front_images: Vec<&ImageLayout> = document_layout
            .images
            .iter()
            .filter(|img| img.image.wrap_mode == WrapMode::InFrontOfText)
            .collect();
        front_images.sort_by_key(|img| img.image.z_index);

        for image in &front_images {
            let Some(image_rect) = image_screen_rect(image) else {
                continue;
            };
            paint_image_on_page(
                ui,
                canvas,
                &page_clipped_painter,
                image,
                image_rect,
                &palette,
                1.0,
            );
            new_image_rects.push((image.image.id, image_rect));
        }
    }

    canvas.image_rects = new_image_rects;
    let image_pointer_captured = handle_image_interaction(ui, &response, canvas, document);

    if !image_pointer_captured {
        handle_pointer_interaction(
            ui,
            &response,
            &page_layout,
            &document_layout.galley,
            canvas,
            document,
        );
    }

    // Draw ghost image if dragging
    if let Some(move_drag) = &canvas.move_drag {
        if move_drag.current_ptr != move_drag.start_ptr {
            let offset = move_drag.current_ptr - move_drag.start_ptr;
            let ghost_rect = move_drag.start_rect.translate(offset);
            if let Some(image_layout) = document_layout
                .images
                .iter()
                .find(|i| i.image.id == move_drag.image_id)
            {
                paint_image_on_page(
                    ui,
                    canvas,
                    &painter,
                    image_layout,
                    ghost_rect,
                    &palette,
                    0.5,
                );
            }
        }
    }

    // Draw selection border + handles with unclipped painter so they aren't cut at page margins
    if let Some((_, selected_rect)) = canvas
        .image_rects
        .iter()
        .find(|(id, _)| Some(*id) == canvas.selected_image_id)
    {
        paint_image_selection(&painter, *selected_rect);
    }

    if has_focus && canvas.selected_image_id.is_none() {
        if let Some(caret_rect) =
            page_layout.caret_rect(&document_layout.galley, canvas.selection.primary)
        {
            paint_text_cursor(
                ui,
                &painter,
                caret_rect,
                ui.input(|i| i.time) - canvas.last_interaction_time,
            );
        }
    }

    let page_pixels = (
        document_points_to_pixels(
            document.page_size.width_points,
            pixels_per_point,
            canvas.zoom,
        ),
        document_points_to_pixels(
            document.page_size.height_points,
            pixels_per_point,
            canvas.zoom,
        ),
    );
    let footer = format!(
        "{:.0} x {:.0} px  |  {} pages  |  y {:.0}",
        page_pixels.0,
        page_pixels.1,
        page_layout.pages.len(),
        canvas.pan.y
    );
    let footer_galley = painter.layout_no_wrap(
        footer,
        FontId::new(11.0, FontFamily::Monospace),
        palette.footer_text,
    );
    let footer_rect = Rect::from_min_size(
        egui::pos2(
            viewport.left() + 22.0,
            viewport.bottom() - footer_galley.size().y - 24.0,
        ),
        footer_galley.size() + egui::vec2(20.0, 14.0),
    );
    painter.rect_filled(footer_rect, CornerRadius::same(3), palette.footer_bg);
    painter.rect_stroke(
        footer_rect,
        CornerRadius::same(3),
        Stroke::new(1.0, palette.footer_stroke),
        StrokeKind::Outside,
    );
    painter.galley(
        egui::pos2(footer_rect.left() + 10.0, footer_rect.top() + 7.0),
        footer_galley,
        palette.footer_text,
    );
}

fn paint_image_on_page(
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

fn resize_handle_rects(image_rect: Rect) -> [(ResizeHandle, Rect); 8] {
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

fn paint_image_selection(painter: &egui::Painter, image_rect: Rect) {
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

fn handle_image_interaction(
    ui: &mut egui::Ui,
    response: &egui::Response,
    canvas: &mut CanvasState,
    document: &mut DocumentState,
) -> bool {
    const HANDLE_HIT_PADDING: f32 = 6.0;

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
            document.set_image_offset_by_id(
                move_drag.image_id,
                move_drag.start_x_points + dx,
                move_drag.start_y_points + dy,
            );
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
            )
        });
        if let Some((image_id, handle, start_ptr, start_w, start_h)) = drag_data {
            if let Some(pointer_pos) = response.interact_pointer_pos() {
                let zoom = canvas.zoom;
                let dx = (pointer_pos.x - start_ptr.x) / zoom;
                let dy = (pointer_pos.y - start_ptr.y) / zoom;
                let shift = ui.input(|i| i.modifiers.shift);
                let lock_aspect = document
                    .paragraph_images
                    .iter()
                    .flatten()
                    .find(|img| img.id == image_id)
                    .map(|img| img.lock_aspect_ratio)
                    .unwrap_or(false);
                let lock_ratio = lock_aspect ^ shift;
                let aspect = start_h / start_w.max(1.0);
                let (new_w, new_h) = match handle {
                    ResizeHandle::SE => {
                        if lock_ratio {
                            let w = start_w + dx;
                            (w, w * aspect)
                        } else {
                            (start_w + dx, start_h + dy)
                        }
                    }
                    ResizeHandle::SW => {
                        if lock_ratio {
                            let w = start_w - dx;
                            (w, w * aspect)
                        } else {
                            (start_w - dx, start_h + dy)
                        }
                    }
                    ResizeHandle::NE => {
                        if lock_ratio {
                            let w = start_w + dx;
                            (w, w * aspect)
                        } else {
                            (start_w + dx, start_h - dy)
                        }
                    }
                    ResizeHandle::NW => {
                        if lock_ratio {
                            let w = start_w - dx;
                            (w, w * aspect)
                        } else {
                            (start_w - dx, start_h - dy)
                        }
                    }
                    ResizeHandle::E => (start_w + dx, start_h),
                    ResizeHandle::W => (start_w - dx, start_h),
                    ResizeHandle::S => (start_w, start_h + dy),
                    ResizeHandle::N => (start_w, start_h - dy),
                };
                document.resize_image_by_id(image_id, new_w.max(24.0), new_h.max(24.0));
            }
            return true;
        }

        if let Some(move_drag) = canvas.move_drag.as_mut() {
            if let Some(pointer_pos) = response.interact_pointer_pos() {
                move_drag.current_ptr = pointer_pos;
                ui.ctx().set_cursor_icon(egui::CursorIcon::Grabbing);
            }
            return true;
        }
    }

    let Some(pointer_pos) = response.interact_pointer_pos() else {
        return false;
    };

    // Start a new resize drag when dragging from a handle
    if response.drag_started() {
        if let Some((image_id, handle)) = image_handle_hit(canvas, pointer_pos, HANDLE_HIT_PADDING)
        {
            let size = document
                .paragraph_images
                .iter()
                .flatten()
                .find(|img| img.id == image_id)
                .map(|img| (img.width_points, img.height_points));
            if let Some((w, h)) = size {
                canvas.selected_image_id = Some(image_id);
                canvas.move_drag = None;
                canvas.resize_drag = Some(ImageResizeDrag {
                    image_id,
                    handle,
                    start_ptr: pointer_pos,
                    start_width_points: w,
                    start_height_points: h,
                });
            }
            return true;
        }

        if let Some((image_id, image_rect)) = selected_image_rect(canvas) {
            if image_rect.contains(pointer_pos) {
                let offset = document
                    .paragraph_images
                    .iter()
                    .flatten()
                    .find(|img| img.id == image_id)
                    .map(|img| (img.offset_x_points(), img.offset_y_points()))
                    .unwrap_or((0.0, 0.0));
                canvas.resize_drag = None;
                canvas.move_drag = Some(ImageMoveDrag {
                    image_id,
                    start_ptr: pointer_pos,
                    current_ptr: pointer_pos,
                    start_rect: image_rect,
                    start_x_points: offset.0,
                    start_y_points: offset.1,
                });
                return true;
            }
        }

        if let Some(image_id) = canvas
            .image_rects
            .iter()
            .find(|(_, rect)| rect.contains(pointer_pos))
            .map(|(id, _)| *id)
        {
            canvas.selected_image_id = Some(image_id);
            canvas.resize_drag = None;
            canvas.move_drag = None;
            return true;
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
        if hit.is_none() {
            canvas.resize_drag = None;
            canvas.move_drag = None;
            return false;
        }
        return true;
    }

    false
}

fn selected_image_rect(canvas: &CanvasState) -> Option<(usize, Rect)> {
    let selected_id = canvas.selected_image_id?;
    canvas
        .image_rects
        .iter()
        .find(|(id, _)| *id == selected_id)
        .copied()
}

fn image_handle_hit(
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

    for &(image_id, image_rect) in canvas.image_rects.iter().rev() {
        if Some(image_id) == canvas.selected_image_id {
            continue;
        }
        for &(handle, handle_rect) in &resize_handle_rects(image_rect) {
            if handle_rect.expand(padding).contains(pointer_pos) {
                return Some((image_id, handle));
            }
        }
    }

    None
}

fn layout_document(
    ui: &mut egui::Ui,
    document: &DocumentState,
    canvas: &CanvasState,
    wrap_width: f32,
) -> DocumentLayout {
    let marker_gutter = document_points_to_screen_points(24.0, canvas.zoom);
    let marker_gap = document_points_to_screen_points(6.0, canvas.zoom);
    let default_style = CharacterStyle::default();
    let painter = ui.painter();

    let mut paragraph_galleys = Vec::new();
    let mut list_markers = Vec::new();
    let mut images = Vec::new();
    let mut manual_page_break_rows = Vec::new();
    let mut paragraph_spacing_ranges = Vec::new();
    let mut row_index = 0usize;
    let mut tight_wrap_flow: Option<ActiveTightWrapFlow> = None;

    for paragraph in document.paragraphs() {
        if tight_wrap_flow
            .as_ref()
            .is_some_and(|flow| flow.remaining_height <= 0.0)
        {
            tight_wrap_flow = None;
        }

        let base_indent = if paragraph.list_marker.is_some() {
            marker_gutter
        } else {
            0.0
        };
        let (indent, paragraph_wrap_width) = tight_wrap_flow
            .as_ref()
            .filter(|flow| flow.pending_top_height <= 0.0)
            .map_or_else(
                || (base_indent, (wrap_width - base_indent).max(1.0)),
                |flow| {
                    let start_x = flow.text_start_x.max(base_indent).clamp(0.0, wrap_width);
                    let end_x = (flow.text_start_x + flow.text_width).clamp(start_x, wrap_width);
                    (start_x, (end_x - start_x).max(1.0))
                },
            );
        let mut job = egui::epaint::text::LayoutJob::default();
        job.wrap.max_width = paragraph_wrap_width;
        job.break_on_newline = true;
        job.halign = egui::Align::LEFT;
        job.justify = paragraph.style.alignment == ParagraphAlignment::Justify;

        let has_visible_text = paragraph
            .runs
            .iter()
            .any(|run| run.text.chars().any(|ch| ch != OBJECT_REPLACEMENT_CHAR));

        if paragraph.runs.is_empty() {
            job.append("", 0.0, text_format(default_style, canvas.zoom));
        } else {
            for run in &paragraph.runs {
                append_run_with_placeholders(
                    &mut job,
                    run,
                    canvas.zoom,
                    paragraph.image.is_some() && !has_visible_text,
                );
            }
        }

        let marker_style = paragraph
            .runs
            .first()
            .map(|run| run.style)
            .unwrap_or(default_style);
        let mut paragraph_galley = painter.layout_job(job);

        if paragraph.style.page_break_before && row_index > 0 {
            manual_page_break_rows.push(row_index);
        }

        let mut tight_wrap_image_spec: Option<(f32, f32, f32, f32)> = None;
        if let Some(image) = paragraph.image.clone().filter(|_| !has_visible_text) {
            let wrap_mode = image.wrap_mode;
            let image_offset_x_points = image.offset_x_points();
            let image_offset_y_points = image.offset_y_points();
            let display_size = image_display_size(&image, paragraph_wrap_width, canvas.zoom);
            let reservation =
                block_image_reservation(wrap_mode, display_size, paragraph_wrap_width, canvas.zoom);
            if reserve_block_image_space(&mut paragraph_galley, reservation.row_size) {
                images.push(ImageLayout {
                    row_index,
                    size: display_size,
                    offset: reservation.image_offset,
                    image,
                });
            }

            if wrap_mode == WrapMode::Tight {
                tight_wrap_image_spec = Some((
                    display_size.x,
                    display_size.y,
                    reservation.image_offset.x
                        + document_points_to_screen_points(image_offset_x_points, canvas.zoom),
                    reservation.image_offset.y
                        + document_points_to_screen_points(image_offset_y_points, canvas.zoom),
                ));
            }
        }
        align_paragraph_galley(
            &mut paragraph_galley,
            indent,
            paragraph_wrap_width,
            paragraph.style.alignment,
        );
        apply_line_spacing(
            &mut paragraph_galley,
            paragraph.style.line_spacing,
            canvas.zoom,
        );

        if let Some((image_width, image_height, image_offset_x, image_offset_y)) =
            tight_wrap_image_spec
        {
            let tight_pad = tight_wrap_pad(canvas.zoom);
            let image_row_x = paragraph_galley
                .rows
                .first()
                .map(|row| row.pos.x)
                .unwrap_or(indent);
            let image_left = image_row_x + image_offset_x;
            let image_right = image_left + image_width;
            let left_width = (image_left - tight_pad).max(0.0);
            let right_start = (image_right + tight_pad).clamp(0.0, wrap_width);
            let right_width = (wrap_width - right_start).max(0.0);
            let min_side_width = document_points_to_screen_points(72.0, canvas.zoom);

            let (text_start_x, text_width) = if right_width >= left_width {
                (right_start, right_width)
            } else {
                (0.0, left_width)
            };

            let zone_top = image_offset_y - tight_pad;
            let zone_bottom = image_offset_y + image_height + tight_pad;
            let pending_top_height = zone_top.max(0.0);
            let remaining_height = (zone_bottom - pending_top_height).max(0.0);

            if text_width >= min_side_width && remaining_height > 0.0 {
                tight_wrap_flow = Some(ActiveTightWrapFlow {
                    pending_top_height,
                    remaining_height,
                    text_start_x,
                    text_width,
                });
            } else {
                tight_wrap_flow = None;
            }
        }

        if let Some(marker_text) = paragraph.list_marker {
            list_markers.push(ListMarkerLayout {
                row_index,
                text: marker_text,
                x: indent - marker_gap,
                font_id: text_format(marker_style, canvas.zoom).font_id,
                color: marker_style.text_color,
            });
        }

        if let Some(flow) = tight_wrap_flow.as_mut() {
            let paragraph_height = paragraph_galley.rect.height().max(0.0);
            if flow.pending_top_height > 0.0 {
                let consumed_top = paragraph_height.min(flow.pending_top_height);
                flow.pending_top_height -= consumed_top;
                flow.remaining_height -= (paragraph_height - consumed_top).max(0.0);
            } else {
                flow.remaining_height -= paragraph_height;
            }
        }

        let paragraph_row_count = paragraph_galley.rows.len();
        let paragraph_spacing_top = document_points_to_screen_points(
            f32::from(paragraph.style.spacing_before_points),
            canvas.zoom,
        );
        let paragraph_spacing_bottom = document_points_to_screen_points(
            f32::from(paragraph.style.spacing_after_points),
            canvas.zoom,
        );
        if paragraph_row_count > 0 {
            paragraph_spacing_ranges.push(ParagraphSpacingRange {
                row_start: row_index,
                row_end: row_index + paragraph_row_count,
                top: paragraph_spacing_top,
                bottom: paragraph_spacing_bottom,
            });
        }

        row_index += paragraph_row_count;
        paragraph_galleys.push(paragraph_galley);
    }

    let plain_text = document.plain_text();
    let mut merged_job = egui::epaint::text::LayoutJob::default();
    merged_job.wrap.max_width = wrap_width;
    merged_job.break_on_newline = true;
    merged_job.append(&plain_text, 0.0, text_format(default_style, canvas.zoom));

    let mut galley = Arc::new(egui::Galley::concat(
        Arc::new(merged_job),
        &paragraph_galleys,
        painter.pixels_per_point(),
    ));
    apply_paragraph_row_spacing(&mut galley, &paragraph_spacing_ranges);
    apply_tight_wrap_row_offsets(&mut galley, &images, canvas.zoom);

    DocumentLayout {
        galley,
        list_markers,
        images,
        manual_page_break_rows,
    }
}

fn append_run_with_placeholders(
    job: &mut egui::epaint::text::LayoutJob,
    run: &crate::document::TextRun,
    zoom: f32,
    keep_visible_placeholder: bool,
) {
    let mut segment = String::new();
    for ch in run.text.chars() {
        if ch == OBJECT_REPLACEMENT_CHAR {
            if !segment.is_empty() {
                job.append(&segment, 0.0, text_format(run.style, zoom));
                segment.clear();
            }

            let mut placeholder_style = run.style;
            if !keep_visible_placeholder {
                placeholder_style.text_color = Color32::TRANSPARENT;
            }
            job.append(
                &OBJECT_REPLACEMENT_CHAR.to_string(),
                0.0,
                text_format(placeholder_style, zoom),
            );
        } else {
            segment.push(ch);
        }
    }

    if !segment.is_empty() {
        job.append(&segment, 0.0, text_format(run.style, zoom));
    }
}

fn apply_line_spacing(
    galley: &mut Arc<egui::Galley>,
    line_spacing: crate::document::LineSpacing,
    zoom: f32,
) {
    if galley.rows.len() < 2 {
        return;
    }

    let galley = Arc::make_mut(galley);
    let original_rect = galley.rect;
    let mut mesh_bounds = egui::Rect::NOTHING;
    let mut cumulative_shift = 0.0;

    for row_index in 1..galley.rows.len() {
        let previous_row_height = galley.rows[row_index - 1].row.height();
        let desired_advance = match line_spacing.kind {
            LineSpacingKind::AutoMultiplier => previous_row_height * line_spacing.value.max(0.0),
            LineSpacingKind::AtLeastPoints => previous_row_height.max(
                document_points_to_screen_points(line_spacing.value.max(0.0), zoom),
            ),
            LineSpacingKind::ExactPoints => {
                document_points_to_screen_points(line_spacing.value.max(0.0), zoom)
            }
        };
        cumulative_shift += desired_advance - previous_row_height;
        galley.rows[row_index].pos.y += cumulative_shift;
    }

    for row in &galley.rows {
        mesh_bounds |= row.visuals.mesh_bounds.translate(row.pos.to_vec2());
    }

    galley.rect = egui::Rect::from_min_max(
        original_rect.min,
        egui::pos2(original_rect.max.x, original_rect.max.y + cumulative_shift),
    );
    galley.mesh_bounds = mesh_bounds;
}

struct BlockImageReservation {
    row_size: egui::Vec2,
    image_offset: egui::Vec2,
}

fn block_image_reservation(
    wrap_mode: WrapMode,
    image_size: egui::Vec2,
    wrap_width: f32,
    zoom: f32,
) -> BlockImageReservation {
    let square_pad = document_points_to_screen_points(12.0, zoom);
    let tight_pad = document_points_to_screen_points(4.0, zoom);

    match wrap_mode {
        WrapMode::Inline => BlockImageReservation {
            row_size: image_size,
            image_offset: egui::Vec2::ZERO,
        },
        WrapMode::Square => {
            let row_width = (image_size.x + square_pad * 2.0).min(wrap_width);
            let row_height = image_size.y + square_pad * 2.0;
            BlockImageReservation {
                row_size: egui::vec2(row_width, row_height),
                image_offset: egui::vec2(((row_width - image_size.x) * 0.5).max(0.0), square_pad),
            }
        }
        WrapMode::Tight => {
            let row_width = (image_size.x + tight_pad * 2.0).min(wrap_width);
            let row_height = tight_wrap_row_height(zoom);
            BlockImageReservation {
                row_size: egui::vec2(row_width, row_height),
                image_offset: egui::vec2(tight_pad, 0.0),
            }
        }
        WrapMode::Through | WrapMode::BehindText | WrapMode::InFrontOfText => {
            BlockImageReservation {
                row_size: egui::Vec2::ZERO,
                image_offset: egui::Vec2::ZERO,
            }
        }
        WrapMode::TopAndBottom => BlockImageReservation {
            row_size: egui::vec2(wrap_width, image_size.y),
            image_offset: egui::vec2(((wrap_width - image_size.x) * 0.5).max(0.0), 0.0),
        },
    }
}

fn tight_wrap_pad(zoom: f32) -> f32 {
    document_points_to_screen_points(4.0, zoom)
}

fn tight_wrap_row_height(zoom: f32) -> f32 {
    document_points_to_screen_points(14.0, zoom).max(1.0)
}

fn apply_tight_wrap_row_offsets(galley: &mut Arc<egui::Galley>, images: &[ImageLayout], zoom: f32) {
    let zones = tight_wrap_zones(galley, images, zoom);
    if zones.is_empty() {
        return;
    }

    let galley = Arc::make_mut(galley);
    let mut min_rect = galley.rect.min;
    let mut max_rect = galley.rect.max;
    let mut mesh_bounds = egui::Rect::NOTHING;

    for (index, row) in galley.rows.iter_mut().enumerate() {
        let row_text = row.row.text();
        if !row_text.chars().all(|ch| ch == OBJECT_REPLACEMENT_CHAR) {
            if let Some(zone) = zones.iter().find(|zone| {
                zone.row_index != index && row.max_y() > zone.top && row.min_y() < zone.bottom
            }) {
                if row.row.size.x <= zone.text_width {
                    row.pos.x = zone.text_start_x;
                }
            }
        }

        let row_rect = row.rect();
        min_rect.x = min_rect.x.min(row_rect.min.x);
        min_rect.y = min_rect.y.min(row_rect.min.y);
        max_rect.x = max_rect.x.max(row_rect.max.x);
        max_rect.y = max_rect.y.max(row_rect.max.y);
        mesh_bounds |= row.visuals.mesh_bounds.translate(row.pos.to_vec2());
    }

    galley.rect = Rect::from_min_max(min_rect, max_rect);
    galley.mesh_bounds = mesh_bounds;
}

fn tight_wrap_zones(
    galley: &egui::Galley,
    images: &[ImageLayout],
    zoom: f32,
) -> Vec<TightWrapZone> {
    let tight_pad = tight_wrap_pad(zoom);
    let min_side_width = document_points_to_screen_points(72.0, zoom);
    let mut zones = Vec::new();

    for image in images
        .iter()
        .filter(|image| image.image.wrap_mode == WrapMode::Tight)
    {
        let Some(row) = galley.rows.get(image.row_index) else {
            continue;
        };

        let image_left = row.pos.x
            + image.offset.x
            + document_points_to_screen_points(image.image.offset_x_points(), zoom);
        let image_top = row.pos.y
            + image.offset.y
            + document_points_to_screen_points(image.image.offset_y_points(), zoom);
        let image_right = image_left + image.size.x;
        let image_bottom = image_top + image.size.y;
        let left_width = (image_left - tight_pad).max(0.0);
        let right_start = (image_right + tight_pad).max(0.0);
        let right_width = (galley.rect.width() - right_start).max(0.0);
        let (text_start_x, text_width) = if right_width >= left_width {
            (right_start, right_width)
        } else {
            (0.0, left_width)
        };

        if text_width >= min_side_width {
            zones.push(TightWrapZone {
                row_index: image.row_index,
                top: image_top - tight_pad,
                bottom: image_bottom + tight_pad,
                text_start_x,
                text_width,
            });
        }
    }

    zones
}

fn reserve_block_image_space(galley: &mut Arc<egui::Galley>, row_size: egui::Vec2) -> bool {
    let galley = Arc::make_mut(galley);
    let Some(placed_row) = galley.rows.first_mut() else {
        return false;
    };
    let row = Arc::make_mut(&mut placed_row.row);
    if row.glyphs.len() != 1 || row.glyphs[0].chr != OBJECT_REPLACEMENT_CHAR {
        return false;
    }

    row.glyphs[0].advance_width = row_size.x;
    row.glyphs[0].line_height = row_size.y;
    row.glyphs[0].font_ascent = row_size.y;
    row.glyphs[0].font_height = row_size.y;
    row.glyphs[0].font_face_ascent = row_size.y;
    row.glyphs[0].font_face_height = row_size.y;
    row.size = row_size;
    row.visuals = Default::default();

    galley.rect = Rect::from_min_size(galley.rect.min, row_size);
    galley.mesh_bounds = Rect::NOTHING;
    true
}

fn image_display_size(image: &DocumentImage, wrap_width: f32, zoom: f32) -> egui::Vec2 {
    let width = document_points_to_screen_points(image.width_points.max(24.0), zoom);
    let height = document_points_to_screen_points(image.height_points.max(24.0), zoom);
    if width <= wrap_width {
        return egui::vec2(width, height);
    }

    let scale = wrap_width / width;
    egui::vec2(wrap_width, (height * scale).max(24.0))
}

fn texture_for_image<'a>(
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

fn align_paragraph_galley(
    galley: &mut Arc<egui::Galley>,
    indent: f32,
    wrap_width: f32,
    alignment: ParagraphAlignment,
) {
    let target_offsets: Vec<f32> = galley
        .rows
        .iter()
        .enumerate()
        .map(|(index, row)| {
            let base_offset = match alignment {
                ParagraphAlignment::Left | ParagraphAlignment::Justify => 0.0,
                ParagraphAlignment::Center => ((wrap_width - row.size.x) * 0.5).max(0.0),
                ParagraphAlignment::Right => (wrap_width - row.size.x).max(0.0),
            };
            let current_x = galley.rows[index].pos.x;
            indent + base_offset - current_x
        })
        .collect();

    if target_offsets
        .iter()
        .all(|delta| delta.abs() <= f32::EPSILON)
    {
        return;
    }

    let galley = Arc::make_mut(galley);
    let mut min_rect = galley.rect.min;
    let mut max_rect = galley.rect.max;
    let mut mesh_bounds = egui::Rect::NOTHING;

    for (row, delta) in galley.rows.iter_mut().zip(target_offsets) {
        row.pos.x += delta;
        let row_rect = row.rect();
        min_rect.x = min_rect.x.min(row_rect.min.x);
        min_rect.y = min_rect.y.min(row_rect.min.y);
        max_rect.x = max_rect.x.max(row_rect.max.x);
        max_rect.y = max_rect.y.max(row_rect.max.y);
        mesh_bounds |= row.visuals.mesh_bounds.translate(row.pos.to_vec2());
    }

    galley.rect = egui::Rect::from_min_max(min_rect, max_rect);
    galley.mesh_bounds = mesh_bounds;
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        app::{CanvasState, ZoomMode},
        document::{
            CharacterStyle, DocumentImage, DocumentState, ImageLayoutMode, ImageRendering,
            LineSpacing, LineSpacingKind, ListKind, PageMargins, PageSize, ParagraphAlignment,
            ParagraphStyle, TextRun, WrapMode, OBJECT_REPLACEMENT_CHAR,
        },
        layout::fit_page_zoom,
    };

    /// Run `layout_document` inside a headless egui context and return
    /// the layout result for assertion.
    fn run_headless_layout(
        document: &DocumentState,
        canvas: &CanvasState,
        wrap_width: f32,
    ) -> DocumentLayout {
        let ctx = egui::Context::default();
        let mut layout = None;

        let _ = ctx.run_ui(egui::RawInput::default(), |ui| {
            layout = Some(layout_document(ui, document, canvas, wrap_width));
        });

        layout.expect("layout_document should have been called inside the egui frame")
    }

    fn make_document(
        runs: Vec<TextRun>,
        paragraph_styles: Vec<ParagraphStyle>,
        paragraph_images: Vec<Option<DocumentImage>>,
    ) -> DocumentState {
        DocumentState {
            title: "Test".to_owned(),
            runs,
            paragraph_styles,
            paragraph_images,
            page_size: PageSize::a4(),
            margins: PageMargins::standard(),
        }
    }

    fn make_test_image(id: usize, width: f32, height: f32, wrap_mode: WrapMode) -> DocumentImage {
        DocumentImage {
            id,
            bytes: vec![],
            alt_text: "test".to_owned(),
            width_points: width,
            height_points: height,
            lock_aspect_ratio: true,
            opacity: 1.0,
            layout_mode: ImageLayoutMode::Inline,
            wrap_mode,
            rendering: ImageRendering::Smooth,
            horizontal_position: Default::default(),
            vertical_position: Default::default(),
            distance_from_text: Default::default(),
            z_index: 0,
            move_with_text: true,
            allow_overlap: false,
        }
    }

    #[test]
    fn single_paragraph_produces_at_least_one_row() {
        let document = make_document(
            vec![TextRun {
                text: "Hello world".to_owned(),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle::default()],
            vec![None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        assert!(
            !layout.galley.rows.is_empty(),
            "galley should have at least one row"
        );
        assert!(
            layout.manual_page_break_rows.is_empty(),
            "no manual page breaks expected"
        );
        assert!(layout.images.is_empty(), "no images expected");
    }

    #[test]
    fn long_text_wraps_into_multiple_rows() {
        let long_text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. ".repeat(20);
        let document = make_document(
            vec![TextRun {
                text: long_text,
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle::default()],
            vec![None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 300.0);

        assert!(
            layout.galley.rows.len() > 1,
            "long text at narrow width should produce multiple rows, got {}",
            layout.galley.rows.len()
        );
    }

    #[test]
    fn manual_page_break_is_recorded() {
        let document = make_document(
            vec![TextRun {
                text: "First paragraph\nSecond paragraph".to_owned(),
                style: CharacterStyle::default(),
            }],
            vec![
                ParagraphStyle::default(),
                ParagraphStyle {
                    page_break_before: true,
                    ..ParagraphStyle::default()
                },
            ],
            vec![None, None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        assert!(
            !layout.manual_page_break_rows.is_empty(),
            "should record a manual page break"
        );
        // The page break should be at the row index where the second paragraph starts.
        assert!(
            layout.manual_page_break_rows[0] > 0,
            "page break row index should be > 0"
        );
    }

    #[test]
    fn block_image_paragraph_produces_image_layout() {
        let document = make_document(
            vec![TextRun {
                text: format!("{OBJECT_REPLACEMENT_CHAR}"),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle::default()],
            vec![Some(make_test_image(1, 200.0, 100.0, WrapMode::Inline))],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        assert_eq!(layout.images.len(), 1, "should have one image layout");
        assert_eq!(layout.images[0].row_index, 0);
        assert!(
            layout.images[0].size.x > 0.0 && layout.images[0].size.y > 0.0,
            "image size should be positive"
        );
    }

    #[test]
    fn centered_alignment_offsets_row_positions() {
        let document = make_document(
            vec![TextRun {
                text: "Short".to_owned(),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle {
                alignment: ParagraphAlignment::Center,
                ..ParagraphStyle::default()
            }],
            vec![None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        assert!(!layout.galley.rows.is_empty());
        let row = &layout.galley.rows[0];
        // Centered text on a 600px wrap should have a positive x offset.
        assert!(
            row.pos.x > 0.0,
            "centered text should be offset from left edge, got x={}",
            row.pos.x
        );
    }

    #[test]
    fn multiple_paragraphs_with_varying_styles() {
        let document = make_document(
            vec![TextRun {
                text: "Left aligned\nCenter aligned\nRight aligned".to_owned(),
                style: CharacterStyle::default(),
            }],
            vec![
                ParagraphStyle {
                    alignment: ParagraphAlignment::Left,
                    ..ParagraphStyle::default()
                },
                ParagraphStyle {
                    alignment: ParagraphAlignment::Center,
                    ..ParagraphStyle::default()
                },
                ParagraphStyle {
                    alignment: ParagraphAlignment::Right,
                    ..ParagraphStyle::default()
                },
            ],
            vec![None, None, None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        // Should have at least 3 rows (one per paragraph).
        assert!(
            layout.galley.rows.len() >= 3,
            "expected at least 3 rows for 3 paragraphs, got {}",
            layout.galley.rows.len()
        );
    }

    #[test]
    fn paragraph_spacing_offsets_following_rows() {
        let document = make_document(
            vec![TextRun {
                text: "First\nSecond".to_owned(),
                style: CharacterStyle::default(),
            }],
            vec![
                ParagraphStyle {
                    spacing_after_points: 24,
                    ..ParagraphStyle::default()
                },
                ParagraphStyle::default(),
            ],
            vec![None, None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        assert!(layout.galley.rows.len() >= 2);
        let first_row = &layout.galley.rows[0];
        let second_row = &layout.galley.rows[1];
        assert!(
            second_row.pos.y - first_row.pos.y > first_row.rect().height(),
            "paragraph spacing should increase the gap between rows"
        );
    }

    #[test]
    fn auto_multiplier_line_spacing_offsets_following_line() {
        let document = make_document(
            vec![TextRun {
                text: "First line in paragraph that wraps on purpose because the width is narrow"
                    .to_owned(),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle {
                line_spacing: LineSpacing {
                    kind: LineSpacingKind::AutoMultiplier,
                    value: 1.5,
                },
                ..ParagraphStyle::default()
            }],
            vec![None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 240.0);

        assert!(layout.galley.rows.len() >= 2);
        let first_row = &layout.galley.rows[0];
        let second_row = &layout.galley.rows[1];
        let default_gap = first_row.row.height();
        let actual_gap = second_row.pos.y - first_row.pos.y;
        assert!(
            actual_gap > default_gap * 1.45,
            "1.5x line spacing should enlarge row advance, got actual_gap={actual_gap}, default_gap={default_gap}"
        );
    }

    #[test]
    fn exact_line_spacing_uses_requested_row_advance() {
        let document = make_document(
            vec![TextRun {
                text: "First line in paragraph that wraps on purpose because the width is narrow"
                    .to_owned(),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle {
                line_spacing: LineSpacing {
                    kind: LineSpacingKind::ExactPoints,
                    value: 24.0,
                },
                ..ParagraphStyle::default()
            }],
            vec![None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 240.0);

        assert!(layout.galley.rows.len() >= 2);
        let first_row = &layout.galley.rows[0];
        let second_row = &layout.galley.rows[1];
        let actual_gap = second_row.pos.y - first_row.pos.y;
        assert!(
            (actual_gap - 24.0).abs() < 1.5,
            "exact line spacing should follow the requested advance, got {actual_gap}"
        );
    }

    #[test]
    fn fit_page_zoom_uses_manual_override_rules() {
        let viewport = egui::Rect::from_min_size(egui::Pos2::ZERO, egui::vec2(400.0, 500.0));
        let fit = fit_page_zoom(viewport, PageSize::a4());
        assert!(
            fit < 1.0,
            "fit zoom should shrink an A4 page in a small viewport"
        );

        let mut canvas = CanvasState::default();
        canvas.imported_docx_view = true;
        canvas.zoom_mode = ZoomMode::FitPage;
        canvas.zoom = fit;
        canvas.zoom_mode = ZoomMode::Manual;
        canvas.zoom = (canvas.zoom * 1.1).clamp(0.5, 3.0);
        assert_eq!(canvas.zoom_mode, ZoomMode::Manual);
        assert!(canvas.zoom > fit);
    }

    #[test]
    fn image_with_page_break_in_multi_paragraph_document() {
        let document = make_document(
            vec![TextRun {
                text: format!("Intro paragraph\n{OBJECT_REPLACEMENT_CHAR}\nClosing paragraph"),
                style: CharacterStyle::default(),
            }],
            vec![
                ParagraphStyle::default(),
                ParagraphStyle {
                    page_break_before: true,
                    ..ParagraphStyle::default()
                },
                ParagraphStyle::default(),
            ],
            vec![
                None,
                Some(make_test_image(2, 400.0, 300.0, WrapMode::Inline)),
                None,
            ],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        assert_eq!(layout.images.len(), 1, "should have one image");
        assert!(
            !layout.manual_page_break_rows.is_empty(),
            "should have a manual page break"
        );
        // Image row should match the page break row.
        assert_eq!(
            layout.images[0].row_index, layout.manual_page_break_rows[0],
            "image should be on the page-break row"
        );
    }

    #[test]
    fn list_markers_are_produced_for_bullet_paragraphs() {
        let document = make_document(
            vec![TextRun {
                text: "Item one\nItem two".to_owned(),
                style: CharacterStyle::default(),
            }],
            vec![
                ParagraphStyle {
                    list_kind: ListKind::Bullet,
                    ..ParagraphStyle::default()
                },
                ParagraphStyle {
                    list_kind: ListKind::Bullet,
                    ..ParagraphStyle::default()
                },
            ],
            vec![None, None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        assert_eq!(layout.list_markers.len(), 2, "should have two list markers");
        assert_eq!(layout.list_markers[0].text, "•");
        assert_eq!(layout.list_markers[1].text, "•");
    }

    #[test]
    fn wrap_modes_change_image_row_geometry() {
        let make_doc = |wrap_mode| {
            make_document(
                vec![TextRun {
                    text: format!("{OBJECT_REPLACEMENT_CHAR}"),
                    style: CharacterStyle::default(),
                }],
                vec![ParagraphStyle::default()],
                vec![Some(make_test_image(10, 180.0, 90.0, wrap_mode))],
            )
        };

        let canvas = CanvasState::default();
        let wrap_width = 600.0;

        let inline = run_headless_layout(&make_doc(WrapMode::Inline), &canvas, wrap_width);
        let square = run_headless_layout(&make_doc(WrapMode::Square), &canvas, wrap_width);
        let tight = run_headless_layout(&make_doc(WrapMode::Tight), &canvas, wrap_width);
        let through = run_headless_layout(&make_doc(WrapMode::Through), &canvas, wrap_width);
        let top_bottom =
            run_headless_layout(&make_doc(WrapMode::TopAndBottom), &canvas, wrap_width);

        let inline_row = &inline.galley.rows[inline.images[0].row_index];
        let square_row = &square.galley.rows[square.images[0].row_index];
        let tight_row = &tight.galley.rows[tight.images[0].row_index];
        let through_row = &through.galley.rows[through.images[0].row_index];
        let top_bottom_row = &top_bottom.galley.rows[top_bottom.images[0].row_index];

        assert!(
            square_row.size.x > tight_row.size.x,
            "square wrap should reserve more horizontal space than tight"
        );
        assert!(
            square_row.size.y > tight_row.size.y,
            "square wrap should reserve more vertical space than tight"
        );
        assert!(
            through_row.size.x <= 1.0 && through_row.size.y <= 1.0,
            "through wrap should not reserve layout space"
        );
        assert!(
            (top_bottom_row.size.x - wrap_width).abs() < 1.0,
            "top-and-bottom wrap should reserve the full row width"
        );
        assert!(
            top_bottom.images[0].offset.x > through.images[0].offset.x,
            "top-and-bottom should center image while through should anchor without horizontal reservation"
        );
        assert!(
            inline_row.size.x < top_bottom_row.size.x,
            "inline wrap should keep a tighter row than top-and-bottom"
        );
    }

    #[test]
    fn tight_wrap_chooses_side_from_image_position() {
        let make_doc = |offset_x: f32| {
            let mut img = make_test_image(21, 180.0, 90.0, WrapMode::Tight);
            img.horizontal_position.offset_points = offset_x;
            make_document(
                vec![TextRun {
                    text: format!(
                        "{OBJECT_REPLACEMENT_CHAR}\nthis paragraph should flow beside the image and expose side choice"
                    ),
                    style: CharacterStyle::default(),
                }],
                vec![ParagraphStyle::default(), ParagraphStyle::default()],
                vec![
                    Some(img),
                    None,
                ],
            )
        };

        let canvas = CanvasState::default();
        let wrap_width = 600.0;
        let left_placed = run_headless_layout(&make_doc(0.0), &canvas, wrap_width);
        let right_placed = run_headless_layout(&make_doc(220.0), &canvas, wrap_width);

        let left_row_x = left_placed.galley.rows[1].pos.x;
        let right_row_x = right_placed.galley.rows[1].pos.x;

        assert!(
            left_row_x > right_row_x,
            "left-placed image should push text to the right; got left_row_x={left_row_x}, right_row_x={right_row_x}"
        );
    }

    #[test]
    fn tight_wrap_vertical_offset_delays_text_wrapping() {
        let make_doc = |offset_y: f32| {
            let mut img = make_test_image(22, 180.0, 90.0, WrapMode::Tight);
            img.vertical_position.offset_points = offset_y;
            make_document(
                vec![TextRun {
                    text: format!(
                        "{OBJECT_REPLACEMENT_CHAR}\nthis paragraph should start unwrapped when image is moved down enough"
                    ),
                    style: CharacterStyle::default(),
                }],
                vec![ParagraphStyle::default(), ParagraphStyle::default()],
                vec![
                    Some(img),
                    None,
                ],
            )
        };

        let canvas = CanvasState::default();
        let wrap_width = 600.0;
        let normal = run_headless_layout(&make_doc(0.0), &canvas, wrap_width);
        let moved_down = run_headless_layout(&make_doc(220.0), &canvas, wrap_width);

        let normal_row_x = normal.galley.rows[1].pos.x;
        let moved_down_row_x = moved_down.galley.rows[1].pos.x;

        assert!(
            normal_row_x > moved_down_row_x + 40.0,
            "moving image down should reduce early side-wrapping; got normal_row_x={normal_row_x}, moved_down_row_x={moved_down_row_x}"
        );
    }
}

struct ParagraphSpacingRange {
    row_start: usize,
    row_end: usize,
    top: f32,
    bottom: f32,
}

fn apply_paragraph_row_spacing(
    galley: &mut Arc<egui::Galley>,
    paragraph_spacing_ranges: &[ParagraphSpacingRange],
) {
    if paragraph_spacing_ranges.is_empty() {
        return;
    }

    let galley = Arc::make_mut(galley);
    let original_rect = galley.rect;
    let mut mesh_bounds = egui::Rect::NOTHING;
    let mut cumulative_shift = 0.0;
    let row_count = galley.rows.len();

    for range in paragraph_spacing_ranges {
        let paragraph_shift = cumulative_shift + range.top;
        let start = range.row_start.min(row_count);
        let end = range.row_end.min(row_count);
        for row in galley.rows[start..end].iter_mut() {
            row.pos.y += paragraph_shift;
        }
        cumulative_shift += range.top + range.bottom;
    }

    for row in &galley.rows {
        mesh_bounds |= row.visuals.mesh_bounds.translate(row.pos.to_vec2());
    }

    galley.rect = egui::Rect::from_min_max(
        original_rect.min,
        egui::pos2(original_rect.max.x, original_rect.max.y + cumulative_shift),
    );
    galley.mesh_bounds = mesh_bounds;
}
