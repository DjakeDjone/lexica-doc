mod editor_input;
mod page_layout;
mod palette;

use std::sync::Arc;

use eframe::egui::{
    self,
    epaint::CornerRadius,
    text_selection::visuals::{paint_text_cursor, paint_text_selection},
    Color32, FontFamily, FontId, Id, Rect, Sense, Stroke, StrokeKind,
};

use crate::{
    app::{CanvasState, ThemeMode},
    document::DocumentState,
    layout::{
        centered_page_rect, document_points_to_pixels, document_points_to_screen_points,
        page_content_rect,
    },
};

use editor_input::{apply_viewport_input, handle_keyboard_input, handle_pointer_interaction};
use page_layout::layout_page_stack;
use palette::canvas_palette;

pub fn paint_document_canvas(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    theme_mode: ThemeMode,
) {
    let palette = canvas_palette(theme_mode);
    let viewport = ui.available_rect_before_wrap();
    let editor_id = Id::new("document_canvas");
    let response = ui.interact(viewport, editor_id, Sense::click_and_drag());
    let painter = ui.painter_at(viewport);
    let pixels_per_point = ui.ctx().pixels_per_point();
    apply_viewport_input(ui, &response, canvas);

    painter.rect_filled(viewport, CornerRadius::ZERO, palette.canvas_bg);

    let base_page_rect =
        centered_page_rect(viewport, document.page_size, canvas.zoom, egui::Vec2::ZERO);
    let content_size = page_content_rect(base_page_rect, document.margins, canvas.zoom).size();
    let mut galley = layout_document(ui, document, canvas, content_size.x);
    let page_layout = layout_page_stack(viewport, document, canvas, &galley);

    handle_pointer_interaction(ui, &response, &page_layout, &galley, canvas, document);

    let has_focus = ui.memory(|mem| mem.has_focus(editor_id));
    if has_focus && handle_keyboard_input(ui, document, canvas, &galley) {
        galley = layout_document(ui, document, canvas, content_size.x);
    }

    if has_focus && !canvas.selection.is_empty() {
        paint_text_selection(&mut galley, ui.visuals(), &canvas.selection, None);
    }

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

        let galley_origin = page.content_rect.min - egui::vec2(0.0, page.start_y);
        painter.with_clip_rect(page.content_rect).galley(
            galley_origin,
            galley.clone(),
            Color32::BLACK,
        );
    }

    if has_focus {
        if let Some(caret_rect) = page_layout.caret_rect(&galley, canvas.selection.primary) {
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

fn layout_document(
    ui: &mut egui::Ui,
    document: &DocumentState,
    canvas: &CanvasState,
    wrap_width: f32,
) -> Arc<egui::Galley> {
    ui.painter()
        .layout_job(document.layout_job(canvas.zoom, wrap_width))
}
