mod editor_input;
mod page_layout;
mod palette;

use std::sync::Arc;

use eframe::egui::{
    self,
    epaint::CornerRadius,
    text_selection::visuals::{paint_text_cursor, paint_text_selection},
    Align2, Color32, FontFamily, FontId, Id, Rect, Sense, Stroke, StrokeKind,
};

use crate::{
    app::{CanvasState, ThemeMode},
    document::{text_format, CharacterStyle, DocumentState, ParagraphAlignment},
    layout::{
        centered_page_rect, document_points_to_pixels, document_points_to_screen_points,
        page_content_rect,
    },
};

use editor_input::{apply_viewport_input, handle_keyboard_input, handle_pointer_interaction};
use page_layout::layout_page_stack;
use palette::canvas_palette;

struct DocumentLayout {
    galley: Arc<egui::Galley>,
    list_markers: Vec<ListMarkerLayout>,
}

struct ListMarkerLayout {
    row_index: usize,
    text: String,
    x: f32,
    font_id: FontId,
    color: Color32,
}

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
    let mut document_layout = layout_document(ui, document, canvas, content_size.x);
    let page_layout = layout_page_stack(viewport, document, canvas, &document_layout.galley);

    handle_pointer_interaction(
        ui,
        &response,
        &page_layout,
        &document_layout.galley,
        canvas,
        document,
    );

    let has_focus = ui.memory(|mem| mem.has_focus(editor_id));
    if has_focus && handle_keyboard_input(ui, document, canvas, &document_layout.galley) {
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
            document_layout.galley.clone(),
            Color32::BLACK,
        );

        let clipped_painter = painter.with_clip_rect(page.content_rect);
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
    }

    if has_focus {
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
    let mut row_index = 0usize;

    for paragraph in document.paragraphs() {
        let indent = if paragraph.list_marker.is_some() {
            marker_gutter
        } else {
            0.0
        };
        let paragraph_wrap_width = (wrap_width - indent).max(1.0);
        let mut job = egui::epaint::text::LayoutJob::default();
        job.wrap.max_width = paragraph_wrap_width;
        job.break_on_newline = true;
        job.halign = egui::Align::LEFT;
        job.justify = paragraph.style.alignment == ParagraphAlignment::Justify;

        if paragraph.runs.is_empty() {
            job.append("", 0.0, text_format(default_style, canvas.zoom));
        } else {
            for run in &paragraph.runs {
                job.append(&run.text, 0.0, text_format(run.style, canvas.zoom));
            }
        }

        let marker_style = paragraph
            .runs
            .first()
            .map(|run| run.style)
            .unwrap_or(default_style);
        let mut paragraph_galley = painter.layout_job(job);
        align_paragraph_galley(
            &mut paragraph_galley,
            indent,
            paragraph_wrap_width,
            paragraph.style.alignment,
        );

        if let Some(marker_text) = paragraph.list_marker {
            list_markers.push(ListMarkerLayout {
                row_index,
                text: marker_text,
                x: indent - marker_gap,
                font_id: text_format(marker_style, canvas.zoom).font_id,
                color: marker_style.text_color,
            });
        }

        row_index += paragraph_galley.rows.len();
        paragraph_galleys.push(paragraph_galley);
    }

    let plain_text = document.plain_text();
    let mut merged_job = egui::epaint::text::LayoutJob::default();
    merged_job.wrap.max_width = wrap_width;
    merged_job.break_on_newline = true;
    merged_job.append(&plain_text, 0.0, text_format(default_style, canvas.zoom));

    DocumentLayout {
        galley: Arc::new(egui::Galley::concat(
            Arc::new(merged_job),
            &paragraph_galleys,
            painter.pixels_per_point(),
        )),
        list_markers,
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
