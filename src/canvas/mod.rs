mod editor_input;
mod image;
mod page_layout;
mod palette;
mod table;

pub mod header_footer;
pub mod image_wrap;
pub mod layout;

#[cfg(test)]
mod tests;

use eframe::egui::{
    self, epaint::text::cursor::CCursor, epaint::CornerRadius,
    text_selection::visuals::paint_text_cursor, text_selection::visuals::paint_text_selection,
    text_selection::CCursorRange, Align2, Color32, EventFilter, Id, Rect, Sense, Stroke,
    StrokeKind,
};

use crate::{
    app::{
        ActiveHeaderFooter, CanvasState, ChangeHistory, ResizeHandle, TableResizeHandleRect,
        TableResizeKind, ThemeMode,
    },
    document::{CharacterStyle, DocumentState, TextRun, WrapMode},
    grammar::GrammarError,
    layout::{
        centered_page_rect, document_points_to_screen_points, fit_page_zoom,
        section_page_content_rect,
    },
    ui::squiggles::{paint_grammar_squiggles, ReplacementAction, SquigglePageSlice},
};

use editor_input::{apply_viewport_input, handle_keyboard_input, handle_pointer_interaction};
use image::{
    handle_image_interaction, image_body_hit, image_drag_preview_rect, image_handle_hit,
    paint_image_on_page, paint_image_selection,
};
use page_layout::{layout_page_stack, PageLayout};
use palette::canvas_palette;
use table::{
    handle_table_interaction, paint_table, table_cell_hit, table_resize_handle_hit,
    TablePaintParams,
};

use header_footer::{
    header_footer_hit, paint_active_header_footer_editor, paint_page_header_footer,
};
pub(crate) use image_wrap::ImageLayout;
pub(crate) use layout::layout_document;

#[derive(Clone, Copy, Debug, Default)]
pub struct CanvasOutput {
    pub text_changed: bool,
    pub current_page: usize,
    pub page_count: usize,
}

pub fn paint_document_canvas(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    theme_mode: ThemeMode,
    history: &mut ChangeHistory,
    grammar_errors: &[GrammarError],
) -> CanvasOutput {
    let mut output = CanvasOutput::default();
    let palette = canvas_palette(theme_mode);
    let viewport = ui.available_rect_before_wrap();
    let editor_id = Id::new("document_canvas");
    let response = ui.interact(viewport, editor_id, Sense::click_and_drag());
    let painter = ui.painter_at(viewport);
    apply_viewport_input(ui, &response, canvas);
    if canvas.zoom_mode == crate::app::ZoomMode::FitPage {
        canvas.zoom = fit_page_zoom(viewport, document.default_page_setup().page_size);
    }

    painter.rect_filled(viewport, CornerRadius::ZERO, palette.canvas_bg);

    let default_setup = document.default_page_setup();
    let base_page_rect = centered_page_rect(
        viewport,
        default_setup.page_size,
        canvas.zoom,
        egui::Vec2::ZERO,
    );
    let content_size =
        section_page_content_rect(base_page_rect, default_setup, 14.0, 14.0, canvas.zoom).size();

    if canvas.active_header_footer.is_some()
        && ui.input(|input| input.key_pressed(egui::Key::Escape))
    {
        canvas.active_header_footer = None;
    }

    let has_focus = ui.memory(|mem| mem.has_focus(editor_id));
    if has_focus {
        ui.memory_mut(|mem| {
            mem.set_focus_lock_filter(
                editor_id,
                EventFilter {
                    tab: true,
                    horizontal_arrows: true,
                    vertical_arrows: true,
                    escape: false,
                },
            );
        });
    }

    let (mut document_layout, page_layout) = if has_focus && canvas.active_header_footer.is_none() {
        let dl = layout_document(ui, document, canvas, content_size.x);
        let changed = handle_keyboard_input(ui, document, canvas, &dl.galley, history);
        if changed {
            output.text_changed = true;
            let dl2 = layout_document(ui, document, canvas, content_size.x);
            let pl = layout_page_stack(
                viewport,
                document,
                canvas,
                &dl2.galley,
                &dl2.manual_page_break_rows,
                &dl2.paragraph_start_rows,
            );
            (dl2, pl)
        } else {
            let pl = layout_page_stack(
                viewport,
                document,
                canvas,
                &dl.galley,
                &dl.manual_page_break_rows,
                &dl.paragraph_start_rows,
            );
            (dl, pl)
        }
    } else {
        let dl = layout_document(ui, document, canvas, content_size.x);
        let pl = layout_page_stack(
            viewport,
            document,
            canvas,
            &dl.galley,
            &dl.manual_page_break_rows,
            &dl.paragraph_start_rows,
        );
        (dl, pl)
    };

    if response.double_clicked() {
        if let Some(pointer_pos) = response.interact_pointer_pos() {
            if let Some(active) = header_footer_hit(&page_layout, document, canvas, pointer_pos) {
                canvas.active_header_footer = Some(active);
                canvas.active_header_footer_cursor =
                    runs_total_chars(active_header_footer_runs(document, active));
                canvas.active_header_footer_selection =
                    CCursorRange::one(CCursor::new(canvas.active_header_footer_cursor));
                canvas.active_table_cell = None;
                canvas.selected_image_id = None;
                response.request_focus();
            } else if canvas.active_header_footer.is_some()
                && page_layout
                    .pages
                    .iter()
                    .any(|page| page.content_rect.contains(pointer_pos))
            {
                canvas.active_header_footer = None;
            }
        }
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
    let mut new_table_cell_rects: Vec<(usize, usize, usize, Rect)> = Vec::new();
    let mut new_table_cell_content_rects: Vec<(usize, usize, usize, Rect)> = Vec::new();
    let mut new_table_resize_handles: Vec<TableResizeHandleRect> = Vec::new();
    let move_preview = canvas.move_drag.as_ref().map(|drag| {
        (
            drag.image_id,
            drag.start_rect,
            drag.current_ptr - drag.start_ptr,
        )
    });

    for (page_index, page) in page_layout.pages.iter().enumerate() {
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

        paint_page_header_footer(
            &painter,
            document,
            canvas,
            page.page_rect,
            page.section_id,
            page.page_index_within_section,
            page.section_page_count,
            page_index + 1,
            page_layout.pages.len(),
            canvas.active_header_footer,
            palette.footer_text,
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
            let image_rect = Rect::from_min_size(
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
            );
            Some(image_drag_preview_rect(
                image.image.id,
                image_rect,
                move_preview,
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

        // Tables
        for table_layout in &document_layout.tables {
            let Some(row) = document_layout.galley.rows.get(table_layout.row_index) else {
                continue;
            };
            let table_y = row.pos.y;
            if table_y < page.start_y || table_y > page.end_y {
                continue;
            }

            let table_origin = egui::pos2(
                page.content_rect.left() + row.pos.x,
                page.content_rect.top() + table_y - page.start_y,
            );

            let active_cell = canvas.active_table_cell;
            let geometry = paint_table(
                ui,
                canvas,
                &page_clipped_painter,
                &table_layout.table,
                TablePaintParams {
                    origin: table_origin,
                    zoom,
                    active_cell,
                    time: ui.input(|i| i.time),
                },
            );
            new_table_cell_rects.extend(geometry.cell_rects);
            new_table_cell_content_rects.extend(geometry.cell_content_rects);
            new_table_resize_handles.extend(geometry.resize_handles);
        }

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

    if paint_active_header_footer_editor(
        ui,
        document,
        canvas,
        history,
        &page_layout,
        &response,
        editor_id,
    ) {
        output.text_changed = true;
    }

    let squiggle_pages: Vec<SquigglePageSlice> = page_layout
        .pages
        .iter()
        .map(|page| SquigglePageSlice {
            content_rect: page.content_rect,
            start_y: page.start_y,
            end_y: page.end_y,
        })
        .collect();
    let pending_replacement = paint_grammar_squiggles(
        ui,
        &painter,
        &document_layout.galley,
        &squiggle_pages,
        grammar_errors,
    );

    canvas.image_rects = new_image_rects;
    canvas.table_cell_rects = new_table_cell_rects;
    canvas.table_cell_content_rects = new_table_cell_content_rects;
    canvas.table_resize_handles = new_table_resize_handles;

    let (table_pointer_captured, table_document_changed) = if canvas.active_header_footer.is_some()
    {
        (true, false)
    } else {
        handle_table_interaction(ui, &response, canvas, document, history)
    };
    output.text_changed |= table_document_changed;

    let (image_pointer_captured, image_document_changed) = if table_pointer_captured {
        (true, false)
    } else {
        handle_image_interaction(ui, &response, canvas, document, history)
    };
    output.text_changed |= image_document_changed;

    if !table_pointer_captured && !image_pointer_captured && canvas.active_header_footer.is_none() {
        handle_pointer_interaction(
            ui,
            &response,
            &page_layout,
            &document_layout.galley,
            canvas,
            document,
        );
    }
    update_canvas_hover_cursor(ui, &response, canvas, &page_layout);

    // Draw selection border + handles with unclipped painter so they aren't cut at page margins
    if let Some((_, selected_rect)) = canvas
        .image_rects
        .iter()
        .find(|(id, _)| Some(*id) == canvas.selected_image_id)
    {
        paint_image_selection(&painter, *selected_rect);
    }

    if has_focus
        && canvas.selected_image_id.is_none()
        && canvas.active_table_cell.is_none()
        && canvas.active_header_footer.is_none()
    {
        if let Some(caret_rect) = page_layout.caret_rect(
            &document_layout.galley,
            canvas.selection.primary,
            caret_height(canvas.active_style, canvas.zoom),
        ) {
            paint_text_cursor(
                ui,
                &painter,
                caret_rect,
                ui.input(|i| i.time) - canvas.last_interaction_time,
            );

            if canvas.ai_working {
                let spinner_rect = egui::Rect::from_min_size(
                    caret_rect.max + egui::vec2(4.0, -caret_rect.height() * 0.8),
                    egui::vec2(caret_rect.height() * 0.8, caret_rect.height() * 0.8),
                );
                ui.put(
                    spinner_rect,
                    egui::Spinner::new()
                        .size(caret_rect.height() * 0.8)
                        .color(palette.page_border),
                );
            }
        }
    }

    output.page_count = page_layout.pages.len();
    output.current_page =
        page_layout.current_page(&document_layout.galley, canvas.selection.primary);

    if let Some(replacement) = pending_replacement {
        if apply_grammar_replacement(document, canvas, history, ui, replacement) {
            output.text_changed = true;
        }
    }

    output
}

fn update_canvas_hover_cursor(
    ui: &egui::Ui,
    response: &egui::Response,
    canvas: &CanvasState,
    page_layout: &PageLayout,
) {
    if response.dragged() {
        return;
    }

    let Some(hover_pos) = ui.ctx().pointer_hover_pos() else {
        return;
    };
    if !response.rect.contains(hover_pos) {
        return;
    }

    let cursor_icon = if let Some(handle) = table_resize_handle_hit(canvas, hover_pos) {
        match handle.kind {
            TableResizeKind::Column { .. } => egui::CursorIcon::ResizeEast,
            TableResizeKind::Row { .. } => egui::CursorIcon::ResizeSouth,
        }
    } else if table_cell_hit(canvas, hover_pos).is_some() {
        egui::CursorIcon::Text
    } else if let Some((_, handle)) = image_handle_hit(canvas, hover_pos, 6.0) {
        match handle {
            ResizeHandle::NW | ResizeHandle::SE => egui::CursorIcon::ResizeNwSe,
            ResizeHandle::NE | ResizeHandle::SW => egui::CursorIcon::ResizeNeSw,
            ResizeHandle::N | ResizeHandle::S => egui::CursorIcon::ResizeSouth,
            ResizeHandle::E | ResizeHandle::W => egui::CursorIcon::ResizeEast,
        }
    } else if image_body_hit(canvas, hover_pos).is_some() {
        egui::CursorIcon::Grab
    } else if page_layout.document_pos(hover_pos).is_some() {
        egui::CursorIcon::Text
    } else {
        return;
    };

    ui.ctx().set_cursor_icon(cursor_icon);
}

fn caret_height(style: CharacterStyle, zoom: f32) -> f32 {
    style.font_size_points.max(1.0) * zoom * 1.25
}

fn apply_grammar_replacement(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    ui: &egui::Ui,
    replacement: ReplacementAction,
) -> bool {
    let text = document.plain_text();
    if replacement.byte_start >= text.len() || replacement.byte_start > replacement.byte_end {
        return false;
    }

    let start_char = byte_to_char_index(&text, replacement.byte_start);
    let end_char = byte_to_char_index(&text, replacement.byte_end).max(start_char);
    let style = document.style_at(start_char);

    let now = ui.input(|i| i.time);
    history.checkpoint(document, now);
    document.delete_range(start_char..end_char);
    document.insert_text(start_char, &replacement.replacement, style);

    let cursor_char = start_char + replacement.replacement.chars().count();
    canvas.selection = egui::text_selection::CCursorRange::one(CCursor::new(cursor_char));
    canvas.active_style = document.typing_style_at(cursor_char);
    canvas.active_paragraph_style = document.paragraph_style_at(cursor_char);
    canvas.last_interaction_time = now;
    true
}

fn byte_to_char_index(text: &str, byte_offset: usize) -> usize {
    let clamped = byte_offset.min(text.len());
    let mut count = 0usize;
    for (idx, _) in text.char_indices() {
        if idx >= clamped {
            break;
        }
        count += 1;
    }
    if clamped == text.len() {
        text.chars().count()
    } else {
        count
    }
}

fn runs_total_chars(runs: &[TextRun]) -> usize {
    runs.iter().map(|run| run.text.chars().count()).sum()
}

fn active_header_footer_runs(document: &DocumentState, active: ActiveHeaderFooter) -> &[TextRun] {
    document
        .resolve_header_footer_slot(active.section_id, active.kind, active.variant)
        .story
        .runs
        .as_slice()
}
