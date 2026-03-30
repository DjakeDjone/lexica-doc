use std::sync::Arc;

use eframe::egui::{
    self,
    epaint::{text::cursor::CCursor, CornerRadius},
    text_selection::{
        visuals::{paint_text_cursor, paint_text_selection},
        CCursorRange,
    },
    Color32, Event, FontFamily, FontId, Id, Key, Modifiers, Rect, Sense, Stroke, StrokeKind,
};

use crate::{
    app::{CanvasState, ThemeMode},
    document::DocumentState,
    layout::{
        centered_page_rect, document_points_to_pixels, document_points_to_screen_points,
        page_content_rect,
    },
};

#[derive(Clone, Copy)]
struct CanvasPalette {
    canvas_bg: Color32,
    page_bg: Color32,
    page_border: Color32,
    page_focus: Color32,
    page_shadow: Color32,
    footer_bg: Color32,
    footer_stroke: Color32,
    footer_text: Color32,
}

fn canvas_palette(mode: ThemeMode) -> CanvasPalette {
    match mode {
        ThemeMode::Light => CanvasPalette {
            canvas_bg: Color32::from_rgb(217, 219, 223),
            page_bg: Color32::from_rgb(255, 255, 255),
            page_border: Color32::from_rgb(186, 193, 204),
            page_focus: Color32::from_rgb(79, 129, 209),
            page_shadow: Color32::from_black_alpha(24),
            footer_bg: Color32::from_rgba_premultiplied(241, 244, 248, 220),
            footer_stroke: Color32::from_rgba_premultiplied(173, 183, 198, 200),
            footer_text: Color32::from_rgb(88, 96, 108),
        },
        ThemeMode::Dark => CanvasPalette {
            canvas_bg: Color32::from_rgb(45, 49, 56),
            page_bg: Color32::from_rgb(250, 250, 252),
            page_border: Color32::from_rgb(124, 132, 147),
            page_focus: Color32::from_rgb(126, 171, 236),
            page_shadow: Color32::from_black_alpha(54),
            footer_bg: Color32::from_rgba_premultiplied(29, 33, 40, 220),
            footer_stroke: Color32::from_rgba_premultiplied(98, 108, 126, 200),
            footer_text: Color32::from_rgb(196, 206, 224),
        },
    }
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

fn apply_viewport_input(ui: &mut egui::Ui, response: &egui::Response, canvas: &mut CanvasState) {
    if !(response.hovered() || response.has_focus()) {
        return;
    }

    let scroll_delta = ui.input(|input| input.smooth_scroll_delta());
    if ui.input(|input| input.modifiers.command) {
        let zoom_delta = ui.input(|input| input.zoom_delta());
        if zoom_delta != 1.0 {
            canvas.zoom = (canvas.zoom * zoom_delta).clamp(0.5, 3.0);
        }
    } else if scroll_delta != egui::Vec2::ZERO {
        canvas.pan += egui::vec2(scroll_delta.x, scroll_delta.y);
    }
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

fn handle_pointer_interaction(
    ui: &mut egui::Ui,
    response: &egui::Response,
    page_layout: &PageLayout,
    galley: &Arc<egui::Galley>,
    canvas: &mut CanvasState,
    document: &DocumentState,
) {
    if response.clicked() || response.drag_started() {
        response.request_focus();
    }

    if let Some(pointer_pos) = response.interact_pointer_pos() {
        let Some(local_pos) = page_layout.document_pos(pointer_pos) else {
            return;
        };
        let cursor = galley.cursor_from_pos(local_pos);

        if response.drag_started() {
            let extend = ui.input(|i| i.modifiers.shift);
            if extend {
                canvas.selection.primary = cursor;
            } else {
                canvas.selection = CCursorRange::one(cursor);
            }
            canvas.selection.h_pos = None;
            canvas.last_interaction_time = ui.input(|i| i.time);
        }

        if response.dragged() {
            canvas.selection.primary = cursor;
            canvas.selection.h_pos = None;
            canvas.active_style = document.typing_style_at(canvas.selection.primary.index);
            canvas.last_interaction_time = ui.input(|i| i.time);
        } else if response.clicked() {
            if ui.input(|i| i.modifiers.shift) {
                canvas.selection.primary = cursor;
            } else {
                canvas.selection = CCursorRange::one(cursor);
            }
            canvas.selection.h_pos = None;
            canvas.active_style = document.typing_style_at(canvas.selection.primary.index);
            canvas.last_interaction_time = ui.input(|i| i.time);
        }
    }
}

fn handle_keyboard_input(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    galley: &Arc<egui::Galley>,
) -> bool {
    let os = ui.ctx().os();
    let events = ui.input(|i| i.events.clone());
    let mut changed = false;

    for event in events {
        match event {
            Event::Text(text) if !text.is_empty() => {
                replace_selection_or_insert(document, canvas, &text);
                changed = true;
            }
            Event::Paste(text) => {
                replace_selection_or_insert(document, canvas, &text);
                changed = true;
            }
            Event::Copy => {
                let selected = canvas.selection.as_sorted_char_range();
                if selected.start < selected.end {
                    ui.copy_text(document.selected_text(selected));
                }
            }
            Event::Cut => {
                let selected = canvas.selection.as_sorted_char_range();
                if selected.start < selected.end {
                    ui.copy_text(document.selected_text(selected.clone()));
                    document.delete_range(selected.clone());
                    canvas.selection = CCursorRange::one(CCursor::new(selected.start));
                    changed = true;
                }
            }
            Event::Key {
                key,
                pressed: true,
                modifiers,
                ..
            } => {
                if handle_shortcut_key(document, canvas, key, modifiers) {
                    changed = true;
                    continue;
                }

                match key {
                    Key::PageUp => {
                        canvas.pan.y += 120.0;
                    }
                    Key::PageDown => {
                        canvas.pan.y -= 120.0;
                    }
                    Key::Backspace => changed |= delete_backward(document, canvas),
                    Key::Delete => changed |= delete_forward(document, canvas),
                    Key::Enter => {
                        replace_selection_or_insert(document, canvas, "\n");
                        changed = true;
                    }
                    Key::Tab => {
                        replace_selection_or_insert(document, canvas, "    ");
                        changed = true;
                    }
                    Key::ArrowLeft
                    | Key::ArrowRight
                    | Key::ArrowUp
                    | Key::ArrowDown
                    | Key::Home
                    | Key::End
                    | Key::A => {
                        if canvas.selection.on_key_press(os, galley, &modifiers, key) {
                            canvas.active_style =
                                document.typing_style_at(canvas.selection.primary.index);
                            canvas.last_interaction_time = ui.input(|i| i.time);
                        }
                    }
                    _ => {}
                }
            }
            _ => {}
        }
    }

    if changed {
        canvas.active_style = document.typing_style_at(canvas.selection.primary.index);
        canvas.last_interaction_time = ui.input(|i| i.time);
    }

    changed
}

fn handle_shortcut_key(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    key: Key,
    modifiers: Modifiers,
) -> bool {
    if !modifiers.command {
        return false;
    }

    let range = canvas.selection.as_sorted_char_range();
    match key {
        Key::B => {
            let next = !canvas.active_style.bold;
            if range.start < range.end {
                document.apply_style_to_range(range, |style| style.bold = next);
            }
            canvas.active_style.bold = next;
            true
        }
        Key::I => {
            let next = !canvas.active_style.italic;
            if range.start < range.end {
                document.apply_style_to_range(range, |style| style.italic = next);
            }
            canvas.active_style.italic = next;
            true
        }
        Key::U => {
            let next = !canvas.active_style.underline;
            if range.start < range.end {
                document.apply_style_to_range(range, |style| style.underline = next);
            }
            canvas.active_style.underline = next;
            true
        }
        _ => false,
    }
}

fn replace_selection_or_insert(document: &mut DocumentState, canvas: &mut CanvasState, text: &str) {
    let selected = canvas.selection.as_sorted_char_range();
    let insert_at = selected.start;
    if selected.start < selected.end {
        document.delete_range(selected);
    }
    document.insert_text(insert_at, text, canvas.active_style);
    let inserted_chars = text.chars().count();
    let next_index = insert_at + inserted_chars;

    let transformed_insert_at = document.apply_markdown_shortcuts_at(insert_at);
    if text.contains('\n') {
        let new_line_start = document
            .line_range_at(transformed_insert_at)
            .end
            .saturating_add(1);
        canvas.selection = CCursorRange::one(CCursor::new(new_line_start));
        canvas.active_style = document.typing_style_at(new_line_start);
        return;
    }

    let transformed_next_index = document.apply_markdown_shortcuts_at(next_index);
    let line_end = document.line_range_at(transformed_next_index).end;
    let cursor_index = transformed_next_index.min(line_end);
    canvas.selection = CCursorRange::one(CCursor::new(cursor_index));
    canvas.active_style = document.typing_style_at(cursor_index);
}

fn delete_backward(document: &mut DocumentState, canvas: &mut CanvasState) -> bool {
    let selected = canvas.selection.as_sorted_char_range();
    if selected.start < selected.end {
        document.delete_range(selected.clone());
        canvas.selection = CCursorRange::one(CCursor::new(selected.start));
        return true;
    }

    if canvas.selection.primary.index == 0 {
        return false;
    }

    let delete_start = canvas.selection.primary.index - 1;
    document.delete_range(delete_start..canvas.selection.primary.index);
    canvas.selection = CCursorRange::one(CCursor::new(delete_start));
    true
}

fn delete_forward(document: &mut DocumentState, canvas: &mut CanvasState) -> bool {
    let selected = canvas.selection.as_sorted_char_range();
    if selected.start < selected.end {
        document.delete_range(selected.clone());
        canvas.selection = CCursorRange::one(CCursor::new(selected.start));
        return true;
    }

    let total_chars = document.total_chars();
    if canvas.selection.primary.index >= total_chars {
        return false;
    }

    let delete_start = canvas.selection.primary.index;
    document.delete_range(delete_start..delete_start + 1);
    canvas.selection = CCursorRange::one(CCursor::new(delete_start));
    true
}

fn caret_rect(galley: &egui::Galley, cursor: CCursor) -> Rect {
    let layout_cursor = galley.layout_from_cursor(cursor);
    let mut rect = galley.pos_from_cursor(cursor);
    if let Some(row) = galley.rows.get(layout_cursor.row) {
        let height = row.row.height();
        rect.min.y = row.pos.y;
        rect.max.y = row.pos.y + height;
    }
    rect.expand2(egui::vec2(0.75, 0.75))
}

struct PageSlice {
    page_rect: Rect,
    content_rect: Rect,
    start_y: f32,
    end_y: f32,
}

struct PageLayout {
    pages: Vec<PageSlice>,
}

impl PageLayout {
    fn document_pos(&self, pointer_pos: egui::Pos2) -> Option<egui::Vec2> {
        self.pages.iter().find_map(|page| {
            if page.content_rect.contains(pointer_pos) {
                let local_y = pointer_pos.y - page.content_rect.top();
                Some(egui::vec2(
                    pointer_pos.x - page.content_rect.left(),
                    (page.start_y + local_y).clamp(page.start_y, page.end_y),
                ))
            } else {
                None
            }
        })
    }

    fn caret_rect(&self, galley: &egui::Galley, cursor: CCursor) -> Option<Rect> {
        let document_rect = caret_rect(galley, cursor);
        self.pages.iter().find_map(|page| {
            if document_rect.center().y >= page.start_y && document_rect.center().y <= page.end_y {
                Some(
                    document_rect
                        .translate(page.content_rect.min.to_vec2() - egui::vec2(0.0, page.start_y)),
                )
            } else {
                None
            }
        })
    }
}

fn layout_page_stack(
    viewport: Rect,
    document: &DocumentState,
    canvas: &CanvasState,
    galley: &egui::Galley,
) -> PageLayout {
    let page_gap = document_points_to_screen_points(24.0, canvas.zoom);
    let base_page_rect =
        centered_page_rect(viewport, document.page_size, canvas.zoom, egui::Vec2::ZERO);
    let page_size = base_page_rect.size();
    let content_height = page_content_rect(base_page_rect, document.margins, canvas.zoom).height();
    let page_ranges = compute_page_ranges(galley, content_height);
    let page_count = page_ranges.len().max(1);
    let stack_height =
        page_count as f32 * page_size.y + (page_count.saturating_sub(1) as f32 * page_gap);

    let top = if stack_height < viewport.height() {
        viewport.center().y - stack_height * 0.5 + canvas.pan.y
    } else {
        viewport.top() + document_points_to_screen_points(24.0, canvas.zoom) + canvas.pan.y
    };
    let left = viewport.center().x - page_size.x * 0.5 + canvas.pan.x;

    let mut pages = Vec::with_capacity(page_count);
    for (index, (start_y, end_y)) in page_ranges.into_iter().enumerate() {
        let min = egui::pos2(left, top + index as f32 * (page_size.y + page_gap));
        let page_rect = Rect::from_min_size(min, page_size);
        let content_rect = page_content_rect(page_rect, document.margins, canvas.zoom);
        pages.push(PageSlice {
            page_rect,
            content_rect,
            start_y,
            end_y,
        });
    }

    PageLayout { pages }
}

fn compute_page_ranges(galley: &egui::Galley, page_height: f32) -> Vec<(f32, f32)> {
    if galley.rows.is_empty() {
        return vec![(0.0, page_height)];
    }

    let mut pages = Vec::new();
    let mut page_start: f32 = 0.0;
    let mut last_row_end: f32 = 0.0;

    for row in &galley.rows {
        let row_start = row.pos.y;
        let row_end = row.pos.y + row.row.height();

        if row_end - page_start > page_height && row_start > page_start {
            pages.push((page_start, last_row_end.max(page_start)));
            page_start = row_start;
        }

        last_row_end = row_end;
    }

    pages.push((page_start, last_row_end.max(page_start)));
    pages
}
