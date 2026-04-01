use std::sync::Arc;

use eframe::egui::{
    self, epaint::text::cursor::CCursor, text_selection::CCursorRange, Event, Key, Modifiers,
};

use crate::{app::CanvasState, document::DocumentState};

use super::page_layout::PageLayout;

pub(super) fn apply_viewport_input(
    ui: &mut egui::Ui,
    response: &egui::Response,
    canvas: &mut CanvasState,
) {
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

pub(super) fn handle_pointer_interaction(
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

        let handled_multi_click = if response.triple_clicked() {
            canvas.selection =
                selection_range_from_char_range(document.line_range_at(cursor.index));
            canvas.selection.h_pos = None;
            canvas.active_style = document.typing_style_at(canvas.selection.primary.index);
            canvas.active_paragraph_style =
                document.paragraph_style_at(canvas.selection.primary.index);
            canvas.last_interaction_time = ui.input(|i| i.time);
            true
        } else if response.double_clicked() {
            let range = document
                .word_range_at(cursor.index)
                .unwrap_or_else(|| document.line_range_at(cursor.index));
            canvas.selection = selection_range_from_char_range(range);
            canvas.selection.h_pos = None;
            canvas.active_style = document.typing_style_at(canvas.selection.primary.index);
            canvas.active_paragraph_style =
                document.paragraph_style_at(canvas.selection.primary.index);
            canvas.last_interaction_time = ui.input(|i| i.time);
            true
        } else if response.drag_started() {
            let extend = ui.input(|i| i.modifiers.shift);
            if extend {
                canvas.selection.primary = cursor;
            } else {
                canvas.selection = CCursorRange::one(cursor);
            }
            canvas.selection.h_pos = None;
            canvas.last_interaction_time = ui.input(|i| i.time);
            false
        } else {
            false
        };

        if response.dragged() {
            canvas.selection.primary = cursor;
            canvas.selection.h_pos = None;
            canvas.active_style = document.typing_style_at(canvas.selection.primary.index);
            canvas.active_paragraph_style =
                document.paragraph_style_at(canvas.selection.primary.index);
            canvas.last_interaction_time = ui.input(|i| i.time);
        } else if response.clicked() && !handled_multi_click {
            if ui.input(|i| i.modifiers.shift) {
                canvas.selection.primary = cursor;
            } else {
                canvas.selection = CCursorRange::one(cursor);
            }
            canvas.selection.h_pos = None;
            canvas.active_style = document.typing_style_at(canvas.selection.primary.index);
            canvas.active_paragraph_style =
                document.paragraph_style_at(canvas.selection.primary.index);
            canvas.last_interaction_time = ui.input(|i| i.time);
        }
    }
}

fn selection_range_from_char_range(range: std::ops::Range<usize>) -> CCursorRange {
    CCursorRange::two(CCursor::new(range.start), CCursor::new(range.end))
}

pub(super) fn handle_keyboard_input(
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
                            canvas.active_paragraph_style =
                                document.paragraph_style_at(canvas.selection.primary.index);
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
        canvas.active_paragraph_style = document.paragraph_style_at(canvas.selection.primary.index);
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
        canvas.active_paragraph_style = document.paragraph_style_at(new_line_start);
        return;
    }

    let transformed_next_index = document.apply_markdown_shortcuts_at(next_index);
    let line_end = document.line_range_at(transformed_next_index).end;
    let cursor_index = transformed_next_index.min(line_end);
    canvas.selection = CCursorRange::one(CCursor::new(cursor_index));
    canvas.active_style = document.typing_style_at(cursor_index);
    canvas.active_paragraph_style = document.paragraph_style_at(cursor_index);
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
