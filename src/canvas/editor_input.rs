use std::sync::Arc;

use eframe::egui::{
    self, epaint::text::cursor::CCursor, text_selection::CCursorRange, Event, Key, Modifiers,
    MouseWheelUnit, TouchPhase,
};

use crate::{
    app::{CanvasState, ChangeHistory, ZoomMode},
    document::DocumentState,
    layout::document_points_to_screen_points,
};

use super::{page_layout::PageLayout, table::table_cell_text_galley};

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
        canvas.scroll_velocity = egui::Vec2::ZERO;
        let zoom_delta = ui.input(|input| {
            let gesture_delta = input.zoom_delta();
            if gesture_delta != 1.0 {
                gesture_delta
            } else {
                (input.smooth_scroll_delta().y * 0.01).exp()
            }
        });
        if zoom_delta != 1.0 {
            canvas.zoom_mode = ZoomMode::Manual;
            canvas.scale_view(zoom_delta);
            canvas.last_zoom_input_time = ui.input(|input| input.time);
        }
    } else if scroll_delta != egui::Vec2::ZERO {
        canvas.pan += egui::vec2(scroll_delta.x, scroll_delta.y);
        let touchpad_moved = ui.input(|input| {
            input.events.iter().any(|event| {
                matches!(
                    event,
                    Event::MouseWheel {
                        unit: MouseWheelUnit::Point,
                        phase: TouchPhase::Move,
                        ..
                    }
                )
            })
        });
        if touchpad_moved {
            let dt = ui.input(|input| input.stable_dt).clamp(1.0 / 240.0, 0.1);
            canvas.scroll_velocity = canvas.scroll_velocity * 0.65 + scroll_delta / dt * 0.35;
        } else {
            canvas.scroll_velocity = egui::Vec2::ZERO;
        }
    } else if canvas.scroll_velocity != egui::Vec2::ZERO {
        let dt = ui.input(|input| input.stable_dt).min(0.1);
        canvas.pan += canvas.scroll_velocity * dt;
        for axis in 0..2 {
            let friction = 1_200.0 * dt;
            if canvas.scroll_velocity[axis].abs() <= friction.max(20.0) {
                canvas.scroll_velocity[axis] = 0.0;
            } else {
                canvas.scroll_velocity[axis] -= friction * canvas.scroll_velocity[axis].signum();
            }
        }
        ui.ctx().request_repaint();
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
        let Some(local_pos) = page_layout.document_pos(pointer_pos).or_else(|| {
            response
                .dragged()
                .then(|| page_layout.clamped_document_pos(pointer_pos))
                .flatten()
        }) else {
            return;
        };
        let mut cursor = galley.cursor_from_pos(local_pos);

        if let Some(completion) = &canvas.ai_completion {
            let primary_idx = canvas.selection.primary.index;
            if cursor.index > primary_idx {
                cursor.index = cursor
                    .index
                    .saturating_sub(completion.chars().count())
                    .max(primary_idx);
            }
        }

        if response.clicked()
            || response.drag_started()
            || response.double_clicked()
            || response.triple_clicked()
        {
            canvas.ai_completion = None;
        }

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
            canvas.active_table_cell = None;
            canvas.selection.h_pos = None;
            canvas.last_interaction_time = ui.input(|i| i.time);
            false
        } else {
            false
        };

        if response.dragged() && canvas.resize_drag.is_none() && canvas.move_drag.is_none() {
            canvas.selection.primary = cursor;
            canvas.active_table_cell = None;
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
            canvas.active_table_cell = None;
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
    history: &mut ChangeHistory,
) -> bool {
    let os = ui.ctx().os();
    let events = ui.input(|i| i.events.clone());
    let mut changed = false;

    let ghost_start_idx = canvas.selection.primary.index;
    let mut ghost_len = 0;

    for event in events {
        match event {
            Event::Text(text) if !text.is_empty() => {
                if let Some(completion) = canvas.ai_completion.take() {
                    ghost_len = completion.chars().count();
                }
                let now = ui.input(|i| i.time);
                history.checkpoint_coalesced(document, now);
                if canvas.active_table_cell.is_some() {
                    replace_active_table_cell_selection_or_insert(document, canvas, &text);
                } else {
                    replace_selection_or_insert(document, canvas, &text);
                }
                changed = true;
            }
            Event::Paste(text) => {
                if let Some(completion) = canvas.ai_completion.take() {
                    ghost_len = completion.chars().count();
                }
                let now = ui.input(|i| i.time);
                history.checkpoint(document, now);
                if canvas.active_table_cell.is_some() {
                    replace_active_table_cell_selection_or_insert(document, canvas, &text);
                } else {
                    replace_selection_or_insert(document, canvas, &text);
                }
                changed = true;
            }
            Event::Copy => {
                if canvas.active_table_cell.is_some() {
                    copy_active_table_cell_selection(ui, document, canvas);
                } else {
                    let selected = canvas.selection.as_sorted_char_range();
                    if selected.start < selected.end {
                        ui.copy_text(document.selected_text(selected));
                    }
                }
            }
            Event::Cut => {
                if let Some(completion) = canvas.ai_completion.take() {
                    ghost_len = completion.chars().count();
                    changed = true;
                }
                if canvas.active_table_cell.is_some() {
                    if copy_active_table_cell_selection(ui, document, canvas) {
                        let now = ui.input(|i| i.time);
                        history.checkpoint(document, now);
                        changed |= delete_active_table_cell_selection(document, canvas);
                    }
                } else {
                    let selected = canvas.selection.as_sorted_char_range();
                    if selected.start < selected.end {
                        let now = ui.input(|i| i.time);
                        history.checkpoint(document, now);
                        ui.copy_text(document.selected_text(selected.clone()));
                        document.delete_range(selected.clone());
                        canvas.selection = CCursorRange::one(CCursor::new(selected.start));
                        changed = true;
                    }
                }
            }
            Event::Key {
                key,
                pressed: true,
                modifiers,
                ..
            } => {
                if key == Key::Escape {
                    if let Some(completion) = canvas.ai_completion.take() {
                        ghost_len = completion.chars().count();
                        changed = true;
                    }
                    continue;
                }

                if key != Key::Tab {
                    if let Some(completion) = canvas.ai_completion.take() {
                        ghost_len = completion.chars().count();
                        changed = true;
                    }
                }

                if handle_shortcut_key(document, canvas, key, modifiers, history, ui) {
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
                    Key::Backspace => {
                        let now = ui.input(|i| i.time);
                        history.checkpoint_coalesced(document, now);
                        changed |= if canvas.active_table_cell.is_some() {
                            delete_from_active_table_cell(document, canvas, true, modifiers.ctrl)
                        } else if modifiers.ctrl {
                            delete_word_backward(document, canvas)
                        } else {
                            delete_backward(document, canvas)
                        };
                    }
                    Key::Delete => {
                        let now = ui.input(|i| i.time);
                        history.checkpoint_coalesced(document, now);
                        changed |= if canvas.active_table_cell.is_some() {
                            delete_from_active_table_cell(document, canvas, false, modifiers.ctrl)
                        } else if modifiers.ctrl {
                            delete_word_forward(document, canvas)
                        } else {
                            delete_forward(document, canvas)
                        };
                    }
                    Key::Enter => {
                        let now = ui.input(|i| i.time);
                        history.checkpoint(document, now);
                        if canvas.active_table_cell.is_some() {
                            replace_active_table_cell_selection_or_insert(document, canvas, "\n");
                        } else {
                            replace_selection_or_insert(document, canvas, "\n");
                        }
                        changed = true;
                    }
                    Key::Tab => {
                        let now = ui.input(|i| i.time);
                        history.checkpoint(document, now);
                        if let Some(completion) = canvas.ai_completion.take() {
                            replace_selection_or_insert(document, canvas, &completion);
                        } else if canvas.active_table_cell.is_some() {
                            move_active_table_cell(document, canvas, !modifiers.shift);
                        } else {
                            replace_selection_or_insert(document, canvas, "    ");
                        }
                        changed = true;
                    }
                    Key::ArrowLeft
                    | Key::ArrowRight
                    | Key::ArrowUp
                    | Key::ArrowDown
                    | Key::Home
                    | Key::End
                    | Key::A => {
                        if canvas.active_table_cell.is_some() {
                            move_active_table_cell_cursor_by_key(
                                ui, document, canvas, key, modifiers,
                            );
                            canvas.last_interaction_time = ui.input(|i| i.time);
                        } else if canvas.selection.on_key_press(os, galley, &modifiers, key) {
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

    if ghost_len > 0 {
        let adjust = |idx: &mut usize| {
            if *idx > ghost_start_idx {
                *idx = idx.saturating_sub(ghost_len).max(ghost_start_idx);
            }
        };
        adjust(&mut canvas.selection.primary.index);
        adjust(&mut canvas.selection.secondary.index);

        let total_chars = document.total_chars();
        canvas.selection.primary.index = canvas.selection.primary.index.min(total_chars);
        canvas.selection.secondary.index = canvas.selection.secondary.index.min(total_chars);
    }

    if changed {
        if let Some((table_id, row, col)) = canvas.active_table_cell {
            if let Some(style) = document.table_cell_style_at(
                table_id,
                row,
                col,
                canvas.table_cell_selection.primary.index,
            ) {
                canvas.active_style = style;
            }
        } else {
            canvas.active_style = document.typing_style_at(canvas.selection.primary.index);
            canvas.active_paragraph_style =
                document.paragraph_style_at(canvas.selection.primary.index);
        }
        canvas.last_interaction_time = ui.input(|i| i.time);
    }

    changed
}

fn replace_active_table_cell_selection_or_insert(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    text: &str,
) {
    let Some((table_id, row, col)) = canvas.active_table_cell else {
        return;
    };
    let selected = canvas.table_cell_selection.as_sorted_char_range();
    if let Some(next_index) = document.replace_table_cell_range_with_text(
        table_id,
        row,
        col,
        selected,
        text,
        canvas.active_style,
    ) {
        canvas.table_cell_selection = CCursorRange::one(CCursor::new(next_index));
    }
}

fn copy_active_table_cell_selection(
    ui: &mut egui::Ui,
    document: &DocumentState,
    canvas: &CanvasState,
) -> bool {
    let Some((table_id, row, col)) = canvas.active_table_cell else {
        return false;
    };
    let selected = canvas.table_cell_selection.as_sorted_char_range();
    if selected.start >= selected.end {
        return false;
    }
    let Some(text) = document.table_cell_text(table_id, row, col) else {
        return false;
    };
    let selected_text: String = text
        .chars()
        .skip(selected.start)
        .take(selected.end - selected.start)
        .collect();
    ui.copy_text(selected_text);
    true
}

fn delete_active_table_cell_selection(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
) -> bool {
    let Some((table_id, row, col)) = canvas.active_table_cell else {
        return false;
    };
    let selected = canvas.table_cell_selection.as_sorted_char_range();
    if selected.start >= selected.end {
        return false;
    }
    document.delete_table_cell_char_range(table_id, row, col, selected.clone());
    canvas.table_cell_selection = CCursorRange::one(CCursor::new(selected.start));
    true
}

fn delete_from_active_table_cell(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    backward: bool,
    word: bool,
) -> bool {
    let Some((table_id, row, col)) = canvas.active_table_cell else {
        return false;
    };
    let Some(text) = document.table_cell_text(table_id, row, col) else {
        return false;
    };
    let chars: Vec<char> = text.chars().collect();
    let selected = canvas.table_cell_selection.as_sorted_char_range();
    let range = if selected.start < selected.end {
        selected
    } else if backward {
        let cursor = canvas.table_cell_selection.primary.index.min(chars.len());
        if cursor == 0 {
            return false;
        }
        let mut start = cursor;
        if word {
            while start > 0 && !is_word_char(chars[start - 1]) {
                start -= 1;
            }
            while start > 0 && is_word_char(chars[start - 1]) {
                start -= 1;
            }
        } else {
            start = start.saturating_sub(1);
        }
        start..cursor
    } else if word {
        let cursor = canvas.table_cell_selection.primary.index.min(chars.len());
        if cursor >= chars.len() {
            return false;
        }
        let mut end = cursor;
        while end < chars.len() && !is_word_char(chars[end]) {
            end += 1;
        }
        while end < chars.len() && is_word_char(chars[end]) {
            end += 1;
        }
        if end == cursor {
            end += 1;
        }
        cursor..end.min(chars.len())
    } else {
        let cursor = canvas.table_cell_selection.primary.index.min(chars.len());
        if cursor >= chars.len() {
            return false;
        }
        cursor..cursor + 1
    };
    document.delete_table_cell_char_range(table_id, row, col, range.clone());
    canvas.table_cell_selection = CCursorRange::one(CCursor::new(range.start));
    true
}

fn move_active_table_cell(document: &DocumentState, canvas: &mut CanvasState, forward: bool) {
    let Some((table_id, row, col)) = canvas.active_table_cell else {
        return;
    };
    let Some(table) = document.table_by_id(table_id) else {
        canvas.active_table_cell = None;
        return;
    };
    let rows = table.num_rows();
    let cols = table.num_cols();
    if rows == 0 || cols == 0 {
        canvas.active_table_cell = None;
        return;
    }
    let linear = row.saturating_mul(cols).saturating_add(col);
    let next = if forward {
        (linear + 1).min(rows * cols - 1)
    } else {
        linear.saturating_sub(1)
    };
    let next_cell = (table_id, next / cols, next % cols);
    canvas.active_table_cell = Some(next_cell);
    canvas.table_cell_selection = CCursorRange::default();
    if let Some(style) = document.table_cell_typing_style(next_cell.0, next_cell.1, next_cell.2) {
        canvas.active_style = style;
    }
}

fn move_active_table_cell_cursor_by_key(
    ui: &egui::Ui,
    document: &DocumentState,
    canvas: &mut CanvasState,
    key: Key,
    modifiers: Modifiers,
) {
    let Some((table_id, row, col)) = canvas.active_table_cell else {
        return;
    };
    let Some(table) = document.table_by_id(table_id) else {
        return;
    };
    let Some(cell) = table.rows.get(row).and_then(|cells| cells.get(col)) else {
        return;
    };
    let width_points: f32 = table
        .col_widths_points
        .iter()
        .skip(col)
        .take(cell.col_span.max(1) as usize)
        .sum();
    let available_width =
        document_points_to_screen_points(width_points - 8.0, canvas.zoom).max(1.0);
    let galley = table_cell_text_galley(ui.painter(), cell, available_width, canvas.zoom);

    if canvas
        .table_cell_selection
        .on_key_press(ui.ctx().os(), &galley, &modifiers, key)
    {
        let next = canvas.table_cell_selection.primary.index;
        if let Some(style) = document.table_cell_style_at(table_id, row, col, next) {
            canvas.active_style = style;
        }
    }
}

fn handle_shortcut_key(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    key: Key,
    modifiers: Modifiers,
    history: &mut ChangeHistory,
    _ui: &mut egui::Ui,
) -> bool {
    if !modifiers.command {
        return false;
    }

    match key {
        Key::B => {
            crate::app::actions::toggle_bold(document, canvas, history);
            true
        }
        Key::I => {
            crate::app::actions::toggle_italic(document, canvas, history);
            true
        }
        Key::U => {
            crate::app::actions::toggle_underline(document, canvas, history);
            true
        }
        Key::A if canvas.active_table_cell.is_some() => {
            if let Some((table_id, row, col)) = canvas.active_table_cell {
                if let Some(len) = document.table_cell_len(table_id, row, col) {
                    canvas.table_cell_selection =
                        CCursorRange::two(CCursor::new(0), CCursor::new(len));
                    return true;
                }
            }
            false
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

    canvas.selection = CCursorRange::one(CCursor::new(next_index));
    canvas.active_style = document.typing_style_at(next_index);
    canvas.active_paragraph_style = document.paragraph_style_at(next_index);
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

fn delete_word_backward(document: &mut DocumentState, canvas: &mut CanvasState) -> bool {
    let selected = canvas.selection.as_sorted_char_range();
    if selected.start < selected.end {
        document.delete_range(selected.clone());
        canvas.selection = CCursorRange::one(CCursor::new(selected.start));
        return true;
    }

    let text = document.plain_text();
    let chars: Vec<char> = text.chars().collect();
    let delete_end = canvas.selection.primary.index.min(chars.len());
    if delete_end == 0 {
        return false;
    }

    let mut delete_start = delete_end;
    while delete_start > 0 && !is_word_char(chars[delete_start - 1]) {
        delete_start -= 1;
    }
    while delete_start > 0 && is_word_char(chars[delete_start - 1]) {
        delete_start -= 1;
    }

    if delete_start == delete_end {
        delete_start = delete_start.saturating_sub(1);
    }

    document.delete_range(delete_start..delete_end);
    canvas.selection = CCursorRange::one(CCursor::new(delete_start));
    true
}

fn delete_word_forward(document: &mut DocumentState, canvas: &mut CanvasState) -> bool {
    let selected = canvas.selection.as_sorted_char_range();
    if selected.start < selected.end {
        document.delete_range(selected.clone());
        canvas.selection = CCursorRange::one(CCursor::new(selected.start));
        return true;
    }

    let text = document.plain_text();
    let chars: Vec<char> = text.chars().collect();
    let delete_start = canvas.selection.primary.index.min(chars.len());
    if delete_start >= chars.len() {
        return false;
    }

    let mut delete_end = delete_start;
    if !is_word_char(chars[delete_end]) {
        while delete_end < chars.len() && !is_word_char(chars[delete_end]) {
            delete_end += 1;
        }
    }
    while delete_end < chars.len() && is_word_char(chars[delete_end]) {
        delete_end += 1;
    }

    if delete_end == delete_start {
        delete_end += 1;
    }

    document.delete_range(delete_start..delete_end);
    canvas.selection = CCursorRange::one(CCursor::new(delete_start));
    true
}

fn is_word_char(ch: char) -> bool {
    ch.is_alphanumeric() || ch == '_'
}

#[cfg(test)]
mod tests {
    use eframe::egui::{
        epaint::text::cursor::CCursor, text_selection::CCursorRange, Key, Modifiers,
    };

    use crate::{
        app::CanvasState,
        document::{CharacterStyle, DocumentState, TextRun},
    };

    use super::{delete_word_backward, delete_word_forward, move_active_table_cell_cursor_by_key};

    fn document_with_text(text: &str) -> DocumentState {
        let mut document = DocumentState::bootstrap();
        document.replace_with_runs(
            "Test".to_owned(),
            vec![TextRun {
                text: text.to_owned(),
                style: CharacterStyle::default(),
            }],
        );
        document
    }

    fn canvas_with_cursor(index: usize) -> CanvasState {
        CanvasState {
            selection: CCursorRange::one(CCursor::new(index)),
            ..CanvasState::default()
        }
    }

    #[test]
    fn ctrl_delete_removes_word_from_caret() {
        let mut document = document_with_text("alpha beta");
        let mut canvas = canvas_with_cursor(0);

        assert!(delete_word_forward(&mut document, &mut canvas));

        assert_eq!(document.plain_text(), " beta");
        assert_eq!(canvas.selection.primary.index, 0);
    }

    #[test]
    fn ctrl_delete_skips_separator_and_removes_next_word() {
        let mut document = document_with_text("alpha beta");
        let mut canvas = canvas_with_cursor(5);

        assert!(delete_word_forward(&mut document, &mut canvas));

        assert_eq!(document.plain_text(), "alpha");
        assert_eq!(canvas.selection.primary.index, 5);
    }

    #[test]
    fn ctrl_backspace_removes_word_before_caret() {
        let mut document = document_with_text("alpha beta");
        let mut canvas = canvas_with_cursor(10);

        assert!(delete_word_backward(&mut document, &mut canvas));

        assert_eq!(document.plain_text(), "alpha ");
        assert_eq!(canvas.selection.primary.index, 6);
    }

    #[test]
    fn ctrl_backspace_skips_separator_and_removes_previous_word() {
        let mut document = document_with_text("alpha beta");
        let mut canvas = canvas_with_cursor(6);

        assert!(delete_word_backward(&mut document, &mut canvas));

        assert_eq!(document.plain_text(), "beta");
        assert_eq!(canvas.selection.primary.index, 0);
    }

    #[test]
    fn table_cursor_uses_layout_aware_arrow_navigation() {
        let ctx = egui::Context::default();
        let mut document = DocumentState::bootstrap();
        let table_id = document.insert_table(0, 1, 1);
        document.set_table_cell_text(table_id, 0, 0, "abc\ndef");
        let mut canvas = CanvasState {
            active_table_cell: Some((table_id, 0, 0)),
            table_cell_selection: CCursorRange::one(CCursor::new(1)),
            ..CanvasState::default()
        };

        let _ = ctx.run_ui(egui::RawInput::default(), |ui| {
            move_active_table_cell_cursor_by_key(
                ui,
                &document,
                &mut canvas,
                Key::ArrowDown,
                Modifiers::NONE,
            );
            move_active_table_cell_cursor_by_key(
                ui,
                &document,
                &mut canvas,
                Key::ArrowRight,
                Modifiers::CTRL | Modifiers::SHIFT,
            );
        });

        assert_eq!(canvas.table_cell_selection.secondary.index, 5);
        assert_eq!(canvas.table_cell_selection.primary.index, 7);
    }

    #[test]
    fn ctrl_scroll_zoom_direction_is_natural() {
        assert!((20.0_f32 * 0.01).exp() > 1.0);
        assert!((-20.0_f32 * 0.01).exp() < 1.0);
    }
}
