use std::ops::Range;
use eframe::egui::{
    self, epaint::text::cursor::CCursor, text_selection::visuals::paint_text_cursor,
    text_selection::CCursorRange, Align2, Color32, FontFamily, FontId, Id, Rect,
    Stroke,
};

use crate::app::{ActiveHeaderFooter, CanvasState, ChangeHistory};
use crate::document::{CharacterStyle, DocumentState, HeaderFooterKind, TextRun};
use crate::layout::document_points_to_screen_points;
use super::super::page_layout::PageLayout;
use super::rendering::{
    paint_tab_aligned_margin_runs, page_header_rect, page_footer_rect,
    runs_plain_text, runs_total_chars, slice_run_text_chars, HeaderSegment,
    measure_text_width, header_footer_line_height, split_runs_for_header_tabs,
    measure_runs_width,
};

pub(crate) fn paint_active_header_footer_editor(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    page_layout: &PageLayout,
    response: &egui::Response,
    editor_id: Id,
) -> bool {
    let Some(active) = canvas.active_header_footer else {
        return false;
    };
    let Some(page) = page_layout.pages.get(active.page_number.saturating_sub(1)) else {
        canvas.active_header_footer = None;
        return false;
    };

    let margin_rect = match active.kind {
        HeaderFooterKind::Header => {
            page_header_rect(page.page_rect, document, canvas, active.section_id)
        }
        HeaderFooterKind::Footer => {
            page_footer_rect(page.page_rect, document, canvas, active.section_id)
        }
    };
    let horizontal_margin = document_points_to_screen_points(
        document
            .section_by_id(active.section_id)
            .map(|section| section.page_setup.margins.left_points)
            .unwrap_or_else(|| document.default_page_setup().margins.left_points)
            .max(18.0),
        canvas.zoom,
    );
    let editor_height = document_points_to_screen_points(20.0, canvas.zoom).clamp(18.0, 28.0);
    let editor_rect = Rect::from_center_size(
        margin_rect.center(),
        egui::vec2(
            (margin_rect.width() - horizontal_margin * 2.0).max(80.0),
            editor_height,
        ),
    );

    let guide_y = match active.kind {
        HeaderFooterKind::Header => margin_rect.bottom(),
        HeaderFooterKind::Footer => margin_rect.top(),
    };
    ui.painter().line_segment(
        [
            egui::pos2(editor_rect.left(), guide_y),
            egui::pos2(editor_rect.right(), guide_y),
        ],
        Stroke::new(1.0, Color32::from_rgba_unmultiplied(120, 130, 145, 110)),
    );

    let active_runs = active_header_footer_runs(document, active).to_vec();

    if let Some(pointer_pos) = response.interact_pointer_pos() {
        let press_origin = ui.input(|i| i.pointer.press_origin());
        let interact_header = editor_rect.expand(20.0).contains(pointer_pos)
            || press_origin.is_some_and(|origin| editor_rect.expand(20.0).contains(origin));

        if interact_header {
            if response.clicked() || response.drag_started() {
                response.request_focus();
            }
            let cursor = CCursor::new(header_footer_cursor_from_pos(
                ui.painter(),
                &active_runs,
                canvas.zoom,
                editor_rect,
                pointer_pos.x,
            ));
            if response.drag_started() {
                if ui.input(|input| input.modifiers.shift) {
                    canvas.active_header_footer_selection.primary = cursor;
                } else {
                    canvas.active_header_footer_selection = CCursorRange::one(cursor);
                }
                canvas.active_header_footer_cursor = cursor.index;
                canvas.active_header_footer_selection.h_pos = None;
                canvas.last_interaction_time = ui.input(|input| input.time);
            } else if response.dragged() {
                canvas.active_header_footer_selection.primary = cursor;
                canvas.active_header_footer_cursor = cursor.index;
                canvas.active_header_footer_selection.h_pos = None;
                canvas.last_interaction_time = ui.input(|input| input.time);
            } else if response.clicked() {
                if ui.input(|input| input.modifiers.shift) {
                    canvas.active_header_footer_selection.primary = cursor;
                } else {
                    canvas.active_header_footer_selection = CCursorRange::one(cursor);
                }
                canvas.active_header_footer_cursor = cursor.index;
                canvas.active_header_footer_selection.h_pos = None;
                canvas.last_interaction_time = ui.input(|input| input.time);
            }
        }
    }

    let has_focus = ui.memory(|memory| memory.has_focus(editor_id));

    let mut edited_runs = active_runs;
    let total_chars = runs_total_chars(&edited_runs);
    canvas.active_header_footer_cursor = canvas.active_header_footer_cursor.min(total_chars);
    canvas.active_header_footer_selection.primary.index = canvas
        .active_header_footer_selection
        .primary
        .index
        .min(total_chars);
    canvas.active_header_footer_selection.secondary.index = canvas
        .active_header_footer_selection
        .secondary
        .index
        .min(total_chars);
    let before_runs = edited_runs.clone();
    let changed = if has_focus {
        handle_header_footer_keyboard_input(ui, &mut edited_runs, canvas, history, document)
    } else {
        false
    };
    if runs_plain_text(&edited_runs).trim().is_empty() {
        let hint = match active.kind {
            HeaderFooterKind::Header => "Header",
            HeaderFooterKind::Footer => "Footer",
        };
        let inherited = document
            .resolve_header_footer_slot(active.section_id, active.kind, active.variant)
            .inherited;
        let section_label = format!(
            "{hint} - Section {}{}",
            active.section_id,
            if inherited { " (Same as Previous)" } else { "" }
        );
        ui.painter().text(
            editor_rect.left_top(),
            Align2::LEFT_TOP,
            section_label,
            FontId::new(editor_height * 0.72, FontFamily::Proportional),
            Color32::from_rgba_premultiplied(96, 104, 118, 140),
        );
    } else {
        let selection = has_focus
            .then(|| canvas.active_header_footer_selection.as_sorted_char_range())
            .filter(|range| range.start < range.end);
        paint_tab_aligned_margin_runs(
            ui.painter(),
            &edited_runs,
            canvas.zoom,
            Color32::from_rgb(36, 39, 46),
            editor_rect,
            selection,
        );
    }

    if has_focus {
        let cursor_pos = header_footer_cursor_pos(
            ui.painter(),
            &edited_runs,
            canvas.zoom,
            editor_rect,
            canvas.active_header_footer_cursor,
        );
        let time = ui.input(|input| input.time) - canvas.last_interaction_time;
        paint_text_cursor(
            ui,
            ui.painter(),
            Rect::from_min_size(cursor_pos, egui::vec2(1.5, editor_height * 0.85)),
            time,
        );
    }

    if changed && edited_runs != before_runs {
        normalize_header_footer_runs(&mut edited_runs);
        *active_header_footer_runs_mut(document, active) = edited_runs;
        document.sync_compat_from_first_section();
        canvas.active_style = header_footer_style_at(
            active_header_footer_runs(document, active),
            canvas.active_header_footer_selection.primary.index,
        );
        true
    } else {
        false
    }
}

fn handle_header_footer_keyboard_input(
    ui: &egui::Ui,
    runs: &mut Vec<TextRun>,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    document: &DocumentState,
) -> bool {
    let mut changed = false;
    let events = ui.input(|input| input.events.clone());
    for event in events {
        match event {
            egui::Event::Copy => {
                copy_header_footer_selection(ui, runs, canvas);
            }
            egui::Event::Cut => {
                if copy_header_footer_selection(ui, runs, canvas) {
                    history.checkpoint(document, ui.input(|input| input.time));
                    let selected = canvas.active_header_footer_selection.as_sorted_char_range();
                    delete_header_footer_range(runs, selected.clone());
                    canvas.active_header_footer_selection =
                        CCursorRange::one(CCursor::new(selected.start));
                    canvas.active_header_footer_cursor = selected.start;
                    changed = true;
                }
            }
            egui::Event::Text(inserted) if !inserted.is_empty() => {
                history.checkpoint_coalesced(document, ui.input(|input| input.time));
                let selected = canvas.active_header_footer_selection.as_sorted_char_range();
                let next = replace_header_footer_range_with_text(
                    runs,
                    selected,
                    &inserted,
                    canvas.active_style,
                );
                canvas.active_header_footer_selection = CCursorRange::one(CCursor::new(next));
                canvas.active_header_footer_cursor = next;
                changed = true;
            }
            egui::Event::Paste(pasted) => {
                history.checkpoint(document, ui.input(|input| input.time));
                let selected = canvas.active_header_footer_selection.as_sorted_char_range();
                let next = replace_header_footer_range_with_text(
                    runs,
                    selected,
                    &pasted,
                    canvas.active_style,
                );
                canvas.active_header_footer_selection = CCursorRange::one(CCursor::new(next));
                canvas.active_header_footer_cursor = next;
                changed = true;
            }
            egui::Event::Key {
                key: egui::Key::Tab,
                pressed: true,
                modifiers,
                ..
            } => {
                history.checkpoint(document, ui.input(|input| input.time));
                let next_cursor = if modifiers.shift {
                    remove_previous_header_footer_tab(
                        runs,
                        canvas.active_header_footer_selection.as_sorted_char_range(),
                    )
                } else {
                    insert_header_footer_tab(
                        runs,
                        canvas.active_header_footer_selection.as_sorted_char_range(),
                        canvas.active_style,
                    )
                };
                if let Some(next_cursor) = next_cursor {
                    canvas.active_header_footer_selection =
                        CCursorRange::one(CCursor::new(next_cursor));
                    canvas.active_header_footer_cursor = next_cursor;
                    changed = true;
                }
            }
            egui::Event::Key {
                key: egui::Key::Backspace,
                pressed: true,
                ..
            } => {
                let selected = canvas.active_header_footer_selection.as_sorted_char_range();
                if selected.start < selected.end {
                    history.checkpoint_coalesced(document, ui.input(|input| input.time));
                    delete_header_footer_range(runs, selected.clone());
                    canvas.active_header_footer_selection =
                        CCursorRange::one(CCursor::new(selected.start));
                    canvas.active_header_footer_cursor = selected.start;
                    changed = true;
                } else if canvas.active_header_footer_cursor > 0 {
                    history.checkpoint_coalesced(document, ui.input(|input| input.time));
                    let start = canvas.active_header_footer_cursor - 1;
                    delete_header_footer_range(runs, start..canvas.active_header_footer_cursor);
                    canvas.active_header_footer_selection = CCursorRange::one(CCursor::new(start));
                    canvas.active_header_footer_cursor = start;
                    changed = true;
                }
            }
            egui::Event::Key {
                key: egui::Key::Delete,
                pressed: true,
                ..
            } => {
                let selected = canvas.active_header_footer_selection.as_sorted_char_range();
                if selected.start < selected.end {
                    history.checkpoint_coalesced(document, ui.input(|input| input.time));
                    delete_header_footer_range(runs, selected.clone());
                    canvas.active_header_footer_selection =
                        CCursorRange::one(CCursor::new(selected.start));
                    canvas.active_header_footer_cursor = selected.start;
                    changed = true;
                } else {
                    let total_chars = runs_total_chars(runs);
                    if canvas.active_header_footer_cursor < total_chars {
                        history.checkpoint_coalesced(document, ui.input(|input| input.time));
                        delete_header_footer_range(
                            runs,
                            canvas.active_header_footer_cursor
                                ..canvas.active_header_footer_cursor + 1,
                        );
                        changed = true;
                    }
                }
            }
            egui::Event::Key {
                key,
                pressed: true,
                modifiers,
                ..
            } => {
                if handle_header_footer_shortcut_key(
                    ui, document, canvas, history, runs, key, modifiers,
                ) {
                    changed = true;
                    continue;
                }
                match key {
                    egui::Key::ArrowLeft => {
                        let next = canvas.active_header_footer_cursor.saturating_sub(1);
                        set_header_footer_cursor(canvas, next, modifiers.shift);
                    }
                    egui::Key::ArrowRight => {
                        let next =
                            (canvas.active_header_footer_cursor + 1).min(runs_total_chars(runs));
                        set_header_footer_cursor(canvas, next, modifiers.shift);
                    }
                    egui::Key::Home => {
                        set_header_footer_cursor(canvas, 0, modifiers.shift);
                    }
                    egui::Key::End => {
                        set_header_footer_cursor(canvas, runs_total_chars(runs), modifiers.shift);
                    }
                    _ => {}
                }
            }
            _ => {}
        }
    }
    changed
}

fn handle_header_footer_shortcut_key(
    ui: &egui::Ui,
    document: &DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    runs: &mut Vec<TextRun>,
    key: egui::Key,
    modifiers: egui::Modifiers,
) -> bool {
    if !modifiers.command {
        return false;
    }

    match key {
        egui::Key::A => {
            let total_chars = runs_total_chars(runs);
            canvas.active_header_footer_selection =
                CCursorRange::two(CCursor::new(0), CCursor::new(total_chars));
            canvas.active_header_footer_cursor = total_chars;
            false
        }
        egui::Key::B => {
            history.checkpoint(document, ui.input(|input| input.time));
            let next = !canvas.active_style.bold;
            apply_header_footer_style_change(runs, canvas, |style| style.bold = next)
        }
        egui::Key::I => {
            history.checkpoint(document, ui.input(|input| input.time));
            let next = !canvas.active_style.italic;
            apply_header_footer_style_change(runs, canvas, |style| style.italic = next)
        }
        egui::Key::U => {
            history.checkpoint(document, ui.input(|input| input.time));
            let next = !canvas.active_style.underline;
            apply_header_footer_style_change(runs, canvas, |style| style.underline = next)
        }
        _ => false,
    }
}

fn apply_header_footer_style_change(
    runs: &mut Vec<TextRun>,
    canvas: &mut CanvasState,
    mutate: impl Fn(&mut CharacterStyle) + Copy,
) -> bool {
    let selected = canvas.active_header_footer_selection.as_sorted_char_range();
    let changed = if selected.start < selected.end {
        apply_style_to_header_footer_range(runs, selected, mutate);
        true
    } else {
        false
    };
    mutate(&mut canvas.active_style);
    changed
}

fn set_header_footer_cursor(canvas: &mut CanvasState, cursor: usize, extend_selection: bool) {
    canvas.active_header_footer_cursor = cursor;
    if extend_selection {
        canvas.active_header_footer_selection.primary = CCursor::new(cursor);
    } else {
        canvas.active_header_footer_selection = CCursorRange::one(CCursor::new(cursor));
    }
}

fn copy_header_footer_selection(ui: &egui::Ui, runs: &[TextRun], canvas: &CanvasState) -> bool {
    let selected = canvas.active_header_footer_selection.as_sorted_char_range();
    if selected.start >= selected.end {
        return false;
    }
    ui.copy_text(selected_header_footer_text(runs, selected));
    true
}

fn selected_header_footer_text(runs: &[TextRun], range: Range<usize>) -> String {
    runs_plain_text(runs)
        .chars()
        .skip(range.start)
        .take(range.end.saturating_sub(range.start))
        .collect()
}

fn insert_header_footer_tab(
    runs: &mut Vec<TextRun>,
    range: Range<usize>,
    style: CharacterStyle,
) -> Option<usize> {
    let plain = runs_plain_text(runs);
    let tab_count_before_cursor = plain
        .chars()
        .take(range.start)
        .filter(|ch| *ch == '\t')
        .count();
    if tab_count_before_cursor >= 2 {
        return None;
    }

    Some(replace_header_footer_range_with_text(
        runs, range, "\t", style,
    ))
}

fn remove_previous_header_footer_tab(
    runs: &mut Vec<TextRun>,
    range: Range<usize>,
) -> Option<usize> {
    if range.start < range.end {
        delete_header_footer_range(runs, range.clone());
        return Some(range.start);
    }

    let plain = runs_plain_text(runs);
    let previous_tab = plain
        .chars()
        .take(range.start)
        .enumerate()
        .filter_map(|(index, ch)| (ch == '\t').then_some(index))
        .last()?;
    delete_header_footer_range(runs, previous_tab..previous_tab + 1);
    Some(previous_tab)
}

fn char_to_byte_index_for_header_footer(text: &str, char_index: usize) -> usize {
    text.char_indices()
        .nth(char_index)
        .map(|(byte_index, _)| byte_index)
        .unwrap_or(text.len())
}

fn header_footer_style_at(runs: &[TextRun], char_index: usize) -> CharacterStyle {
    let total = runs_total_chars(runs);
    let target = char_index.min(total);
    let mut offset = 0usize;
    for run in runs {
        let len = run.text.chars().count();
        if target < offset + len {
            return run.style;
        }
        offset += len;
    }
    runs.last().map(|run| run.style).unwrap_or_default()
}

fn normalize_header_footer_runs(runs: &mut Vec<TextRun>) {
    runs.retain(|run| !run.text.is_empty());
    let mut normalized: Vec<TextRun> = Vec::with_capacity(runs.len().max(1));
    for run in runs.drain(..) {
        if let Some(last) = normalized.last_mut() {
            if last.style == run.style {
                last.text.push_str(&run.text);
                continue;
            }
        }
        normalized.push(run);
    }
    if normalized.is_empty() {
        normalized.push(TextRun {
            text: String::new(),
            style: CharacterStyle::default(),
        });
    }
    *runs = normalized;
}

fn split_header_footer_runs_at(runs: &mut Vec<TextRun>, char_index: usize) {
    if char_index == 0 || char_index >= runs_total_chars(runs) {
        return;
    }

    let mut offset = 0usize;
    for idx in 0..runs.len() {
        let len = runs[idx].text.chars().count();
        if char_index > offset && char_index < offset + len {
            let local = char_index - offset;
            let byte_index = char_to_byte_index_for_header_footer(&runs[idx].text, local);
            let right = runs[idx].text.split_off(byte_index);
            let style = runs[idx].style;
            runs.insert(idx + 1, TextRun { text: right, style });
            break;
        }
        offset += len;
    }
}

fn replace_header_footer_range_with_text(
    runs: &mut Vec<TextRun>,
    range: Range<usize>,
    text: &str,
    style: CharacterStyle,
) -> usize {
    let start = range.start.min(runs_total_chars(runs));
    let end = range.end.min(runs_total_chars(runs));
    delete_header_footer_range(runs, start..end);
    insert_header_footer_text(runs, start, text, style);
    start + text.chars().count()
}

fn insert_header_footer_text(
    runs: &mut Vec<TextRun>,
    char_index: usize,
    text: &str,
    style: CharacterStyle,
) {
    if text.is_empty() {
        return;
    }

    let insertion_index = char_index.min(runs_total_chars(runs));
    split_header_footer_runs_at(runs, insertion_index);

    let mut offset = 0usize;
    let mut target = runs.len();
    for (idx, run) in runs.iter().enumerate() {
        if offset == insertion_index {
            target = idx;
            break;
        }
        offset += run.text.chars().count();
    }

    runs.insert(
        target,
        TextRun {
            text: text.to_owned(),
            style,
        },
    );
    normalize_header_footer_runs(runs);
}

fn delete_header_footer_range(runs: &mut Vec<TextRun>, range: Range<usize>) {
    if range.start >= range.end {
        return;
    }

    let start = range.start.min(runs_total_chars(runs));
    let end = range.end.min(runs_total_chars(runs));
    split_header_footer_runs_at(runs, start);
    split_header_footer_runs_at(runs, end);

    let mut next_runs = Vec::new();
    let mut offset = 0usize;
    for run in runs.drain(..) {
        let len = run.text.chars().count();
        if offset >= start && offset + len <= end {
            // fully inside range, delete
        } else {
            next_runs.push(run);
        }
        offset += len;
    }
    *runs = next_runs;
    normalize_header_footer_runs(runs);
}

fn apply_style_to_header_footer_range(
    runs: &mut Vec<TextRun>,
    range: Range<usize>,
    mutate: impl Fn(&mut CharacterStyle),
) {
    if range.start >= range.end {
        return;
    }

    let start = range.start.min(runs_total_chars(runs));
    let end = range.end.min(runs_total_chars(runs));
    split_header_footer_runs_at(runs, start);
    split_header_footer_runs_at(runs, end);

    let mut offset = 0usize;
    for run in runs {
        let len = run.text.chars().count();
        if offset >= start && offset + len <= end {
            mutate(&mut run.style);
        }
        offset += len;
    }
}

fn header_footer_cursor_from_pos(
    painter: &egui::Painter,
    runs: &[TextRun],
    zoom: f32,
    rect: Rect,
    x: f32,
) -> usize {
    let total_chars = runs_total_chars(runs);
    (0..=total_chars)
        .min_by(|left, right| {
            let left_x = header_footer_cursor_pos(painter, runs, zoom, rect, *left).x;
            let right_x = header_footer_cursor_pos(painter, runs, zoom, rect, *right).x;
            (left_x - x)
                .abs()
                .partial_cmp(&(right_x - x).abs())
                .unwrap_or(std::cmp::Ordering::Equal)
        })
        .unwrap_or(total_chars)
}

fn header_footer_cursor_pos(
    painter: &egui::Painter,
    runs: &[TextRun],
    zoom: f32,
    rect: Rect,
    cursor: usize,
) -> egui::Pos2 {
    let segments = split_runs_for_header_tabs(runs);
    let mut segment_cursor = cursor;
    let mut slot = 0usize;

    for s in 0..3 {
        let Some(segment) = segments.get(s) else {
            continue;
        };
        let segment_chars = segment
            .runs
            .iter()
            .map(|r| r.range.as_ref().map_or(0, |rng| rng.end - rng.start))
            .sum::<usize>();
        let end_char = segment.end;
        let start_char = segment.end - segment_chars;

        if cursor >= start_char && cursor <= end_char {
            slot = s;
            segment_cursor = cursor;
            break;
        }
        if cursor < start_char && s == 0 {
            slot = 0;
            segment_cursor = cursor;
            break;
        }
        if cursor > end_char && s == segments.len().saturating_sub(1) {
            slot = s;
            segment_cursor = cursor;
            break;
        }
    }

    let Some(segment) = segments.get(slot) else {
        return rect.left_top();
    };

    let prefix_width = measure_segment_prefix_width(painter, segment, segment_cursor, zoom);
    let segment_width = measure_runs_width(painter, &segment.runs, zoom);
    let start_x = match slot {
        0 => rect.left(),
        1 => rect.center().x - segment_width * 0.5,
        _ => rect.right() - segment_width,
    };

    egui::pos2(start_x + prefix_width, rect.top() + (rect.height() - header_footer_line_height(CharacterStyle::default(), zoom)) * 0.5)
}

fn measure_segment_prefix_width(
    painter: &egui::Painter,
    segment: &HeaderSegment,
    cursor: usize,
    zoom: f32,
) -> f32 {
    let mut width = 0.0;
    for piece in &segment.runs {
        let Some(range) = piece.range.clone() else {
            continue;
        };
        if cursor >= range.end {
            width += measure_text_width(painter, &piece.text, piece.style, zoom);
        } else if cursor > range.start {
            let prefix = slice_run_text_chars(&piece.text, 0..cursor - range.start);
            width += measure_text_width(painter, &prefix, piece.style, zoom);
            break;
        } else {
            break;
        }
    }
    width
}

fn active_header_footer_runs(document: &DocumentState, active: ActiveHeaderFooter) -> &[TextRun] {
    document
        .resolve_header_footer_slot(active.section_id, active.kind, active.variant)
        .story
        .runs
        .as_slice()
}

fn active_header_footer_runs_mut(
    document: &mut DocumentState,
    active: ActiveHeaderFooter,
) -> &mut Vec<TextRun> {
    &mut document
        .header_footer_story_mut_materialized(active.section_id, active.kind, active.variant)
        .expect("active header/footer section should exist")
        .runs
}
