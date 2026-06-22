use std::ops::Range;

use eframe::egui::{self, epaint::text::cursor::CCursor, text_selection::CCursorRange};

use crate::{
    app::{CanvasState, ChangeHistory},
    document::DocumentState,
};

#[derive(Default)]
pub(crate) struct FindReplaceState {
    pub visible: bool,
    pub find_text: String,
    pub replace_text: String,
    pub match_case: bool,
    pub whole_word: bool,
    pub last_match: Option<Range<usize>>,
    pub message: String,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(crate) struct FindOptions {
    pub match_case: bool,
    pub whole_word: bool,
}

impl FindReplaceState {
    pub(crate) fn options(&self) -> FindOptions {
        FindOptions {
            match_case: self.match_case,
            whole_word: self.whole_word,
        }
    }
}

pub(crate) fn paint_find_replace_window(
    ctx: &egui::Context,
    state: &mut FindReplaceState,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    status_message: &mut String,
) -> bool {
    if !state.visible {
        return false;
    }

    let mut changed = false;
    let mut visible = state.visible;
    egui::Window::new("Find and Replace")
        .open(&mut visible)
        .collapsible(false)
        .resizable(false)
        .default_width(360.0)
        .show(ctx, |ui| {
            let find_response = ui.horizontal(|ui| {
                ui.label("Find");
                ui.text_edit_singleline(&mut state.find_text)
            });
            if find_response.inner.changed() {
                state.last_match = None;
                state.message.clear();
            }

            ui.horizontal(|ui| {
                ui.label("Replace");
                ui.text_edit_singleline(&mut state.replace_text);
            });

            ui.horizontal(|ui| {
                if ui.checkbox(&mut state.match_case, "Match case").changed() {
                    state.last_match = None;
                }
                if ui.checkbox(&mut state.whole_word, "Whole word").changed() {
                    state.last_match = None;
                }
            });

            ui.label(
                egui::RichText::new("Searches document body text")
                    .size(11.0)
                    .color(ui.visuals().weak_text_color()),
            );

            let has_query = !state.find_text.is_empty();
            if !has_query && state.message.is_empty() {
                state.message = "Enter text to find".to_owned();
            }

            ui.horizontal(|ui| {
                if ui
                    .add_enabled(has_query, egui::Button::new("Previous"))
                    .clicked()
                {
                    let start = canvas.selection.as_sorted_char_range().start;
                    select_previous_match(document, canvas, state, start);
                    *status_message = state.message.clone();
                }
                if ui
                    .add_enabled(has_query, egui::Button::new("Next"))
                    .clicked()
                {
                    let start = canvas.selection.as_sorted_char_range().end;
                    select_next_match(document, canvas, state, start);
                    *status_message = state.message.clone();
                }
                if ui
                    .add_enabled(has_query, egui::Button::new("Replace"))
                    .clicked()
                {
                    if replace_current_match(document, canvas, history, state) {
                        *status_message = "Replaced match".to_owned();
                        changed = true;
                    } else {
                        *status_message = state.message.clone();
                    }
                }
                if ui
                    .add_enabled(has_query, egui::Button::new("Replace All"))
                    .clicked()
                {
                    let count = replace_all(document, canvas, history, state);
                    changed |= count > 0;
                    *status_message = match count {
                        0 => "No matches".to_owned(),
                        1 => "Replaced 1 match".to_owned(),
                        _ => format!("Replaced {count} matches"),
                    };
                }
            });

            if !state.message.is_empty() {
                ui.label(
                    egui::RichText::new(&state.message)
                        .size(11.0)
                        .color(ui.visuals().weak_text_color()),
                );
            }
        });
    state.visible = visible;
    changed
}

pub(crate) fn find_next_match(
    document_text: &str,
    query: &str,
    start_char: usize,
    options: FindOptions,
) -> Option<Range<usize>> {
    if query.is_empty() {
        return None;
    }
    let total_chars = document_text.chars().count();
    let start = start_char.min(total_chars);
    find_in_range(document_text, query, start..total_chars, options)
        .or_else(|| find_in_range(document_text, query, 0..start, options))
}

pub(crate) fn find_previous_match(
    document_text: &str,
    query: &str,
    start_char: usize,
    options: FindOptions,
) -> Option<Range<usize>> {
    if query.is_empty() {
        return None;
    }
    let total_chars = document_text.chars().count();
    let start = start_char.min(total_chars);
    all_matches(document_text, query, options)
        .into_iter()
        .filter(|range| range.start < start)
        .last()
        .or_else(|| {
            all_matches(document_text, query, options)
                .into_iter()
                .last()
        })
}

pub(crate) fn replace_current_match(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    state: &mut FindReplaceState,
) -> bool {
    if state.find_text.is_empty() {
        state.message = "Enter text to find".to_owned();
        return false;
    }

    let selected = canvas.selection.as_sorted_char_range();
    let text = document.plain_text();
    let range = if selected.start < selected.end
        && range_matches(&text, &state.find_text, selected.clone(), state.options())
    {
        selected
    } else if let Some(found) = find_next_match(
        &text,
        &state.find_text,
        canvas.selection.primary.index,
        state.options(),
    ) {
        found
    } else {
        state.message = "No matches".to_owned();
        state.last_match = None;
        return false;
    };

    history.checkpoint(document, f64::NAN);
    replace_range(document, canvas, range.clone(), &state.replace_text);
    let end = range.start + state.replace_text.chars().count();
    state.last_match = Some(range.start..end);
    state.message = "Replaced match".to_owned();
    true
}

pub(crate) fn replace_all(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    state: &mut FindReplaceState,
) -> usize {
    if state.find_text.is_empty() {
        state.message = "Enter text to find".to_owned();
        return 0;
    }

    let matches = all_matches(&document.plain_text(), &state.find_text, state.options());
    if matches.is_empty() {
        state.message = "No matches".to_owned();
        state.last_match = None;
        return 0;
    }

    history.checkpoint(document, f64::NAN);
    for range in matches.iter().rev() {
        replace_range(document, canvas, range.clone(), &state.replace_text);
    }

    let first_start = matches[0].start;
    let first_end = first_start + state.replace_text.chars().count();
    canvas.selection = CCursorRange::two(CCursor::new(first_start), CCursor::new(first_end));
    state.last_match = Some(first_start..first_end);
    state.message = match matches.len() {
        1 => "Replaced 1 match".to_owned(),
        count => format!("Replaced {count} matches"),
    };
    matches.len()
}

fn select_next_match(
    document: &DocumentState,
    canvas: &mut CanvasState,
    state: &mut FindReplaceState,
    start_char: usize,
) {
    let found = find_next_match(
        &document.plain_text(),
        &state.find_text,
        start_char,
        state.options(),
    );
    select_match(canvas, state, found);
}

fn select_previous_match(
    document: &DocumentState,
    canvas: &mut CanvasState,
    state: &mut FindReplaceState,
    start_char: usize,
) {
    let found = find_previous_match(
        &document.plain_text(),
        &state.find_text,
        start_char,
        state.options(),
    );
    select_match(canvas, state, found);
}

fn select_match(
    canvas: &mut CanvasState,
    state: &mut FindReplaceState,
    found: Option<Range<usize>>,
) {
    if let Some(range) = found {
        canvas.active_header_footer = None;
        canvas.active_table_cell = None;
        canvas.selected_image_id = None;
        canvas.selection = CCursorRange::two(CCursor::new(range.start), CCursor::new(range.end));
        state.last_match = Some(range);
        state.message = "Match selected".to_owned();
    } else {
        state.last_match = None;
        state.message = "No matches".to_owned();
    }
}

fn replace_range(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    range: Range<usize>,
    replacement: &str,
) {
    let style = document.style_at(range.start);
    document.delete_range(range.clone());
    document.insert_text(range.start, replacement, style);
    let end = range.start + replacement.chars().count();
    canvas.selection = CCursorRange::two(CCursor::new(range.start), CCursor::new(end));
    canvas.active_style = document.typing_style_at(end);
    canvas.active_paragraph_style = document.paragraph_style_at(end);
}

fn find_in_range(
    document_text: &str,
    query: &str,
    search_range: Range<usize>,
    options: FindOptions,
) -> Option<Range<usize>> {
    let query_len = query.chars().count();
    let total_chars = document_text.chars().count();
    if query_len == 0 || search_range.start >= search_range.end || query_len > total_chars {
        return None;
    }

    let max_start = search_range.end.saturating_sub(query_len);
    (search_range.start..=max_start)
        .map(|start| start..start + query_len)
        .find(|range| range_matches(document_text, query, range.clone(), options))
}

fn all_matches(document_text: &str, query: &str, options: FindOptions) -> Vec<Range<usize>> {
    let mut matches = Vec::new();
    let mut start = 0usize;
    while let Some(range) = find_next_match(document_text, query, start, options) {
        if range.start < start {
            break;
        }
        start = range.end.max(range.start + 1);
        matches.push(range);
    }
    matches
}

fn range_matches(
    document_text: &str,
    query: &str,
    range: Range<usize>,
    options: FindOptions,
) -> bool {
    if options.whole_word && !is_whole_word_match(document_text, range.clone()) {
        return false;
    }
    let candidate = slice_chars(document_text, range);
    if options.match_case {
        candidate == query
    } else {
        candidate.to_lowercase() == query.to_lowercase()
    }
}

fn is_whole_word_match(document_text: &str, range: Range<usize>) -> bool {
    let chars: Vec<char> = document_text.chars().collect();
    let before = range.start.checked_sub(1).and_then(|idx| chars.get(idx));
    let after = chars.get(range.end);
    !before.is_some_and(|ch| is_word_char(*ch)) && !after.is_some_and(|ch| is_word_char(*ch))
}

fn is_word_char(ch: char) -> bool {
    ch.is_alphanumeric() || ch == '_'
}

fn slice_chars(text: &str, range: Range<usize>) -> String {
    text.chars()
        .skip(range.start)
        .take(range.end.saturating_sub(range.start))
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn opts(match_case: bool, whole_word: bool) -> FindOptions {
        FindOptions {
            match_case,
            whole_word,
        }
    }

    #[test]
    fn next_match_starts_from_cursor() {
        assert_eq!(
            find_next_match("alpha beta alpha", "alpha", 1, opts(false, false)),
            Some(11..16)
        );
    }

    #[test]
    fn previous_match_wraps() {
        assert_eq!(
            find_previous_match("alpha beta alpha", "alpha", 0, opts(false, false)),
            Some(11..16)
        );
    }

    #[test]
    fn case_insensitive_match_finds_text() {
        assert_eq!(
            find_next_match("Alpha", "alpha", 0, opts(false, false)),
            Some(0..5)
        );
    }

    #[test]
    fn match_case_can_miss() {
        assert_eq!(
            find_next_match("Alpha", "alpha", 0, opts(true, false)),
            None
        );
    }

    #[test]
    fn whole_word_excludes_embedded_matches() {
        assert_eq!(
            find_next_match("alphabet alpha", "alpha", 0, opts(false, true)),
            Some(9..14)
        );
    }

    #[test]
    fn replace_all_handles_offset_drift() {
        let mut document = DocumentState::bootstrap();
        document.replace_with_runs(
            "Test".to_owned(),
            vec![crate::document::TextRun {
                text: "one two one".to_owned(),
                style: Default::default(),
            }],
        );
        let mut canvas = CanvasState::default();
        let mut history = ChangeHistory::new();
        let mut state = FindReplaceState {
            find_text: "one".to_owned(),
            replace_text: "three".to_owned(),
            ..Default::default()
        };

        assert_eq!(
            replace_all(&mut document, &mut canvas, &mut history, &mut state),
            2
        );
        assert_eq!(document.plain_text(), "three two three");
    }
}
