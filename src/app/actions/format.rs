use eframe::egui;

use crate::document::{
    CharacterStyle, DocumentState, FontChoice, ListKind, ParagraphAlignment, ParagraphStyle,
    TextRun,
};
use crate::app::{ActiveHeaderFooter, CanvasState, ChangeHistory};

pub fn toggle_bold(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    let next_value = !canvas.active_style.bold;
    apply_selection_or_active_style(document, canvas, move |style| style.bold = next_value);
}

pub fn toggle_italic(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    let next_value = !canvas.active_style.italic;
    apply_selection_or_active_style(document, canvas, move |style| style.italic = next_value);
}

pub fn toggle_underline(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    let next_value = !canvas.active_style.underline;
    apply_selection_or_active_style(document, canvas, move |style| style.underline = next_value);
}

pub fn toggle_strikethrough(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    let next_value = !canvas.active_style.strikethrough;
    apply_selection_or_active_style(document, canvas, move |style| {
        style.strikethrough = next_value
    });
}

pub fn set_font_size(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    font_size: f32,
    history: &mut ChangeHistory,
    now: f64,
) {
    history.checkpoint_coalesced(document, now);
    apply_selection_or_active_style(document, canvas, move |style| {
        style.font_size_points = font_size
    });
}

pub fn set_font_choice(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    font_choice: FontChoice,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    apply_selection_or_active_style(document, canvas, move |style| {
        style.font_choice = font_choice;
        style.font_family_name = font_choice.family_name();
    });
}

pub fn set_text_color(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    color: egui::Color32,
    history: &mut ChangeHistory,
    now: f64,
) {
    history.checkpoint_coalesced(document, now);
    apply_selection_or_active_style(document, canvas, move |style| style.text_color = color);
}

pub fn set_highlight_color(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    color: egui::Color32,
    history: &mut ChangeHistory,
    now: f64,
) {
    history.checkpoint_coalesced(document, now);
    apply_selection_or_active_style(document, canvas, move |style| style.highlight_color = color);
}

pub fn sync_active_style(document: &DocumentState, canvas: &mut CanvasState) {
    if let Some(active) = canvas.active_header_footer {
        canvas.active_style = header_footer_selection_style_at(
            active_header_footer_runs(document, active),
            canvas.active_header_footer_selection.as_sorted_char_range(),
            canvas.active_header_footer_selection.primary.index,
        );
        return;
    }

    if let Some((table_id, row, col)) = canvas.active_table_cell {
        if let Some(style) = document.table_cell_selection_style_at(
            table_id,
            row,
            col,
            canvas.table_cell_selection.as_sorted_char_range(),
        ) {
            canvas.active_style = style;
        }
        return;
    }

    let range = canvas.selection.as_sorted_char_range();
    canvas.active_style = document.selection_style_at(range.clone());
    canvas.active_paragraph_style = document.paragraph_style_at(canvas.selection.primary.index);
}

pub fn set_paragraph_alignment(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    alignment: ParagraphAlignment,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    apply_selection_or_current_paragraph(document, canvas, move |style| {
        style.alignment = alignment
    });
}

pub fn toggle_bullet_list(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    let next = if canvas.active_paragraph_style.list_kind == ListKind::Bullet {
        ListKind::None
    } else {
        ListKind::Bullet
    };
    apply_selection_or_current_paragraph(document, canvas, move |style| style.list_kind = next);
}

pub fn toggle_ordered_list(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    let next = if canvas.active_paragraph_style.list_kind == ListKind::Ordered {
        ListKind::None
    } else {
        ListKind::Ordered
    };
    apply_selection_or_current_paragraph(document, canvas, move |style| style.list_kind = next);
}

fn apply_selection_or_active_style(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    mutate: impl Fn(&mut CharacterStyle) + Copy,
) {
    if let Some(active) = canvas.active_header_footer {
        let range = canvas.active_header_footer_selection.as_sorted_char_range();
        if range.start < range.end {
            apply_style_to_header_footer_range(
                active_header_footer_runs_mut(document, active),
                range,
                mutate,
            );
            document.sync_compat_from_first_section();
        }
        mutate(&mut canvas.active_style);
        return;
    }

    if let Some((table_id, row, col)) = canvas.active_table_cell {
        let range = canvas.table_cell_selection.as_sorted_char_range();
        if range.start < range.end {
            document.apply_style_to_table_cell_range(table_id, row, col, range, mutate);
        } else if document
            .table_cell_len(table_id, row, col)
            .is_some_and(|len| len == 0)
        {
            document.apply_style_to_table_cell(table_id, row, col, mutate);
        }
        mutate(&mut canvas.active_style);
        return;
    }

    let range = canvas.selection.as_sorted_char_range();
    if range.start < range.end {
        document.apply_style_to_range(range, mutate);
    }
    mutate(&mut canvas.active_style);
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

fn header_footer_selection_style_at(
    runs: &[TextRun],
    range: std::ops::Range<usize>,
    cursor: usize,
) -> CharacterStyle {
    if range.start < range.end {
        return header_footer_style_at(runs, range.start);
    }
    header_footer_style_at(runs, cursor)
}

fn header_footer_style_at(runs: &[TextRun], char_index: usize) -> CharacterStyle {
    let target = char_index.min(runs_total_chars(runs));
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

fn apply_style_to_header_footer_range(
    runs: &mut Vec<TextRun>,
    range: std::ops::Range<usize>,
    mutate: impl Fn(&mut CharacterStyle) + Copy,
) {
    let start = range.start.min(runs_total_chars(runs));
    let end = range.end.min(runs_total_chars(runs));
    if start >= end {
        return;
    }
    split_header_footer_runs_at(runs, start);
    split_header_footer_runs_at(runs, end);

    let mut offset = 0usize;
    for run in runs.iter_mut() {
        let len = run.text.chars().count();
        if offset >= start && offset + len <= end {
            mutate(&mut run.style);
        }
        offset += len;
    }
    normalize_header_footer_runs(runs);
}

fn split_header_footer_runs_at(runs: &mut Vec<TextRun>, char_index: usize) {
    if char_index == 0 || char_index >= runs_total_chars(runs) {
        return;
    }

    let mut offset = 0usize;
    for index in 0..runs.len() {
        let len = runs[index].text.chars().count();
        if char_index > offset && char_index < offset + len {
            let local_index = char_index - offset;
            let byte_index = char_to_byte_index(&runs[index].text, local_index);
            let right = runs[index].text.split_off(byte_index);
            let style = runs[index].style;
            runs.insert(index + 1, TextRun { text: right, style });
            break;
        }
        offset += len;
    }
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

fn runs_total_chars(runs: &[TextRun]) -> usize {
    runs.iter().map(|run| run.text.chars().count()).sum()
}

fn char_to_byte_index(text: &str, char_index: usize) -> usize {
    text.char_indices()
        .nth(char_index)
        .map(|(byte_index, _)| byte_index)
        .unwrap_or(text.len())
}

fn apply_selection_or_current_paragraph(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    mutate: impl Fn(&mut ParagraphStyle) + Copy,
) {
    let range = canvas.selection.as_sorted_char_range();
    document.apply_paragraph_style_to_range(range, mutate);
    canvas.active_paragraph_style = document.paragraph_style_at(canvas.selection.primary.index);
}
