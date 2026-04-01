use std::path::PathBuf;

use eframe::egui;
use rfd::FileDialog;

use crate::document::{
    CharacterStyle, DocumentState, FontChoice, ListKind, ParagraphAlignment, ParagraphStyle,
};

use super::CanvasState;

pub(super) fn open_document(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
) {
    if let Some(path) = FileDialog::new()
        .add_filter("supported", &["txt", "md", "markdown", "docx"])
        .pick_file()
    {
        match DocumentState::load_from_path(&path) {
            Ok(new_document) => {
                *document = new_document;
                canvas.selection = egui::text_selection::CCursorRange::default();
                canvas.active_style = CharacterStyle::default();
                canvas.active_paragraph_style = ParagraphStyle::default();
                canvas.zoom = 1.0;
                canvas.pan = egui::Vec2::ZERO;
                *current_path = match path.extension().and_then(|ext| ext.to_str()) {
                    Some("docx") => None,
                    _ => Some(path.clone()),
                };
                *status_message = format!(
                    "Imported {}",
                    path.file_name()
                        .and_then(|name| name.to_str())
                        .unwrap_or("document")
                );
            }
            Err(error) => *status_message = error,
        }
    }
}

pub(super) fn save_document(
    document: &DocumentState,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
) {
    let path = match current_path.clone() {
        Some(path) => path,
        None => match FileDialog::new()
            .add_filter("text", &["txt"])
            .add_filter("markdown", &["md"])
            .set_file_name(&document.title)
            .save_file()
        {
            Some(path) => path,
            None => return,
        },
    };

    match document.save_to_path(&path) {
        Ok(()) => {
            *current_path = Some(path.clone());
            *status_message = format!(
                "Saved {}",
                path.file_name()
                    .and_then(|name| name.to_str())
                    .unwrap_or("document")
            );
        }
        Err(error) => *status_message = error,
    }
}

pub(super) fn handle_global_shortcuts(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    current_path: &mut Option<PathBuf>,
    status_message: &mut String,
) {
    if ui.input_mut(|input| input.consume_key(egui::Modifiers::COMMAND, egui::Key::S)) {
        save_document(document, status_message, current_path);
    }
}

pub(super) fn toggle_bold(document: &mut DocumentState, canvas: &mut CanvasState) {
    let next_value = !canvas.active_style.bold;
    apply_selection_or_active_style(document, canvas, move |style| style.bold = next_value);
}

pub(super) fn toggle_italic(document: &mut DocumentState, canvas: &mut CanvasState) {
    let next_value = !canvas.active_style.italic;
    apply_selection_or_active_style(document, canvas, move |style| style.italic = next_value);
}

pub(super) fn toggle_underline(document: &mut DocumentState, canvas: &mut CanvasState) {
    let next_value = !canvas.active_style.underline;
    apply_selection_or_active_style(document, canvas, move |style| style.underline = next_value);
}

pub(super) fn toggle_strikethrough(document: &mut DocumentState, canvas: &mut CanvasState) {
    let next_value = !canvas.active_style.strikethrough;
    apply_selection_or_active_style(document, canvas, move |style| {
        style.strikethrough = next_value
    });
}

pub(super) fn set_font_size(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    font_size: f32,
) {
    apply_selection_or_active_style(document, canvas, move |style| {
        style.font_size_points = font_size
    });
}

pub(super) fn set_font_choice(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    font_choice: FontChoice,
) {
    apply_selection_or_active_style(document, canvas, move |style| {
        style.font_choice = font_choice
    });
}

pub(super) fn set_text_color(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    color: egui::Color32,
) {
    apply_selection_or_active_style(document, canvas, move |style| style.text_color = color);
}

pub(super) fn set_highlight_color(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    color: egui::Color32,
) {
    apply_selection_or_active_style(document, canvas, move |style| style.highlight_color = color);
}

pub(super) fn sync_active_style(document: &DocumentState, canvas: &mut CanvasState) {
    let range = canvas.selection.as_sorted_char_range();
    let cursor_index = if range.start < range.end {
        range.end
    } else {
        canvas.selection.primary.index
    };
    canvas.active_style = document.typing_style_at(cursor_index);
    canvas.active_paragraph_style = document.paragraph_style_at(cursor_index);
}

pub(super) fn set_paragraph_alignment(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    alignment: ParagraphAlignment,
) {
    apply_selection_or_current_paragraph(document, canvas, move |style| {
        style.alignment = alignment
    });
}

pub(super) fn toggle_bullet_list(document: &mut DocumentState, canvas: &mut CanvasState) {
    let next = if canvas.active_paragraph_style.list_kind == ListKind::Bullet {
        ListKind::None
    } else {
        ListKind::Bullet
    };
    apply_selection_or_current_paragraph(document, canvas, move |style| style.list_kind = next);
}

pub(super) fn toggle_ordered_list(document: &mut DocumentState, canvas: &mut CanvasState) {
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
    let range = canvas.selection.as_sorted_char_range();
    if range.start < range.end {
        document.apply_style_to_range(range, mutate);
    }
    mutate(&mut canvas.active_style);
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
