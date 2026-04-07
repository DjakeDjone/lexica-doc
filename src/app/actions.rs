use std::{fs, path::PathBuf};

use eframe::egui;
use rfd::FileDialog;

use crate::document::{
    CharacterStyle, DocumentImage, DocumentState, FontChoice, ListKind, ParagraphAlignment,
    ParagraphStyle,
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
                canvas.image_textures.clear();
                canvas.selected_image_id = None;
                canvas.image_rects.clear();
                canvas.resize_drag = None;
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

pub(super) fn insert_page_break(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
) {
    let selected = canvas.selection.as_sorted_char_range();
    let insert_at = selected.start;
    if selected.start < selected.end {
        document.delete_range(selected);
    }

    let cursor_index = document.insert_page_break(insert_at);
    canvas.selection = egui::text_selection::CCursorRange::one(
        egui::epaint::text::cursor::CCursor::new(cursor_index),
    );
    canvas.active_style = document.typing_style_at(cursor_index);
    canvas.active_paragraph_style = document.paragraph_style_at(cursor_index);
    *status_message = "Inserted page break".to_owned();
}

pub(super) fn insert_image(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
) {
    let Some(path) = FileDialog::new()
        .add_filter("images", &["png", "jpg", "jpeg", "gif", "bmp"])
        .pick_file()
    else {
        return;
    };

    let image = match load_image_for_document(&path, document) {
        Ok(image) => image,
        Err(error) => {
            *status_message = error;
            return;
        }
    };

    let selected = canvas.selection.as_sorted_char_range();
    let insert_at = selected.start;
    if selected.start < selected.end {
        document.delete_range(selected);
    }

    let cursor_index = document.insert_image(insert_at, image);
    canvas.selection = egui::text_selection::CCursorRange::one(
        egui::epaint::text::cursor::CCursor::new(cursor_index),
    );
    canvas.active_style = document.typing_style_at(cursor_index);
    canvas.active_paragraph_style = document.paragraph_style_at(cursor_index);
    canvas.image_textures.clear();
    *status_message = format!(
        "Inserted {}",
        path.file_name()
            .and_then(|name| name.to_str())
            .unwrap_or("image")
    );
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

pub(super) fn set_image_opacity(
    document: &mut DocumentState,
    image_id: usize,
    opacity: f32,
    status_message: &mut String,
) {
    document.set_image_opacity(image_id, opacity);
    *status_message = format!("Opacity: {:.0}%", opacity * 100.0);
}

pub(super) fn set_image_wrap_mode(
    document: &mut DocumentState,
    image_id: usize,
    wrap_mode: crate::document::WrapMode,
    status_message: &mut String,
) {
    document.set_image_wrap_mode(image_id, wrap_mode);
    *status_message = format!("Wrap: {}", wrap_mode.label());
}

pub(super) fn set_image_rendering(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    image_id: usize,
    rendering: crate::document::ImageRendering,
    status_message: &mut String,
) {
    document.set_image_rendering(image_id, rendering);
    // Clear both possible cache entries for this image so texture is rebuilt
    canvas.image_textures.remove(&(image_id * 2));
    canvas.image_textures.remove(&(image_id * 2 + 1));
    *status_message = match rendering {
        crate::document::ImageRendering::Smooth => "Rendering: Smooth".to_owned(),
        crate::document::ImageRendering::Crisp => "Rendering: Crisp".to_owned(),
    };
}

pub(super) fn reset_image_size(
    document: &mut DocumentState,
    _canvas: &mut CanvasState,
    image_id: usize,
    status_message: &mut String,
) {
    let image_bytes = document
        .paragraph_images
        .iter()
        .flatten()
        .find(|img| img.id == image_id)
        .map(|img| img.bytes.clone());

    let Some(bytes) = image_bytes else {
        return;
    };

    match image::load_from_memory(&bytes) {
        Ok(decoded) => {
            let w = (decoded.width() as f32 * 0.75).clamp(24.0, document.page_size.width_points);
            let h = (decoded.height() as f32 * 0.75).clamp(24.0, document.page_size.height_points);
            document.resize_image_by_id(image_id, w, h);
            *status_message = format!("Image size reset to {:.0} × {:.0} pt", w, h);
        }
        Err(error) => {
            *status_message = format!("Could not decode image: {error}");
        }
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

fn load_image_for_document(
    path: &PathBuf,
    document: &DocumentState,
) -> Result<DocumentImage, String> {
    let bytes =
        fs::read(path).map_err(|error| format!("failed to read {}: {error}", path.display()))?;
    let decoded = image::load_from_memory(&bytes)
        .map_err(|error| format!("failed to decode {}: {error}", path.display()))?;
    let width_points = (decoded.width() as f32 * 0.75).clamp(24.0, document.page_size.width_points);
    let height_points =
        (decoded.height() as f32 * 0.75).clamp(24.0, document.page_size.height_points);
    let next_id = document
        .paragraph_images
        .iter()
        .flatten()
        .map(|image| image.id)
        .max()
        .unwrap_or(0)
        + 1;

    Ok(DocumentImage {
        id: next_id,
        bytes,
        alt_text: path
            .file_stem()
            .and_then(|name| name.to_str())
            .unwrap_or("Image")
            .to_owned(),
        width_points,
        height_points,
        opacity: 1.0,
        wrap_mode: crate::document::WrapMode::Inline,
        rendering: crate::document::ImageRendering::Smooth,
    })
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
