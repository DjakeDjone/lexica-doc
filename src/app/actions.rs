use std::path::PathBuf;

use eframe::egui;
#[cfg(not(target_arch = "wasm32"))]
use rfd::FileDialog;
#[cfg(not(target_arch = "wasm32"))]
use std::fs;
#[cfg(target_arch = "wasm32")]
use wasm_bindgen::{JsCast as _, JsValue};

#[cfg(not(target_arch = "wasm32"))]
use crate::document::DocumentImage;
use crate::document::{
    CharacterStyle, DocumentState, FontChoice, ListKind, ParagraphAlignment, ParagraphStyle,
};

#[cfg(not(target_arch = "wasm32"))]
use super::ZoomMode;
use super::{CanvasState, ChangeHistory};

pub(super) fn open_document(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
    history: &mut ChangeHistory,
) {
    #[cfg(target_arch = "wasm32")]
    {
        *status_message = "Opening local files is not available in the web build yet".to_owned();
        let _ = (document, canvas, current_path, history);
    }

    #[cfg(not(target_arch = "wasm32"))]
    if let Some(path) = FileDialog::new()
        .add_filter("supported", &["txt", "md", "markdown", "docx"])
        .pick_file()
    {
        match DocumentState::load_from_path(&path) {
            Ok(new_document) => {
                let imported_docx =
                    matches!(path.extension().and_then(|ext| ext.to_str()), Some("docx"));
                history.clear();
                *document = new_document;
                canvas.selection = egui::text_selection::CCursorRange::default();
                canvas.active_style = CharacterStyle::default();
                canvas.active_paragraph_style = ParagraphStyle::default();
                canvas.zoom = 1.0;
                canvas.zoom_mode = if imported_docx {
                    ZoomMode::FitPage
                } else {
                    ZoomMode::Manual
                };
                canvas.imported_docx_view = imported_docx;
                canvas.pan = egui::Vec2::ZERO;
                canvas.image_textures.clear();
                canvas.selected_image_id = None;
                canvas.image_rects.clear();
                canvas.resize_drag = None;
                canvas.move_drag = None;
                canvas.active_table_cell = None;
                canvas.table_cell_rects.clear();
                canvas.table_cell_content_rects.clear();
                canvas.table_cell_selection = egui::text_selection::CCursorRange::default();
                canvas.table_resize_handles.clear();
                canvas.table_resize_drag = None;
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
    #[cfg(target_arch = "wasm32")]
    {
        match download_document(document) {
            Ok(filename) => *status_message = format!("Downloaded {filename}"),
            Err(error) => *status_message = error,
        }
        let _ = current_path;
    }

    #[cfg(not(target_arch = "wasm32"))]
    {
        let path = match current_path.clone() {
            Some(path) => path,
            None => match pick_save_path(document) {
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
}

pub(super) fn save_document_as(
    document: &DocumentState,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
) {
    #[cfg(target_arch = "wasm32")]
    {
        match download_document(document) {
            Ok(filename) => *status_message = format!("Downloaded {filename}"),
            Err(error) => *status_message = error,
        }
        let _ = current_path;
    }

    #[cfg(not(target_arch = "wasm32"))]
    {
        let Some(path) = pick_save_path(document) else {
            return;
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
}

#[cfg(not(target_arch = "wasm32"))]
fn pick_save_path(document: &DocumentState) -> Option<PathBuf> {
    FileDialog::new()
        .add_filter("text", &["txt"])
        .add_filter("markdown", &["md", "markdown"])
        .add_filter("web (formatted)", &["html", "htm"])
        .add_filter("pdf", &["pdf"])
        .set_file_name(&document.title)
        .save_file()
}

#[cfg(target_arch = "wasm32")]
fn download_document(document: &DocumentState) -> Result<String, String> {
    let filename = download_filename(&document.title, "html");
    let bytes = document.export_bytes_for_extension("html")?;
    download_bytes(&filename, "text/html;charset=utf-8", &bytes)?;
    Ok(filename)
}

#[cfg(target_arch = "wasm32")]
fn download_filename(title: &str, extension: &str) -> String {
    let mut stem: String = title
        .chars()
        .map(|ch| {
            if ch.is_ascii_alphanumeric() || matches!(ch, '-' | '_' | ' ') {
                ch
            } else {
                '-'
            }
        })
        .collect();
    stem = stem.trim().replace(' ', "-");
    while stem.contains("--") {
        stem = stem.replace("--", "-");
    }
    let stem = stem.trim_matches('-');
    let stem = if stem.is_empty() { "document" } else { stem };
    format!("{stem}.{extension}")
}

#[cfg(target_arch = "wasm32")]
fn download_bytes(filename: &str, mime_type: &str, bytes: &[u8]) -> Result<(), String> {
    let window = web_sys::window().ok_or_else(|| "Browser window is unavailable".to_owned())?;
    let document = window
        .document()
        .ok_or_else(|| "Browser document is unavailable".to_owned())?;
    let body = document
        .body()
        .ok_or_else(|| "Browser document body is unavailable".to_owned())?;

    let byte_array = js_sys::Uint8Array::from(bytes);
    let blob_parts = js_sys::Array::new();
    blob_parts.push(&byte_array.buffer());

    let blob_options = web_sys::BlobPropertyBag::new();
    blob_options.set_type(mime_type);
    let blob = web_sys::Blob::new_with_u8_array_sequence_and_options(&blob_parts, &blob_options)
        .map_err(js_error_message)?;
    let url = web_sys::Url::create_object_url_with_blob(&blob).map_err(js_error_message)?;

    let anchor = document
        .create_element("a")
        .map_err(js_error_message)?
        .dyn_into::<web_sys::HtmlAnchorElement>()
        .map_err(|_| "Failed to create download link".to_owned())?;
    anchor.set_href(&url);
    anchor.set_download(filename);
    anchor
        .style()
        .set_property("display", "none")
        .map_err(js_error_message)?;

    body.append_child(&anchor).map_err(js_error_message)?;
    anchor.click();
    anchor.remove();
    web_sys::Url::revoke_object_url(&url).map_err(js_error_message)?;
    Ok(())
}

#[cfg(target_arch = "wasm32")]
fn js_error_message(value: JsValue) -> String {
    value
        .as_string()
        .unwrap_or_else(|| "Browser download failed".to_owned())
}

pub(super) fn insert_page_break(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
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
    history: &mut ChangeHistory,
) {
    #[cfg(target_arch = "wasm32")]
    {
        *status_message = "Inserting local images is not available in the web build yet".to_owned();
        let _ = (document, canvas, history);
    }

    #[cfg(not(target_arch = "wasm32"))]
    {
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

        history.checkpoint(document, f64::NAN);
        if let Some((table_id, row, col)) = canvas.active_table_cell {
            document.insert_table_cell_image(table_id, row, col, image, canvas.active_style);
            if let Some(len) = document.table_cell_len(table_id, row, col) {
                canvas.table_cell_selection = egui::text_selection::CCursorRange::one(
                    egui::epaint::text::cursor::CCursor::new(len),
                );
            }
            canvas.selected_image_id = None;
            canvas.resize_drag = None;
            canvas.move_drag = None;
            canvas.table_resize_drag = None;
            canvas.image_textures.clear();
            *status_message = format!(
                "Inserted {} into table cell",
                path.file_name()
                    .and_then(|name| name.to_str())
                    .unwrap_or("image")
            );
            return;
        }

        let selected = canvas.selection.as_sorted_char_range();
        let insert_at = selected.start;
        if selected.start < selected.end {
            document.delete_range(selected);
        }

        let image_id = image.id;
        let cursor_index = document.insert_image(insert_at, image);
        canvas.selection = egui::text_selection::CCursorRange::one(
            egui::epaint::text::cursor::CCursor::new(cursor_index),
        );
        canvas.active_style = document.typing_style_at(cursor_index);
        canvas.active_paragraph_style = document.paragraph_style_at(cursor_index);
        canvas.selected_image_id = Some(image_id);
        canvas.active_table_cell = None;
        canvas.resize_drag = None;
        canvas.move_drag = None;
        canvas.table_resize_drag = None;
        canvas.image_textures.clear();
        *status_message = format!(
            "Inserted {}",
            path.file_name()
                .and_then(|name| name.to_str())
                .unwrap_or("image")
        );
    }
}

pub(super) fn handle_global_shortcuts(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    current_path: &mut Option<PathBuf>,
    status_message: &mut String,
) -> bool {
    let mut document_changed = false;

    if ui.input_mut(|input| input.consume_key(egui::Modifiers::COMMAND, egui::Key::S)) {
        save_document(document, status_message, current_path);
    }
    if ui.input_mut(|input| {
        input.consume_key(
            egui::Modifiers::COMMAND | egui::Modifiers::SHIFT,
            egui::Key::S,
        )
    }) {
        save_document_as(document, status_message, current_path);
    }
    if ui.input_mut(|input| input.consume_key(egui::Modifiers::COMMAND, egui::Key::Z)) {
        if ui.input(|i| i.modifiers.shift) {
            if history.redo(document) {
                canvas.image_textures.clear();
                *status_message = "Redo".to_owned();
                document_changed = true;
            }
        } else if history.undo(document) {
            canvas.image_textures.clear();
            *status_message = "Undo".to_owned();
            document_changed = true;
        }
    }
    let shift_redo_pressed = ui.input_mut(|input| {
        input.consume_key(
            egui::Modifiers::COMMAND | egui::Modifiers::SHIFT,
            egui::Key::Z,
        )
    });
    if shift_redo_pressed && history.redo(document) {
        canvas.image_textures.clear();
        *status_message = "Redo".to_owned();
        document_changed = true;
    }
    // Ctrl+Y as an alternative redo shortcut
    let redo_pressed =
        ui.input_mut(|input| input.consume_key(egui::Modifiers::COMMAND, egui::Key::Y));
    if redo_pressed && history.redo(document) {
        canvas.image_textures.clear();
        *status_message = "Redo".to_owned();
        document_changed = true;
    }

    document_changed
}

pub(super) fn set_image_opacity(
    document: &mut DocumentState,
    image_id: usize,
    opacity: f32,
    status_message: &mut String,
    history: &mut ChangeHistory,
    now: f64,
) {
    history.checkpoint_coalesced(document, now);
    document.set_image_opacity(image_id, opacity);
    *status_message = format!("Opacity: {:.0}%", opacity * 100.0);
}

pub(super) fn set_image_wrap_mode(
    document: &mut DocumentState,
    image_id: usize,
    wrap_mode: crate::document::WrapMode,
    status_message: &mut String,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    document.set_image_wrap_mode(image_id, wrap_mode);
    *status_message = format!("Wrap: {}", wrap_mode.label());
}

pub(super) fn set_image_rendering(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    image_id: usize,
    rendering: crate::document::ImageRendering,
    status_message: &mut String,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
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
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
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

pub(super) fn toggle_bold(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    let next_value = !canvas.active_style.bold;
    apply_selection_or_active_style(document, canvas, move |style| style.bold = next_value);
}

pub(super) fn toggle_italic(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    let next_value = !canvas.active_style.italic;
    apply_selection_or_active_style(document, canvas, move |style| style.italic = next_value);
}

pub(super) fn toggle_underline(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    let next_value = !canvas.active_style.underline;
    apply_selection_or_active_style(document, canvas, move |style| style.underline = next_value);
}

pub(super) fn toggle_strikethrough(
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

pub(super) fn set_font_size(
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

pub(super) fn set_font_choice(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    font_choice: FontChoice,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    apply_selection_or_active_style(document, canvas, move |style| {
        style.font_choice = font_choice
    });
}

pub(super) fn set_text_color(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    color: egui::Color32,
    history: &mut ChangeHistory,
    now: f64,
) {
    history.checkpoint_coalesced(document, now);
    apply_selection_or_active_style(document, canvas, move |style| style.text_color = color);
}

pub(super) fn set_highlight_color(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    color: egui::Color32,
    history: &mut ChangeHistory,
    now: f64,
) {
    history.checkpoint_coalesced(document, now);
    apply_selection_or_active_style(document, canvas, move |style| style.highlight_color = color);
}

pub(super) fn sync_active_style(document: &DocumentState, canvas: &mut CanvasState) {
    if let Some((table_id, row, col)) = canvas.active_table_cell {
        if let Some(style) = document.table_cell_style_at(
            table_id,
            row,
            col,
            canvas.table_cell_selection.primary.index,
        ) {
            canvas.active_style = style;
        }
        return;
    }

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
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    apply_selection_or_current_paragraph(document, canvas, move |style| {
        style.alignment = alignment
    });
}

pub(super) fn toggle_bullet_list(
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

pub(super) fn toggle_ordered_list(
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

#[cfg(not(target_arch = "wasm32"))]
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
    let next_id = document.next_image_id();

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
        lock_aspect_ratio: true,
        opacity: 1.0,
        layout_mode: crate::document::ImageLayoutMode::Inline,
        wrap_mode: crate::document::WrapMode::Inline,
        rendering: crate::document::ImageRendering::Smooth,
        horizontal_position: Default::default(),
        vertical_position: Default::default(),
        distance_from_text: Default::default(),
        z_index: 0,
        move_with_text: true,
        allow_overlap: false,
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

pub(super) fn insert_table(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    num_rows: usize,
    num_cols: usize,
    status_message: &mut String,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    let selected = canvas.selection.as_sorted_char_range();
    let insert_at = selected.start;
    if selected.start < selected.end {
        document.delete_range(selected);
    }

    let cursor_index = document.insert_table(insert_at, num_rows, num_cols);
    canvas.selection = egui::text_selection::CCursorRange::one(
        egui::epaint::text::cursor::CCursor::new(cursor_index),
    );
    if let Some(table_id) = document
        .paragraph_tables
        .iter()
        .flatten()
        .map(|table| table.id)
        .max()
    {
        canvas.active_table_cell = Some((table_id, 0, 0));
        canvas.table_cell_selection = egui::text_selection::CCursorRange::default();
    }
    canvas.selected_image_id = None;
    canvas.active_style = document.typing_style_at(cursor_index);
    canvas.active_paragraph_style = document.paragraph_style_at(cursor_index);
    *status_message = format!("Inserted {}×{} table", num_rows, num_cols);
}

#[allow(dead_code)]
pub(super) fn insert_table_row(
    document: &mut DocumentState,
    table_id: usize,
    after_row: usize,
    status_message: &mut String,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    document.insert_table_row(table_id, after_row);
    *status_message = "Row inserted".to_owned();
}

#[allow(dead_code)]
pub(super) fn insert_table_column(
    document: &mut DocumentState,
    table_id: usize,
    after_col: usize,
    status_message: &mut String,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    document.insert_table_column(table_id, after_col);
    *status_message = "Column inserted".to_owned();
}

#[allow(dead_code)]
pub(super) fn delete_table_row(
    document: &mut DocumentState,
    table_id: usize,
    row_index: usize,
    status_message: &mut String,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    document.delete_table_row(table_id, row_index);
    *status_message = "Row deleted".to_owned();
}

#[allow(dead_code)]
pub(super) fn delete_table_column(
    document: &mut DocumentState,
    table_id: usize,
    col_index: usize,
    status_message: &mut String,
    history: &mut ChangeHistory,
) {
    history.checkpoint(document, f64::NAN);
    document.delete_table_column(table_id, col_index);
    *status_message = "Column deleted".to_owned();
}
