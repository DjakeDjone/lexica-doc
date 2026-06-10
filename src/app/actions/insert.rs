use eframe::egui;
#[cfg(not(target_arch = "wasm32"))]
use rfd::FileDialog;
#[cfg(not(target_arch = "wasm32"))]
use std::fs;
#[cfg(not(target_arch = "wasm32"))]
use std::path::PathBuf;

use crate::app::{CanvasState, ChangeHistory};
#[cfg(not(target_arch = "wasm32"))]
use crate::document::DocumentImage;
use crate::document::DocumentState;

pub fn insert_page_break(
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

pub fn insert_section_break(
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
    let paragraph_index = document
        .paragraphs()
        .iter()
        .position(|paragraph| {
            paragraph.range.contains(&cursor_index) || paragraph.range.start == cursor_index
        })
        .unwrap_or_else(|| document.paragraph_count().saturating_sub(1));
    let section_id = document.insert_section_break_before_paragraph(paragraph_index);
    canvas.selection = egui::text_selection::CCursorRange::one(
        egui::epaint::text::cursor::CCursor::new(cursor_index),
    );
    canvas.active_header_footer = None;
    canvas.active_style = document.typing_style_at(cursor_index);
    canvas.active_paragraph_style = document.paragraph_style_at(cursor_index);
    *status_message = format!("Inserted section break before Section {section_id}");
}

pub fn insert_image(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    #[cfg(not(target_arch = "wasm32"))] dialog_tx: &std::sync::mpsc::Sender<
        crate::app::DialogAction,
    >,
) {
    #[cfg(target_arch = "wasm32")]
    {
        *status_message = "Inserting local images is not available in the web build yet".to_owned();
        let _ = (document, canvas, history);
    }

    #[cfg(not(target_arch = "wasm32"))]
    {
        let _ = (document, canvas, status_message, history);
        let tx = dialog_tx.clone();
        std::thread::spawn(move || {
            if let Some(path) = FileDialog::new()
                .add_filter("images", &["png", "jpg", "jpeg", "gif", "bmp"])
                .pick_file()
            {
                let _ = tx.send(crate::app::DialogAction::InsertImage(path));
            }
        });
    }
}

#[cfg(not(target_arch = "wasm32"))]
pub fn finish_insert_image(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    path: &std::path::PathBuf,
) {
    let image = match load_image_for_document(path, document) {
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

pub fn insert_table(
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

pub fn insert_table_row(
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

pub fn insert_table_column(
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

pub fn delete_table_row(
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

pub fn delete_table_column(
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
