pub mod file;
pub mod format;
pub mod insert;

use std::path::PathBuf;
use eframe::egui;

use crate::document::DocumentState;
use super::{CanvasState, ChangeHistory};

// Re-export public functions to keep the parent module's API stable
pub use file::{
    open_document, save_document, save_document_as, save_document_as_with_name,
};
#[cfg(not(target_arch = "wasm32"))]
pub use file::open_document_from_path;

pub use format::{
    set_font_choice, set_font_size, set_highlight_color, set_paragraph_alignment,
    set_text_color, sync_active_style, toggle_bold, toggle_bullet_list, toggle_italic,
    toggle_ordered_list, toggle_strikethrough, toggle_underline,
};

pub use insert::{
    delete_table_column, delete_table_row, insert_image, insert_page_break,
    insert_section_break, insert_table, insert_table_column, insert_table_row,
};

fn redo(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
) -> bool {
    if history.redo(document) {
        canvas.image_textures.clear();
        *status_message = "Redo".to_owned();
        true
    } else {
        false
    }
}

pub fn handle_global_shortcuts(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    current_path: &mut Option<PathBuf>,
    status_message: &mut String,
) -> bool {
    let mut document_changed = false;

    if ui.input_mut(|input| input.consume_key(egui::Modifiers::COMMAND, egui::Key::S)) {
        let _ = save_document(document, status_message, current_path);
    }
    if ui.input_mut(|input| {
        input.consume_key(
            egui::Modifiers::COMMAND | egui::Modifiers::SHIFT,
            egui::Key::S,
        )
    }) {
        let _ = save_document_as(document, status_message, current_path);
    }
    if ui.input_mut(|input| input.consume_key(egui::Modifiers::COMMAND, egui::Key::Z)) {
        if ui.input(|i| i.modifiers.shift) {
            document_changed |= redo(document, canvas, status_message, history);
        } else if history.undo(document) {
            canvas.image_textures.clear();
            *status_message = "Undo".to_owned();
            document_changed = true;
        }
    }
    let redo_shortcut = ui.input_mut(|input| {
        input.consume_key(egui::Modifiers::COMMAND | egui::Modifiers::SHIFT, egui::Key::Z)
            || input.consume_key(egui::Modifiers::COMMAND, egui::Key::Y)
    });
    if redo_shortcut {
        document_changed |= redo(document, canvas, status_message, history);
    }

    document_changed
}

pub fn set_image_opacity(
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

pub fn set_image_wrap_mode(
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

pub fn set_image_rendering(
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

pub fn reset_image_size(
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
