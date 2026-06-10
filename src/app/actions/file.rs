#[cfg(not(target_arch = "wasm32"))]
use std::path::Path;
use std::path::PathBuf;

use eframe::egui;
#[cfg(not(target_arch = "wasm32"))]
use rfd::FileDialog;
#[cfg(target_arch = "wasm32")]
use wasm_bindgen::{JsCast as _, JsValue};

use crate::app::{CanvasState, ChangeHistory};
use crate::document::DocumentState;
#[cfg(not(target_arch = "wasm32"))]
use crate::document::{CharacterStyle, ParagraphStyle};

pub fn open_document(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
    history: &mut ChangeHistory,
    #[cfg(not(target_arch = "wasm32"))] dialog_tx: &std::sync::mpsc::Sender<
        crate::app::DialogAction,
    >,
) -> Option<PathBuf> {
    #[cfg(target_arch = "wasm32")]
    {
        *status_message = "Opening local files is not available in the web build yet".to_owned();
        let _ = (document, canvas, current_path, history);
        None
    }

    #[cfg(not(target_arch = "wasm32"))]
    {
        let _ = (document, canvas, status_message, current_path, history);
        let tx = dialog_tx.clone();
        std::thread::spawn(move || {
            if let Some(path) = FileDialog::new()
                .add_filter("supported", &["txt", "md", "markdown", "docx", "odt"])
                .pick_file()
            {
                let _ = tx.send(crate::app::DialogAction::OpenDocument(path));
            }
        });
        None
    }
}

#[cfg(not(target_arch = "wasm32"))]
pub fn open_document_from_path(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
    history: &mut ChangeHistory,
    path: &Path,
) -> bool {
    match DocumentState::load_from_path(path) {
        Ok(new_document) => {
            let imported_document = matches!(
                path.extension().and_then(|ext| ext.to_str()),
                Some("docx" | "odt")
            );
            history.clear();
            *document = new_document;
            canvas.selection = egui::text_selection::CCursorRange::default();
            canvas.active_style = CharacterStyle::default();
            canvas.active_paragraph_style = ParagraphStyle::default();
            canvas.imported_docx_view = imported_document;
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
            *current_path = Some(path.to_path_buf());
            *status_message = format!(
                "Imported {}",
                path.file_name()
                    .and_then(|name| name.to_str())
                    .unwrap_or("document")
            );
            true
        }
        Err(error) => {
            *status_message = error;
            false
        }
    }
}

pub fn save_document(
    document: &DocumentState,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
    #[cfg(not(target_arch = "wasm32"))] dialog_tx: &std::sync::mpsc::Sender<
        crate::app::DialogAction,
    >,
) -> Option<PathBuf> {
    #[cfg(target_arch = "wasm32")]
    {
        match download_document(document) {
            Ok(filename) => *status_message = format!("Downloaded {filename}"),
            Err(error) => *status_message = error,
        }
        let _ = current_path;
        None
    }

    #[cfg(not(target_arch = "wasm32"))]
    {
        if let Some(path) = current_path.clone() {
            match document.save_to_path(&path) {
                Ok(()) => {
                    *current_path = Some(path.clone());
                    *status_message = format!(
                        "Saved {}",
                        path.file_name()
                            .and_then(|name| name.to_str())
                            .unwrap_or("document")
                    );
                    Some(path)
                }
                Err(error) => {
                    *status_message = error;
                    None
                }
            }
        } else {
            let tx = dialog_tx.clone();
            let title = document.title.clone();
            std::thread::spawn(move || {
                if let Some(path) = pick_save_path_with_file_name(&title) {
                    let _ = tx.send(crate::app::DialogAction::SaveDocument(path));
                }
            });
            None
        }
    }
}

pub fn save_document_as(
    document: &DocumentState,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
    #[cfg(not(target_arch = "wasm32"))] dialog_tx: &std::sync::mpsc::Sender<
        crate::app::DialogAction,
    >,
) -> Option<PathBuf> {
    #[cfg(target_arch = "wasm32")]
    {
        match download_document(document) {
            Ok(filename) => *status_message = format!("Downloaded {filename}"),
            Err(error) => *status_message = error,
        }
        let _ = current_path;
        None
    }

    #[cfg(not(target_arch = "wasm32"))]
    {
        let _ = (status_message, current_path);
        let tx = dialog_tx.clone();
        let title = document.title.clone();
        std::thread::spawn(move || {
            if let Some(path) = pick_save_path_with_file_name(&title) {
                let _ = tx.send(crate::app::DialogAction::SaveDocument(path));
            }
        });
        None
    }
}

pub fn save_document_as_with_name(
    document: &DocumentState,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
    file_name: &str,
    extension: &str,
    #[cfg(not(target_arch = "wasm32"))] dialog_tx: &std::sync::mpsc::Sender<
        crate::app::DialogAction,
    >,
) -> Option<PathBuf> {
    let suggested_name = suggested_save_name(file_name, extension);

    #[cfg(target_arch = "wasm32")]
    {
        match download_document_as(document, &suggested_name, extension) {
            Ok(filename) => *status_message = format!("Downloaded {filename}"),
            Err(error) => *status_message = error,
        }
        let _ = current_path;
        None
    }

    #[cfg(not(target_arch = "wasm32"))]
    {
        let _ = (document, status_message, current_path);
        let tx = dialog_tx.clone();
        std::thread::spawn(move || {
            if let Some(path) = pick_save_path_with_file_name(&suggested_name) {
                let _ = tx.send(crate::app::DialogAction::SaveDocument(path));
            }
        });
        None
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn pick_save_path_with_file_name(file_name: &str) -> Option<PathBuf> {
    FileDialog::new()
        .add_filter("text", &["txt"])
        .add_filter("markdown", &["md", "markdown"])
        .add_filter("web (formatted)", &["html", "htm"])
        .add_filter("Word document", &["docx"])
        .add_filter("OpenDocument text", &["odt"])
        .add_filter("pdf", &["pdf"])
        .set_file_name(file_name)
        .save_file()
}

#[cfg(target_arch = "wasm32")]
fn download_document(document: &DocumentState) -> Result<String, String> {
    let filename = download_document_as(document, &document.title, "html")?;
    Ok(filename)
}

#[cfg(target_arch = "wasm32")]
fn download_document_as(
    document: &DocumentState,
    file_name: &str,
    extension: &str,
) -> Result<String, String> {
    let filename = suggested_save_name(file_name, extension);
    let extension = extension.trim_start_matches('.');
    let bytes = document.export_bytes_for_extension(extension)?;
    download_bytes(&filename, mime_type_for_extension(extension), &bytes)?;
    Ok(filename)
}

fn suggested_save_name(file_name: &str, extension: &str) -> String {
    let extension = extension.trim_start_matches('.').to_ascii_lowercase();
    let fallback = if file_name.trim().is_empty() {
        "document"
    } else {
        file_name.trim()
    };
    let path = std::path::Path::new(fallback);
    let mut name = path
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("document")
        .trim()
        .to_owned();
    if name.is_empty() {
        name = "document".to_owned();
    }
    let has_extension = path
        .extension()
        .and_then(|ext| ext.to_str())
        .is_some_and(|ext| ext.eq_ignore_ascii_case(&extension));
    if !extension.is_empty() && !has_extension {
        if let Some(stem) = path.file_stem().and_then(|stem| stem.to_str()) {
            if !stem.trim().is_empty() {
                name = stem.trim().to_owned();
            }
        }
        name.push('.');
        name.push_str(&extension);
    }
    name
}

#[cfg(target_arch = "wasm32")]
fn mime_type_for_extension(extension: &str) -> &'static str {
    match extension {
        "md" | "markdown" => "text/markdown;charset=utf-8",
        "txt" => "text/plain;charset=utf-8",
        "pdf" => "application/pdf",
        "html" | "htm" => "text/html;charset=utf-8",
        "docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "odt" => "application/vnd.oasis.opendocument.text",
        _ => "application/octet-stream",
    }
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
