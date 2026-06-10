use axum::{
    extract::State,
    routing::{get, post},
    Json, Router,
};
use serde::{Deserialize, Serialize};
use tokio::net::TcpListener;
use tokio::sync::{mpsc, oneshot};

use crate::api::{EditorCommand, EditorSnapshot, UiNode};

pub enum ApiRequest {
    GetState(oneshot::Sender<EditorSnapshot>),
    GetDocumentText(oneshot::Sender<String>),
    ApplyCommand(EditorCommand, oneshot::Sender<Result<(), String>>),
    GetUiTree(oneshot::Sender<UiNode>),
}

#[derive(Clone)]
struct AppState {
    tx: mpsc::Sender<ApiRequest>,
}

#[derive(Serialize)]
struct ErrorResponse {
    error: String,
}

pub async fn start_server(tx: mpsc::Sender<ApiRequest>) -> std::io::Result<()> {
    let state = AppState { tx };

    let app = Router::new()
        .route("/state", get(handle_get_state))
        .route("/document/text", get(handle_get_document_text))
        .route("/command", post(handle_command))
        .route("/ui/tree", get(handle_get_ui_tree))
        .route("/ui/invoke", post(handle_ui_invoke))
        .with_state(state);

    let listener = TcpListener::bind("127.0.0.1:0").await?;
    let addr = listener.local_addr()?;
    
    println!("wors control API listening on {}", addr);
    
    // Write port to ~/.wors/mcp-port
    if let Some(mut path) = dirs::home_dir() {
        path.push(".wors");
        let _ = std::fs::create_dir_all(&path);
        path.push("mcp-port");
        let _ = std::fs::write(path, addr.port().to_string());
    }

    axum::serve(listener, app).await
}

async fn handle_get_state(State(state): State<AppState>) -> Result<Json<EditorSnapshot>, Json<ErrorResponse>> {
    let (resp_tx, resp_rx) = oneshot::channel();
    let _ = state.tx.send(ApiRequest::GetState(resp_tx)).await;
    match resp_rx.await {
        Ok(snapshot) => Ok(Json(snapshot)),
        Err(_) => Err(Json(ErrorResponse { error: "Failed to get state".into() })),
    }
}

async fn handle_get_document_text(State(state): State<AppState>) -> Result<String, Json<ErrorResponse>> {
    let (resp_tx, resp_rx) = oneshot::channel();
    let _ = state.tx.send(ApiRequest::GetDocumentText(resp_tx)).await;
    match resp_rx.await {
        Ok(text) => Ok(text),
        Err(_) => Err(Json(ErrorResponse { error: "Failed to get document text".into() })),
    }
}

async fn handle_command(
    State(state): State<AppState>,
    Json(payload): Json<EditorCommand>,
) -> Result<Json<()>, Json<ErrorResponse>> {
    let (resp_tx, resp_rx) = oneshot::channel();
    let _ = state.tx.send(ApiRequest::ApplyCommand(payload, resp_tx)).await;
    match resp_rx.await {
        Ok(Ok(())) => Ok(Json(())),
        Ok(Err(e)) => Err(Json(ErrorResponse { error: e })),
        Err(_) => Err(Json(ErrorResponse { error: "Failed to apply command".into() })),
    }
}

async fn handle_get_ui_tree(State(state): State<AppState>) -> Result<Json<UiNode>, Json<ErrorResponse>> {
    let (resp_tx, resp_rx) = oneshot::channel();
    let _ = state.tx.send(ApiRequest::GetUiTree(resp_tx)).await;
    match resp_rx.await {
        Ok(tree) => Ok(Json(tree)),
        Err(_) => Err(Json(ErrorResponse { error: "Failed to get UI tree".into() })),
    }
}

#[derive(Deserialize)]
struct UiInvokePayload {
    id: String,
    action: crate::api::UiAction,
}

async fn handle_ui_invoke(
    State(state): State<AppState>,
    Json(payload): Json<UiInvokePayload>,
) -> Result<Json<()>, Json<ErrorResponse>> {
    let command = EditorCommand::InvokeUi {
        id: payload.id,
        action: payload.action,
    };
    let (resp_tx, resp_rx) = oneshot::channel();
    let _ = state.tx.send(ApiRequest::ApplyCommand(command, resp_tx)).await;
    match resp_rx.await {
        Ok(Ok(())) => Ok(Json(())),
        Ok(Err(e)) => Err(Json(ErrorResponse { error: e })),
        Err(_) => Err(Json(ErrorResponse { error: "Failed to invoke UI".into() })),
    }
}

#[cfg(not(target_arch = "wasm32"))]
pub fn handle_api_request(
    request: crate::http_server::ApiRequest,
    document: &mut crate::document::DocumentState,
    canvas: &mut crate::app::CanvasState,
    history: &mut crate::app::ChangeHistory,
    status_message: &mut String,
    current_path: &Option<std::path::PathBuf>,
) {
    use crate::http_server::ApiRequest;
    use crate::api::{EditorSnapshot, UiNode};
    match request {
        ApiRequest::GetState(tx) => {
            let snapshot = EditorSnapshot {
                title: current_path.as_ref().and_then(|p| p.file_name()).and_then(|n| n.to_str()).unwrap_or("document").to_owned(),
                dirty: history.can_undo(), // Approximation for now
                document_text: document.plain_text(),
                selection_anchor: canvas.selection.secondary.index,
                selection_focus: canvas.selection.primary.index,
                cursor_position: canvas.selection.primary.index,
                can_undo: history.can_undo(),
                can_redo: history.can_redo(),
                active_ui: None,
            };
            let _ = tx.send(snapshot);
        }
        ApiRequest::GetDocumentText(tx) => {
            let _ = tx.send(document.plain_text());
        }
        ApiRequest::ApplyCommand(command, tx) => {
            let mut success = true;
            let mut error_msg = String::new();
            use crate::api::EditorCommand;
            match command {
                EditorCommand::InsertText { text } => {
                    let style = canvas.active_style;
                    history.checkpoint(document, f64::NAN);
                    document.insert_text(canvas.selection.primary.index, &text, style);
                    canvas.selection.primary.index += text.chars().count();
                    canvas.selection.secondary.index = canvas.selection.primary.index;
                }
                EditorCommand::ReplaceRange { start, end, text } => {
                    if start <= end && end <= document.plain_text().chars().count() {
                        let style = canvas.active_style;
                        history.checkpoint(document, f64::NAN);
                        document.delete_range(start..end);
                        document.insert_text(start, &text, style);
                        canvas.selection.primary.index = start + text.chars().count();
                        canvas.selection.secondary.index = canvas.selection.primary.index;
                    } else {
                        success = false;
                        error_msg = "Invalid range".to_owned();
                    }
                }
                EditorCommand::SetSelection { anchor, focus } => {
                    let len = document.plain_text().chars().count();
                    if anchor <= len && focus <= len {
                        canvas.selection.secondary.index = anchor;
                        canvas.selection.primary.index = focus;
                    } else {
                        success = false;
                        error_msg = "Invalid selection".to_owned();
                    }
                }
                EditorCommand::ToggleBold => {
                    crate::app::actions::toggle_bold(document, canvas, history);
                }
                EditorCommand::ToggleItalic => {
                    crate::app::actions::toggle_italic(document, canvas, history);
                }
                EditorCommand::Undo => {
                    if history.undo(document) {
                        canvas.image_textures.clear();
                        *status_message = "Undo".to_owned();
                    }
                }
                EditorCommand::Redo => {
                    if history.redo(document) {
                        canvas.image_textures.clear();
                        *status_message = "Redo".to_owned();
                    }
                }
                EditorCommand::Save => {
                }
                EditorCommand::ExportMarkdown { path: _path } => {
                }
                EditorCommand::ExportPdf { path } => {
                    match document.to_pdf_bytes() {
                        Ok(bytes) => {
                            if let Err(e) = std::fs::write(&path, bytes) {
                                success = false;
                                error_msg = format!("Failed to write PDF: {}", e);
                            }
                        }
                        Err(e) => {
                            success = false;
                            error_msg = format!("Failed to export PDF: {}", e);
                        }
                    }
                }
                EditorCommand::InvokeUi { id, action: _action } => {
                    if id == "toolbar.bold" {
                        crate::app::actions::toggle_bold(document, canvas, history);
                    } else if id == "toolbar.italic" {
                        crate::app::actions::toggle_italic(document, canvas, history);
                    }
                }
            }
            if success {
                let _ = tx.send(Ok(()));
            } else {
                let _ = tx.send(Err(error_msg));
            }
        }
        ApiRequest::GetUiTree(tx) => {
            let tree = UiNode {
                id: "document.body".to_owned(),
                role: "document".to_owned(),
                label: "Document".to_owned(),
                enabled: true,
                visible: true,
                bounds: None,
                children: vec![
                    UiNode {
                        id: "toolbar.bold".to_owned(),
                        role: "button".to_owned(),
                        label: "Bold".to_owned(),
                        enabled: true,
                        visible: true,
                        bounds: None,
                        children: vec![],
                    },
                    UiNode {
                        id: "toolbar.italic".to_owned(),
                        role: "button".to_owned(),
                        label: "Italic".to_owned(),
                        enabled: true,
                        visible: true,
                        bounds: None,
                        children: vec![],
                    },
                    UiNode {
                        id: "toolbar.save".to_owned(),
                        role: "button".to_owned(),
                        label: "Save".to_owned(),
                        enabled: true,
                        visible: true,
                        bounds: None,
                        children: vec![],
                    },
                ],
            };
            let _ = tx.send(tree);
        }
    }
}
