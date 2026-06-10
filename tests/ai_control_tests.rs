use wors::app::{CanvasState, ChangeHistory};
use wors::document::DocumentState;
use wors::http_server::ApiRequest;
use wors::api::{EditorCommand, UiAction};
use tokio::sync::oneshot;
use std::path::PathBuf;

#[tokio::test]
async fn test_ai_control_handle_requests() {
    let mut document = DocumentState::bootstrap();
    let mut canvas = CanvasState::default();
    let mut history = ChangeHistory::default();
    let mut status_message = String::new();
    let current_path = Some(PathBuf::from("test.md"));

    // Set up some initial text
    document.delete_range(0..document.plain_text().chars().count());
    document.insert_text(0, "Hello World", Default::default());
    canvas.selection.primary.index = 11;
    canvas.selection.secondary.index = 11;

    // Test 1: Insert Text
    let (tx, rx) = oneshot::channel();
    wors::http_server::handle_api_request(
        ApiRequest::ApplyCommand(EditorCommand::InsertText { text: "!".to_string() }, tx),
        &mut document,
        &mut canvas,
        &mut history,
        &mut status_message,
        &current_path,
    );
    assert_eq!(rx.await.unwrap(), Ok(()));
    assert_eq!(document.plain_text(), "Hello World!");

    // Test 2: Replace Valid Range
    let (tx, rx) = oneshot::channel();
    wors::http_server::handle_api_request(
        ApiRequest::ApplyCommand(EditorCommand::ReplaceRange { start: 0, end: 5, text: "Hi".to_string() }, tx),
        &mut document,
        &mut canvas,
        &mut history,
        &mut status_message,
        &current_path,
    );
    assert_eq!(rx.await.unwrap(), Ok(()));
    assert_eq!(document.plain_text(), "Hi World!");

    // Test 3: Replace Invalid Range
    let (tx, rx) = oneshot::channel();
    wors::http_server::handle_api_request(
        ApiRequest::ApplyCommand(EditorCommand::ReplaceRange { start: 100, end: 105, text: "Invalid".to_string() }, tx),
        &mut document,
        &mut canvas,
        &mut history,
        &mut status_message,
        &current_path,
    );
    assert!(rx.await.unwrap().is_err());
    assert_eq!(document.plain_text(), "Hi World!");

    // Test 4: Set Selection
    let (tx, rx) = oneshot::channel();
    wors::http_server::handle_api_request(
        ApiRequest::ApplyCommand(EditorCommand::SetSelection { anchor: 0, focus: 2 }, tx),
        &mut document,
        &mut canvas,
        &mut history,
        &mut status_message,
        &current_path,
    );
    assert_eq!(rx.await.unwrap(), Ok(()));
    assert_eq!(canvas.selection.secondary.index, 0);
    assert_eq!(canvas.selection.primary.index, 2);

    // Test 5: Toolbar Bold Action
    let (tx, rx) = oneshot::channel();
    wors::http_server::handle_api_request(
        ApiRequest::ApplyCommand(EditorCommand::InvokeUi { id: "toolbar.bold".to_string(), action: UiAction::Click }, tx),
        &mut document,
        &mut canvas,
        &mut history,
        &mut status_message,
        &current_path,
    );
    assert_eq!(rx.await.unwrap(), Ok(()));
    // Verify the selection was made bold
    let style = document.selection_style_at(0..2);
    assert!(style.bold);

    // Test 6: Get UI Tree
    let (tx, rx) = oneshot::channel();
    wors::http_server::handle_api_request(
        ApiRequest::GetUiTree(tx),
        &mut document,
        &mut canvas,
        &mut history,
        &mut status_message,
        &current_path,
    );
    let tree = rx.await.unwrap();
    assert_eq!(tree.id, "document.body");
    assert!(tree.children.iter().any(|c| c.id == "toolbar.bold"));

    // Test 7: Undo
    let (tx, rx) = oneshot::channel();
    wors::http_server::handle_api_request(
        ApiRequest::ApplyCommand(EditorCommand::Undo, tx),
        &mut document,
        &mut canvas,
        &mut history,
        &mut status_message,
        &current_path,
    );
    assert_eq!(rx.await.unwrap(), Ok(()));
    // After undoing the bold action
    let style = document.selection_style_at(0..2);
    assert!(!style.bold);

    // Test 8: Get Document Text
    let (tx, rx) = oneshot::channel();
    wors::http_server::handle_api_request(
        ApiRequest::GetDocumentText(tx),
        &mut document,
        &mut canvas,
        &mut history,
        &mut status_message,
        &current_path,
    );
    let text = rx.await.unwrap();
    assert_eq!(text, "Hi World!");
}
