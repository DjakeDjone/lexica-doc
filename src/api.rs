use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum EditorCommand {
    InsertText {
        text: String,
    },
    ReplaceRange {
        start: usize,
        end: usize,
        text: String,
    },
    SetSelection {
        anchor: usize,
        focus: usize,
    },
    ToggleBold,
    ToggleItalic,
    Undo,
    Redo,
    Save,
    ExportMarkdown {
        path: String,
    },
    ExportPdf {
        path: String,
    },
    InvokeUi {
        id: String,
        action: UiAction,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum UiAction {
    Click,
    Focus,
    Toggle,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct UiNode {
    pub id: String,
    pub role: String,
    pub label: String,
    pub enabled: bool,
    pub visible: bool,
    pub bounds: Option<UiBounds>,
    pub children: Vec<UiNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct UiBounds {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EditorSnapshot {
    pub title: String,
    pub dirty: bool,
    pub document_text: String,
    pub selection_anchor: usize,
    pub selection_focus: usize,
    pub cursor_position: usize,
    pub can_undo: bool,
    pub can_redo: bool,
    pub active_ui: Option<String>,
}
