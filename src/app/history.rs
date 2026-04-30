use std::collections::VecDeque;

use crate::document::DocumentState;

const HISTORY_LIMIT: usize = 200;
const CHECKPOINT_COALESCE_SECONDS: f64 = 0.75;

pub struct ChangeHistory {
    undo_stack: VecDeque<DocumentState>,
    redo_stack: VecDeque<DocumentState>,
    last_checkpoint_time: f64,
}

impl ChangeHistory {
    pub fn new() -> Self {
        Self {
            undo_stack: VecDeque::new(),
            redo_stack: VecDeque::new(),
            last_checkpoint_time: f64::NEG_INFINITY,
        }
    }

    fn push_snapshot(stack: &mut VecDeque<DocumentState>, document: &DocumentState) {
        stack.push_back(document.clone());
        if stack.len() > HISTORY_LIMIT {
            stack.pop_front();
        }
    }

    fn remember_undo_snapshot(&mut self, document: &DocumentState) {
        Self::push_snapshot(&mut self.undo_stack, document);
        self.redo_stack.clear();
    }

    /// Always checkpoint — use before discrete actions (button clicks).
    pub fn checkpoint(&mut self, document: &DocumentState, now: f64) {
        self.remember_undo_snapshot(document);
        self.last_checkpoint_time = now;
    }

    /// Checkpoint only if enough time has elapsed — use before continuous controls (drag values).
    pub fn checkpoint_coalesced(&mut self, document: &DocumentState, now: f64) {
        if now - self.last_checkpoint_time > CHECKPOINT_COALESCE_SECONDS {
            self.remember_undo_snapshot(document);
            self.last_checkpoint_time = now;
        }
    }

    pub fn undo(&mut self, document: &mut DocumentState) -> bool {
        if let Some(prev) = self.undo_stack.pop_back() {
            Self::push_snapshot(&mut self.redo_stack, document);
            *document = prev;
            true
        } else {
            false
        }
    }

    pub fn redo(&mut self, document: &mut DocumentState) -> bool {
        if let Some(next) = self.redo_stack.pop_back() {
            Self::push_snapshot(&mut self.undo_stack, document);
            *document = next;
            true
        } else {
            false
        }
    }

    pub fn can_undo(&self) -> bool {
        !self.undo_stack.is_empty()
    }

    pub fn can_redo(&self) -> bool {
        !self.redo_stack.is_empty()
    }

    pub fn clear(&mut self) {
        self.undo_stack.clear();
        self.redo_stack.clear();
        self.last_checkpoint_time = f64::NEG_INFINITY;
    }
}

impl Default for ChangeHistory {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::{ChangeHistory, HISTORY_LIMIT};
    use crate::document::DocumentState;

    fn document_with_title(title: impl Into<String>) -> DocumentState {
        let mut document = DocumentState::bootstrap();
        document.title = title.into();
        document
    }

    #[test]
    fn caps_undo_history_without_losing_recent_snapshots() {
        let mut history = ChangeHistory::new();
        let mut document = document_with_title("current");

        for index in 0..(HISTORY_LIMIT + 5) {
            document.title = format!("snapshot {index}");
            history.checkpoint(&document, index as f64);
        }

        let mut active = document_with_title("active");
        assert!(history.undo(&mut active));
        assert_eq!(active.title, format!("snapshot {}", HISTORY_LIMIT + 4));

        for _ in 1..HISTORY_LIMIT {
            assert!(history.undo(&mut active));
        }

        assert_eq!(active.title, "snapshot 5");
        assert!(!history.undo(&mut active));
    }

    #[test]
    fn coalesced_checkpoints_respect_time_window() {
        let mut history = ChangeHistory::new();
        let document = DocumentState::bootstrap();

        history.checkpoint_coalesced(&document, 1.0);
        history.checkpoint_coalesced(&document, 1.2);

        let mut active = document_with_title("active");
        assert!(history.undo(&mut active));
        assert!(!history.undo(&mut active));
    }
}
