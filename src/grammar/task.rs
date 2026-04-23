use std::time::Duration;

use tokio::sync::mpsc;
use tokio::time::sleep;

use super::{process, GrammarChecker, GrammarError};

const DEBOUNCE_DELAY_MS: u64 = 500;

#[derive(Clone, Debug)]
pub struct GrammarRequest {
    pub text: String,
    pub language: String,
}

#[derive(Debug)]
pub enum GrammarTaskResult {
    Completed(Vec<GrammarError>),
    Unavailable(String),
}

pub async fn run_grammar_task(
    mut rx: mpsc::Receiver<GrammarRequest>,
    results_tx: mpsc::Sender<GrammarTaskResult>,
    checker: GrammarChecker,
    port: u16,
) {
    if !process::wait_until_ready(checker.client(), port).await {
        let _ = results_tx
            .send(GrammarTaskResult::Unavailable(format!(
                "LanguageTool server did not become ready on port {port}"
            )))
            .await;
        return;
    }

    while let Some(mut pending) = rx.recv().await {
        loop {
            sleep(Duration::from_millis(DEBOUNCE_DELAY_MS)).await;

            let mut saw_newer = false;
            while let Ok(newer) = rx.try_recv() {
                pending = newer;
                saw_newer = true;
            }

            if !saw_newer {
                break;
            }
        }

        match checker.check(&pending.text, &pending.language).await {
            Ok(errors) => {
                let _ = results_tx.send(GrammarTaskResult::Completed(errors)).await;
            }
            Err(err) => {
                let _ = results_tx
                    .send(GrammarTaskResult::Unavailable(format!(
                        "Grammar check failed: {err}"
                    )))
                    .await;
            }
        }
    }
}
