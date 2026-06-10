use serde::{Deserialize, Serialize};
use std::time::Duration;
use tokio::sync::mpsc;
use tokio::time::sleep;

const DEBOUNCE_DELAY_MS: u64 = 750;

#[derive(Clone, Debug)]
pub struct AiRequest {
    pub text_before_cursor: String,
    pub text_after_cursor: String,
    pub endpoint: String,
    pub model: String,
}

#[derive(Debug)]
pub enum AiTaskResult {
    Started,
    Completed(String),
    Unavailable(String),
}

#[derive(Serialize)]
struct OllamaGenerateRequest {
    model: String,
    prompt: String,
    stream: bool,
    options: OllamaOptions,
}

#[derive(Serialize)]
struct OllamaOptions {
    num_predict: usize,
}

#[derive(Deserialize)]
struct OllamaGenerateResponse {
    response: String,
}

pub async fn run_ai_task(
    mut rx: mpsc::Receiver<AiRequest>,
    results_tx: mpsc::Sender<AiTaskResult>,
) {
    let client = reqwest::Client::builder()
        .timeout(Duration::from_secs(10))
        .build()
        .unwrap_or_default();

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

        if pending.text_before_cursor.trim().is_empty() {
            let _ = results_tx
                .send(AiTaskResult::Completed(String::new()))
                .await;
            continue;
        }

        let prompt = format!(
            "Please complete the following text. Return ONLY the completion without any leading or trailing commentary or formatting.\n\n{}",
            pending.text_before_cursor
        );

        let body = OllamaGenerateRequest {
            model: pending.model.clone(),
            prompt,
            stream: false,
            options: OllamaOptions { num_predict: 50 },
        };

        let endpoint = if pending.endpoint.ends_with('/') {
            format!("{}api/generate", pending.endpoint)
        } else {
            format!("{}/api/generate", pending.endpoint)
        };

        let _ = results_tx.send(AiTaskResult::Started).await;

        match client.post(&endpoint).json(&body).send().await {
            Ok(response) => {
                if response.status().is_success() {
                    if let Ok(parsed) = response.json::<OllamaGenerateResponse>().await {
                        let _ = results_tx
                            .send(AiTaskResult::Completed(
                                parsed.response.trim_end().to_owned(),
                            ))
                            .await;
                    } else {
                        let _ = results_tx
                            .send(AiTaskResult::Unavailable(
                                "Failed to parse Ollama response".to_owned(),
                            ))
                            .await;
                    }
                } else {
                    let _ = results_tx
                        .send(AiTaskResult::Unavailable(format!(
                            "Ollama API returned error: {}",
                            response.status()
                        )))
                        .await;
                }
            }
            Err(err) => {
                let _ = results_tx
                    .send(AiTaskResult::Unavailable(format!(
                        "Request to Ollama failed: {err}"
                    )))
                    .await;
            }
        }
    }
}
