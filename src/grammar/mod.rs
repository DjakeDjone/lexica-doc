#[cfg(not(target_arch = "wasm32"))]
pub mod download;
pub mod process;
#[cfg(not(target_arch = "wasm32"))]
pub mod task;

use std::path::PathBuf;

#[cfg(not(target_arch = "wasm32"))]
use anyhow::Result;
#[cfg(not(target_arch = "wasm32"))]
use serde::Deserialize;
use whatlang::Lang;

#[derive(Clone, Debug)]
pub struct GrammarConfig {
    pub lt_jar_path: PathBuf,
    pub port: u16,
    pub language: Language,
}

impl Default for GrammarConfig {
    fn default() -> Self {
        Self {
            lt_jar_path: process::default_lt_jar_path(),
            port: 8081,
            language: Language::Auto,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Language {
    EnUs,
    DeDE,
    Auto,
}

impl Language {
    pub fn to_languagetool_code(self, text: &str) -> &'static str {
        match self {
            Self::EnUs => "en-US",
            Self::DeDE => "de-DE",
            Self::Auto => detect_language(text),
        }
    }
}

#[derive(Clone, Debug)]
pub struct GrammarError {
    pub byte_start: usize,
    pub byte_end: usize,
    pub message: String,
    pub short_message: String,
    pub replacements: Vec<String>,
    pub rule_id: String,
}

#[derive(Clone, Debug)]
pub enum GrammarStatus {
    Idle,
    Checking,
    Done,
    Unavailable(String),
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Clone, Debug)]
pub struct GrammarChecker {
    client: reqwest::Client,
    port: u16,
}

#[cfg(not(target_arch = "wasm32"))]
impl GrammarChecker {
    pub fn new(port: u16) -> Self {
        Self {
            client: reqwest::Client::new(),
            port,
        }
    }

    pub fn client(&self) -> &reqwest::Client {
        &self.client
    }

    pub async fn check(&self, text: &str, language: &str) -> Result<Vec<GrammarError>> {
        if text.trim().is_empty() {
            return Ok(Vec::new());
        }

        let url = format!("http://localhost:{}/v2/check", self.port);
        let send_result = self
            .client
            .post(url)
            .form(&[("text", text), ("language", language)])
            .send()
            .await;

        let response = match send_result {
            Ok(response) => response,
            Err(err) if err.is_connect() => return Ok(Vec::new()),
            Err(err) => return Err(err.into()),
        };

        let payload: LtCheckResponse = response.error_for_status()?.json().await?;
        Ok(payload
            .matches
            .into_iter()
            .map(|m| {
                let byte_start = char_offset_to_byte_index(text, m.offset);
                let byte_end = char_offset_to_byte_index(text, m.offset.saturating_add(m.length))
                    .max(byte_start);
                let replacements = m
                    .replacements
                    .into_iter()
                    .map(|entry| entry.value)
                    .filter(|value| !value.is_empty())
                    .take(5)
                    .collect();
                let short_message = if m.short_message.is_empty() {
                    m.message.clone()
                } else {
                    m.short_message
                };

                GrammarError {
                    byte_start,
                    byte_end,
                    message: m.message,
                    short_message,
                    replacements,
                    rule_id: m.rule.id,
                }
            })
            .collect())
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Deserialize)]
struct LtCheckResponse {
    matches: Vec<LtMatch>,
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Deserialize)]
struct LtMatch {
    message: String,
    #[serde(default, rename = "shortMessage")]
    short_message: String,
    offset: usize,
    length: usize,
    #[serde(default)]
    replacements: Vec<LtReplacement>,
    rule: LtRule,
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Deserialize)]
struct LtReplacement {
    value: String,
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Deserialize)]
struct LtRule {
    id: String,
}

fn detect_language(text: &str) -> &'static str {
    let sample: String = text.chars().take(500).collect();
    let detected = whatlang::detect(&sample).map(|info| info.lang());
    match detected {
        Some(Lang::Deu) => "de-DE",
        _ => "en-US",
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn char_offset_to_byte_index(text: &str, char_offset: usize) -> usize {
    text.char_indices()
        .nth(char_offset)
        .map(|(byte_idx, _)| byte_idx)
        .unwrap_or(text.len())
}
